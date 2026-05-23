from __future__ import annotations

import argparse
import json
import logging
import os
import re
import sys
import traceback
from pathlib import Path
from time import perf_counter


ROOT = Path(__file__).resolve().parents[1]
NOTEBOOK_CANDIDATES = [
    ROOT / "Notebooks" / "Transcription Pipeline.ipynb",
    ROOT / "Nigeria" / "Notebooks" / "Transcription Pipeline.ipynb",
    ROOT / "Nigeria" / "Notebooks" / "Data Acquisition Pipeline.ipynb",
]

for import_path in (ROOT, ROOT / "scripts"):
    import_path_str = str(import_path)
    if import_path_str not in sys.path:
        sys.path.insert(0, import_path_str)

SETUP_CELL_MARKERS = [
    "from __future__ import annotations",
    'YOUTUBE_URL = ""',
    "VIDEO_JOBS = {",
]
PREFLIGHT_CELL_MARKERS = [
    'require_binary("ffmpeg")',
    "def ensure_runtime_is_ready() -> None:",
    "class VideoContextPayload(TypedDict, total=False):",
]
PIPELINE_EXTRA_CELL_MARKERS = [
    "def convert_to_wav_16k_mono(input_path: str, output_path: str) -> str:",
    "def normalize_word_entries(raw_words: list[dict], *, source: str) -> list[dict]:",
    "def derive_diarization_constraints(metadata: dict) -> dict:",
    "def assign_speakers_to_words(words: list[dict], speaker_segments: list[dict], tolerance: float = 0.75) -> list[dict]:",
    "# transcriptsUtils guards",
    "def is_audience_like_turn(text: str) -> bool:",
    "def resolve_transcript_output_path(output_dir: Path, metadata: dict, transcript_filename: str | None) -> Path:",
    "preview_text = Path(saved_path).read_text(encoding='utf-8')",
]

log = logging.getLogger("run_all_transcriptions")
_runtime_plan: dict | None = None


def resolve_notebook_path() -> Path:
    configured = os.getenv("NLC_TRANSCRIPTION_NOTEBOOK", "").strip()
    if configured:
        path = Path(configured).expanduser().resolve()
        if not path.exists():
            raise FileNotFoundError(
                f"NLC_TRANSCRIPTION_NOTEBOOK points to a missing file: {path}"
            )
        return path
    for candidate in NOTEBOOK_CANDIDATES:
        if candidate.exists():
            return candidate
    checked = "\n - ".join(str(c) for c in NOTEBOOK_CANDIDATES)
    raise FileNotFoundError(
        "No supported transcription notebook found. Checked:\n"
        f" - {checked}\n"
        "Set NLC_TRANSCRIPTION_NOTEBOOK to an explicit path."
    )


NOTEBOOK_PATH = resolve_notebook_path()


def load_notebook_code_cells() -> dict[int, str]:
    notebook = json.loads(NOTEBOOK_PATH.read_text(encoding="utf-8"))
    code_cells: dict[int, str] = {}
    for index, cell in enumerate(notebook["cells"]):
        if cell.get("cell_type") != "code":
            continue
        code_cells[index] = "".join(cell.get("source", []))
    return code_cells


def resolve_cell_index(code_cells: dict[int, str], marker: str) -> int:
    matches = [index for index, source in code_cells.items() if marker in source]
    if not matches:
        raise RuntimeError(f"Notebook cell marker not found: {marker}")
    if len(matches) > 1:
        raise RuntimeError(f"Notebook cell marker is ambiguous: {marker} -> {matches}")
    return matches[0]


def _slugify(value: str) -> str:
    value = re.sub(r"[^a-z0-9]+", "_", value.lower()).strip("_")
    return value or "episode"


def _is_legacy_layout(code_cells: dict[int, str]) -> bool:
    blob = "\n".join(code_cells.values())
    return "VIDEO_JOBS" in blob and "apply_video_job" in blob


def _resolve_runtime_plan() -> dict:
    global _runtime_plan
    if _runtime_plan is not None:
        return _runtime_plan

    code_cells = load_notebook_code_cells()
    if _is_legacy_layout(code_cells):
        setup_indexes = [resolve_cell_index(code_cells, marker) for marker in SETUP_CELL_MARKERS]
        preflight_indexes = [resolve_cell_index(code_cells, marker) for marker in PREFLIGHT_CELL_MARKERS]
        pipeline_indexes = preflight_indexes + [
            resolve_cell_index(code_cells, marker) for marker in PIPELINE_EXTRA_CELL_MARKERS
        ]
        _runtime_plan = {
            "mode": "legacy",
            "setup_indexes": setup_indexes,
            "preflight_indexes": preflight_indexes,
            "pipeline_indexes": pipeline_indexes,
        }
        return _runtime_plan

    setup_indexes: list[int] = []
    for marker in ("from __future__ import annotations", "EPISODES = ["):
        matches = [i for i, src in code_cells.items() if marker in src]
        if matches:
            setup_indexes.append(matches[0])
    if not setup_indexes and code_cells:
        setup_indexes = [min(code_cells)]

    _runtime_plan = {
        "mode": "episodes",
        "setup_indexes": sorted(set(setup_indexes)),
        "preflight_indexes": [],
        "pipeline_indexes": [],
    }
    return _runtime_plan


def _derive_output_subdir(runtime: dict) -> str:
    trans_dir = runtime.get("TRANS_DIR")
    if isinstance(trans_dir, Path):
        try:
            rel = trans_dir.resolve().relative_to(ROOT.resolve())
            return str(rel.parent)
        except ValueError:
            pass
    return "Nigeria/Content Analysis/Content - Raw/Banky Wellington"


def _ensure_video_jobs(runtime: dict) -> None:
    if isinstance(runtime.get("VIDEO_JOBS"), dict) and runtime["VIDEO_JOBS"]:
        return

    episodes = runtime.get("EPISODES")
    if not isinstance(episodes, list) or not episodes:
        raise RuntimeError(
            f"Notebook {NOTEBOOK_PATH} did not define VIDEO_JOBS or EPISODES. "
            "Cannot build job runtime."
        )

    audio_dir = runtime.get("AUDIO_DIR")
    output_subdir = _derive_output_subdir(runtime)
    jobs: dict[str, dict] = {}

    for idx, episode in enumerate(episodes, start=1):
        if not isinstance(episode, dict):
            continue
        audio_name = str(episode.get("audio", "")).strip()
        stem = Path(audio_name).stem or f"episode_{idx:02d}"
        job_name = _slugify(stem)
        while job_name in jobs:
            job_name = f"{job_name}_{idx}"

        if isinstance(audio_dir, Path) and audio_name:
            local_audio_path = str((audio_dir / audio_name).resolve())
        else:
            local_audio_path = audio_name

        yt_id = str(episode.get("yt_id", "")).strip()
        youtube_url = (
            f"https://www.youtube.com/watch?v={yt_id}" if yt_id else str(episode.get("youtube_url", "")).strip()
        )

        jobs[job_name] = {
            "local_audio_path": local_audio_path,
            "output_subdir": output_subdir,
            "youtube_url": youtube_url,
            "panel": episode.get("panel", []),
            "_source": "episodes",
        }

    if not jobs:
        raise RuntimeError("Could not synthesize VIDEO_JOBS from EPISODES.")

    runtime["VIDEO_JOBS"] = jobs


def build_runtime() -> dict:
    runtime = {"__name__": "__main__"}
    code_cells = load_notebook_code_cells()
    plan = _resolve_runtime_plan()
    for index in plan["setup_indexes"]:
        exec(compile(code_cells[index], f"{NOTEBOOK_PATH.name}::cell_{index}", "exec"), runtime)
    _ensure_video_jobs(runtime)
    return runtime


def get_job_names(runtime: dict) -> list[str]:
    return list(runtime["VIDEO_JOBS"].keys())


def run_cells(runtime: dict, indexes: list[int]) -> None:
    code_cells = load_notebook_code_cells()
    for index in indexes:
        exec(compile(code_cells[index], f"{NOTEBOOK_PATH.name}::cell_{index}", "exec"), runtime)


def run_single_job(job_name: str, *, preflight_only: bool = False) -> tuple[bool, float, str | None]:
    started = perf_counter()
    plan = _resolve_runtime_plan()
    if plan["mode"] != "legacy":
        elapsed = perf_counter() - started
        return (
            False,
            elapsed,
            "Detected EPISODES-style notebook layout. "
            "Use `python Nigeria/scripts/run_mentality_transcription.py` to execute transcription jobs.",
        )

    runtime = build_runtime()
    runtime["apply_video_job"](job_name)
    try:
        run_cells(runtime, plan["preflight_indexes"] if preflight_only else plan["pipeline_indexes"])
    except Exception as exc:
        elapsed = perf_counter() - started
        return False, elapsed, "".join(traceback.format_exception(exc))
    elapsed = perf_counter() - started
    return True, elapsed, None


def run_preflight(job_names: list[str]) -> bool:
    log.info("Running preflight across %s jobs...", len(job_names))
    failures: list[tuple[str, str]] = []
    for index, job_name in enumerate(job_names, start=1):
        log.info("Preflight %s/%s: %s", index, len(job_names), job_name)
        ok, elapsed, error = run_single_job(job_name, preflight_only=True)
        if ok:
            log.info("Preflight passed for %s in %.1fs", job_name, elapsed)
        else:
            log.error("Preflight failed for %s in %.1fs", job_name, elapsed)
            failures.append((job_name, error or "unknown error"))
    if failures:
        log.error("Preflight failed for %s job(s).", len(failures))
        for job_name, error in failures:
            log.error("FAILED PRECHECK %s\n%s", job_name, error)
        return False
    log.info("All job preflights passed.")
    return True


def run_batch(job_names: list[str]) -> int:
    failures: list[tuple[str, str]] = []
    for index, job_name in enumerate(job_names, start=1):
        log.info("Running job %s/%s: %s", index, len(job_names), job_name)
        ok, elapsed, error = run_single_job(job_name, preflight_only=False)
        if ok:
            log.info("Completed %s in %.1f minutes", job_name, elapsed / 60.0)
            continue
        log.error("Job failed: %s after %.1f minutes", job_name, elapsed / 60.0)
        if error:
            log.error("Traceback for %s\n%s", job_name, error)
        failures.append((job_name, error or "unknown error"))
    if failures:
        log.error("Batch finished with %s failure(s).", len(failures))
        for job_name, error in failures:
            log.error("FAILED JOB %s\n%s", job_name, error)
        return 1
    log.info("Batch completed successfully for %s job(s).", len(job_names))
    return 0


def configure_logging(log_path: Path | None) -> None:
    handlers: list[logging.Handler] = [logging.StreamHandler(sys.stdout)]
    if log_path is not None:
        log_path.parent.mkdir(parents=True, exist_ok=True)
        handlers.append(logging.FileHandler(log_path, encoding="utf-8"))
    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s %(levelname)s %(message)s",
        handlers=handlers,
        force=True,
    )


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Run all transcription jobs from the notebook pipeline.")
    parser.add_argument("--country", choices=["all", "nigeria", "kenya"], default="all")
    parser.add_argument("--skip-preflight", action="store_true")
    parser.add_argument("--jobs", nargs="*", default=None, help="Explicit notebook job names to run.")
    parser.add_argument("--log-file", default=str(ROOT / "temp" / "logs" / "transcription_batch.log"))
    return parser.parse_args()


def select_jobs(runtime: dict, country: str, explicit_jobs: list[str] | None) -> list[str]:
    if explicit_jobs:
        unknown = [job for job in explicit_jobs if job not in runtime["VIDEO_JOBS"]]
        if unknown:
            raise SystemExit(f"Unknown job(s): {', '.join(unknown)}")
        return explicit_jobs

    all_jobs = list(runtime["VIDEO_JOBS"].items())
    if country == "all":
        return [name for name, _ in all_jobs]
    if country == "nigeria":
        return [name for name, job in all_jobs if str(job.get("output_subdir", "")).startswith("Nigeria")]
    return [name for name, job in all_jobs if str(job.get("output_subdir", "")).startswith("Kenya")]


def main() -> int:
    args = parse_args()
    configure_logging(Path(args.log_file) if args.log_file else None)

    try:
        plan = _resolve_runtime_plan()
        runtime = build_runtime()
    except Exception as exc:
        log.error("Failed to initialize transcription runtime from %s", NOTEBOOK_PATH)
        log.error("%s", exc)
        return 1

    if plan["mode"] != "legacy":
        log.error(
            "Notebook layout at %s is EPISODES-style (not legacy VIDEO_JOBS/apply_video_job). "
            "Use `python Nigeria/scripts/run_mentality_transcription.py` to execute full transcription jobs. "
            "This module still provides build_runtime/select_jobs compatibility for downstream scripts.",
            NOTEBOOK_PATH,
        )
        return 1

    job_names = select_jobs(runtime, args.country, args.jobs)
    log.info("Selected %s jobs.", len(job_names))
    for job_name in job_names:
        log.info(" - %s", job_name)
    if not args.skip_preflight:
        if not run_preflight(job_names):
            return 1
    return run_batch(job_names)


if __name__ == "__main__":
    raise SystemExit(main())
