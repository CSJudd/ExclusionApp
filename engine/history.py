from pathlib import Path
from datetime import datetime
import json

BASE_DATA_DIR = Path.home() / "ExclusionAppData"
RUNS_DIR = BASE_DATA_DIR / "runs"

RUNS_DIR.mkdir(parents=True, exist_ok=True)


def create_run_directory(client_name, month):
    run_dir = RUNS_DIR / client_name / month
    run_dir.mkdir(parents=True, exist_ok=True)
    return run_dir


def write_metadata(run_dir, metadata):
    metadata_path = run_dir / "metadata.json"

    # Add timestamp if not already provided
    if "timestamp" not in metadata:
        metadata["timestamp"] = datetime.utcnow().isoformat()

    with open(metadata_path, "w", encoding="utf-8") as f:
        json.dump(metadata, f, indent=4)

    return metadata_path


def write_run_log(run_dir, message):
    log_path = run_dir / "run_log.txt"

    timestamp = datetime.utcnow().isoformat()

    with open(log_path, "a", encoding="utf-8") as f:
        f.write(f"[{timestamp}] {message}\n")

    return log_path