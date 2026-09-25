"""Regression check for warnings preceding prepare's manifest path."""

import contextlib
import io
import subprocess
import sys
import tempfile
from pathlib import Path
from types import SimpleNamespace
from unittest.mock import patch

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "cli" / "editppt" / "runtime"))

import main as runtime_main


def prepare_args():
    return SimpleNamespace(
        out_root=None,
        job_dir=None,
        dpi=None,
        max_concurrent_pages=None,
        inputs=["input.pdf"],
        no_text_hints=True,
        image_backend="builtin-imagegen",
    )


def run_prepare(stdout):
    completed = subprocess.CompletedProcess([], 0, stdout=stdout, stderr="")
    with patch.object(runtime_main.subprocess, "run", return_value=completed), patch.object(
        runtime_main, "run_script", return_value=0
    ), contextlib.redirect_stdout(io.StringIO()), contextlib.redirect_stderr(io.StringIO()) as stderr:
        return runtime_main.cmd_prepare(prepare_args()), stderr.getvalue()


def main():
    with tempfile.TemporaryDirectory() as temp_dir:
        manifest = Path(temp_dir) / "deck_manifest.json"
        manifest.touch()
        warning = "warning: The fitz API is deprecated and will be removed\n"

        result, _ = run_prepare(f"{warning}{manifest}\nrun_id=test\n")
        assert result == 0

        missing = Path(temp_dir) / "missing" / "deck_manifest.json"
        result, stderr = run_prepare(f"{warning}{missing}\nrun_id=test\n")
        assert result == 1
        assert "prepare reported a missing deck_manifest.json path" in stderr


if __name__ == "__main__":
    main()
