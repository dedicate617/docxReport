import platform
import subprocess
import gzip
import json
import os
import sys
import shutil
from pathlib import Path


class RunnerLauncher:
    @staticmethod
    def get_runner_binary_path(base_dir: Path) -> Path:
        system = platform.system().lower()
        machine = platform.machine().lower()

        if system == "windows":
            exe_dir = Path(sys.executable).resolve().parent
            candidates = [
                base_dir / "bin" / "win-x64" / "CrossReportRunner.exe",
                base_dir / "CrossReportRunner.exe",
                exe_dir / "bin" / "win-x64" / "CrossReportRunner.exe",
                exe_dir / "CrossReportRunner.exe",
                Path(__file__).resolve().parent / "bin" / "win-x64" / "CrossReportRunner.exe",
                Path(__file__).resolve().parent.parent / "bin" / "win-x64" / "CrossReportRunner.exe",
                Path(__file__).resolve().parent.parent.parent / "bin" / "win-x64" / "CrossReportRunner.exe",
                base_dir / "src" / "CrossReportRunner" / "bin" / "Debug" / "net8.0" / "win-x64" / "CrossReportRunner.exe",
            ]
            for c in candidates:
                if c.exists():
                    return c
            which_runner = shutil.which("CrossReportRunner.exe") or shutil.which("CrossReportRunner")
            if which_runner:
                return Path(which_runner)
            return base_dir / "bin" / "win-x64" / "CrossReportRunner.exe"
        elif system == "linux":
            if "aarch64" in machine or "arm64" in machine:
                return base_dir / "bin" / "linux-arm64" / "CrossReportRunner"
            elif "arm" in machine:  # ARM32 (armv7l, armhf)
                return base_dir / "bin" / "linux-arm" / "CrossReportRunner"
            else:
                return base_dir / "bin" / "linux-x64" / "CrossReportRunner"
        else:
            raise NotImplementedError(f"Unsupported OS/Machine: {system}/{machine}")

    @classmethod
    def execute_and_stream(cls, base_dir: Path, dsdef: str, dspara: str, vpara: str) -> dict:
        bin_path = cls.get_runner_binary_path(base_dir)
        if not bin_path.exists():
            raise FileNotFoundError(f"Runner binary not found: {bin_path}")

        cmd = [
            str(bin_path),
            f"--dsdef={dsdef}",
            f"--dspara={dspara}",
            f"--vpara={vpara}",
            "--output-mode=pipe-gzip"
        ]

        proc = subprocess.Popen(cmd, stdout=subprocess.PIPE, stderr=subprocess.PIPE)
        raw_gzip, stderr = proc.communicate()

        if proc.returncode != 0:
            err_msg = stderr.decode('utf-8', errors='ignore')
            raise RuntimeError(f"CrossReportRunner failed (code {proc.returncode}): {err_msg}")

        decompressed = gzip.decompress(raw_gzip)
        return json.loads(decompressed.decode('utf-8'))
