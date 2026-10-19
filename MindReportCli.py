import sys
import os
from pathlib import Path

# Add project root to sys.path
repo_root = Path(__file__).resolve().parent
if str(repo_root) not in sys.path:
    sys.path.insert(0, str(repo_root))

from src.MindReport.main import main

if __name__ == "__main__":
    main()
