"""Source-checkout entry point for the reference and decimal audit."""

import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "src"))
from format_research_word.audit import main

if __name__ == "__main__":
    main()
