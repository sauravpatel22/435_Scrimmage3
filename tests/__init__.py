"""
Makes the project root importable from the test modules regardless of how
they're invoked (e.g. `python -m unittest discover -s tests`, which by
default only puts `tests/` itself on sys.path, not its parent). This file
runs once, the first time anything under the `tests` package is imported,
and inserts the project root so `import models`, `import
spreadsheet_processor`, etc. resolve the same way they do for main.py.
"""

import sys
from pathlib import Path

_PROJECT_ROOT = Path(__file__).resolve().parent.parent
if str(_PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(_PROJECT_ROOT))
