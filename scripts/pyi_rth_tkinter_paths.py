import os
import sys
from pathlib import Path


base_dir = Path(getattr(sys, "_MEIPASS", Path(__file__).resolve().parent))
tcl_dir = base_dir / "_tcl_data"
tk_dir = base_dir / "_tk_data"

if tcl_dir.exists():
    os.environ["TCL_LIBRARY"] = str(tcl_dir)
if tk_dir.exists():
    os.environ["TK_LIBRARY"] = str(tk_dir)
