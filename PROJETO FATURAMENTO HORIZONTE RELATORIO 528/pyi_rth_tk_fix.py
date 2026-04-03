import os
import sys

# Garante que o executavel encontre os arquivos Tcl/Tk embarcados.
if getattr(sys, "frozen", False):
    base = getattr(sys, "_MEIPASS", os.path.dirname(sys.executable))
    os.environ.setdefault("TCL_LIBRARY", os.path.join(base, "tcl", "tcl8.6"))
    os.environ.setdefault("TK_LIBRARY", os.path.join(base, "tcl", "tk8.6"))
