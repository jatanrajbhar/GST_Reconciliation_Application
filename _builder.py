import pathlib
lines = []
def a(s): lines.append(s)
fp = pathlib.Path("c:/Users/tech/Documents/backup/product2/_compare_csv_vs_excel.py")
a("import sys, os, re")
a("sys.stdout.reconfigure(encoding="utf-8")")
a("sys.path.insert(0, "c:/Users/tech/Documents/backup/product2")")
a("")
a("import pandas as pd")
a("import numpy as np")
a("print("builder test OK")")
fp.write_text(chr(10).join(lines), encoding="utf-8")
print("builder done")