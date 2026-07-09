# -*- coding: utf-8 -*-
import os, zipfile
from pptx import Presentation

path = r"C:\Users\USER\Desktop\프로젝트\약손_발표.pptx"
print("exists:", os.path.exists(path), "size:", os.path.getsize(path))

# 1) ZIP 무결성
try:
    z = zipfile.ZipFile(path)
    bad = z.testzip()
    print("zip entries:", len(z.namelist()), "bad:", bad)
    print("has [Content_Types].xml:", "[Content_Types].xml" in z.namelist())
except Exception as e:
    print("ZIP ERROR:", e)

# 2) python-pptx 로 열기
try:
    p = Presentation(path)
    print("pptx OK, slides:", len(p.slides._sldIdLst))
except Exception as e:
    print("PPTX ERROR:", repr(e))
