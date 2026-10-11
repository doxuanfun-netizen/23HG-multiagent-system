# -*- coding: utf-8 -*-
"""Dò LibreOffice Calc thật sự dùng được (không chỉ có file soffice).

Máy có `soffice` nhưng thiếu gói libreoffice-calc sẽ có binary nhưng không chuyển đổi được file.
Dò bằng cách chuyển một CSV nhỏ sang XLSX; nếu không ra file thì coi như không có Calc để test bỏ qua
đúng lý do thay vì báo lỗi.
"""

import functools
import os
import shutil
import subprocess
import tempfile


@functools.lru_cache(maxsize=1)
def calc_available():
    soffice = shutil.which("soffice") or shutil.which("libreoffice")
    if not soffice:
        return False
    with tempfile.TemporaryDirectory() as td:
        src = os.path.join(td, "probe.csv")
        with open(src, "w", encoding="utf-8") as f:
            f.write("a,b\n1,2\n")
        out = os.path.join(td, "out")
        os.makedirs(out)
        try:
            subprocess.run([soffice, "--headless", "--norestore",
                            "-env:UserInstallation=file://" + os.path.join(td, "profile"),
                            "--convert-to", "xlsx", "--outdir", out, src],
                           capture_output=True, timeout=180)
        except (OSError, subprocess.SubprocessError):
            return False
        return os.path.exists(os.path.join(out, "probe.xlsx"))
