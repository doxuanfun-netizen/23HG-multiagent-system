# -*- coding: utf-8 -*-
"""
Làm sạch siêu dữ liệu máy cá nhân trong file Office (xlsx/xlsm/xlam/docx) mà không đụng tới nội dung.

- absPath: Excel ghi thư mục lưu file cuối cùng vào xl/workbook.xml (vd. C:\\Users\\<tên>\\Downloads\\<dự án>\\).
  Thông tin này lộ tên người dùng và tên dự án khi file được chia sẻ; xóa không ảnh hưởng gì tới bảng tính.
- Liên kết ngoài (--lien-ket-ngoai): workbook trỏ tới file Excel khác trên máy người tạo (thường sót lại khi chép
  sheet từ file mẫu). Chỉ xóa khi KHÔNG ô công thức nào dùng tới; các tên vùng (definedName) trỏ ra ngoài cũng bị
  xóa và được liệt kê trong báo cáo. Nếu có ô công thức dùng liên kết thì dừng, không sửa file.

Các phần khác của gói (vbaProject.bin, sheet, style...) được chép nguyên byte; ghi lại zip theo đúng thứ tự và
thời gian của bản gốc nên chạy lại cho kết quả giống nhau (các bản sao cùng tên vẫn giống hệt nhau).

Chạy:  python -m tools.office_scrub <file...> [--lien-ket-ngoai] [--kiem-tra]
"""

from __future__ import annotations

import argparse
import os
import re
import shutil
import sys
import tempfile
import zipfile
from typing import Dict, List

ABS_PATH_BLOCK = re.compile(
    r'<mc:AlternateContent[^>]*>\s*<mc:Choice Requires="x15[a-z]*">\s*<x15ac:absPath\b[^>]*/>\s*</mc:Choice>\s*'
    r'</mc:AlternateContent>')
ABS_PATH_ANY = re.compile(r"<x15ac:absPath\b")
EXT_REL_TYPE = "relationships/externalLink"
OFFICE_SUFFIXES = (".xlsx", ".xlsm", ".xlam", ".xltx", ".xltm", ".docx", ".docm", ".pptx")


class ScrubError(RuntimeError):
    pass


def _external_link_parts(names: List[str]) -> List[str]:
    return [n for n in names if n.startswith("xl/externalLinks/")]


def inspect(path: str) -> Dict[str, object]:
    """Báo cáo, không sửa: còn absPath không, có bao nhiêu liên kết ngoài, ô công thức nào dùng liên kết ngoài."""
    with zipfile.ZipFile(path) as z:
        names = z.namelist()
        wb = z.read("xl/workbook.xml").decode("utf-8") if "xl/workbook.xml" in names else ""
        formulas = []
        for n in names:
            if n.startswith("xl/worksheets/") and n.endswith(".xml"):
                text = z.read(n).decode("utf-8", "ignore")
                if re.search(r"<f\b[^>]*>[^<]*\[\d+\][^<]*</f>", text):
                    formulas.append(n)
    ext_names = re.findall(r'<definedName\b[^>]*name="([^"]+)"[^>]*>[^<]*\[\d+\][^<]*</definedName>', wb)
    return {"absPath": bool(ABS_PATH_ANY.search(wb)),
            "lien_ket_ngoai": len([n for n in _external_link_parts(names) if n.endswith(".xml")
                                   and "/_rels/" not in n]),
            "ten_vung_tro_ra_ngoai": ext_names,
            "o_cong_thuc_dung_lien_ket": formulas}


def _drop_external_links(parts: Dict[str, bytes], report: Dict[str, object]) -> None:
    wb = parts["xl/workbook.xml"].decode("utf-8")
    rels_name = "xl/_rels/workbook.xml.rels"
    rels = parts[rels_name].decode("utf-8")
    ext_ids = re.findall(r'<Relationship\b[^>]*Id="([^"]+)"[^>]*Type="[^"]*' + EXT_REL_TYPE + '"', rels)
    ext_ids += [i for i in re.findall(r'<Relationship\b[^>]*Type="[^"]*' + EXT_REL_TYPE + '"[^>]*Id="([^"]+)"', rels)
                if i not in ext_ids]
    if not ext_ids:
        return
    rels = re.sub(r'<Relationship\b[^>]*' + EXT_REL_TYPE + r'"[^>]*/>', "", rels)
    wb = re.sub(r"<externalReferences>.*?</externalReferences>", "", wb, flags=re.S)
    dropped = re.findall(r'<definedName\b[^>]*name="([^"]+)"[^>]*>[^<]*\[\d+\][^<]*</definedName>', wb)
    wb = re.sub(r'<definedName\b[^>]*>[^<]*\[\d+\][^<]*</definedName>', "", wb)
    wb = re.sub(r"<definedNames>\s*</definedNames>", "", wb)
    parts["xl/workbook.xml"] = wb.encode("utf-8")
    parts[rels_name] = rels.encode("utf-8")
    for n in [n for n in parts if n.startswith("xl/externalLinks/")]:
        del parts[n]
    ct = parts["[Content_Types].xml"].decode("utf-8")
    parts["[Content_Types].xml"] = re.sub(r'<Override\b[^>]*PartName="/xl/externalLinks/[^"]*"[^>]*/>', "",
                                          ct).encode("utf-8")
    report["lien_ket_ngoai_da_xoa"] = len(ext_ids)
    report["ten_vung_da_xoa"] = dropped


def scrub(path: str, lien_ket_ngoai: bool = False) -> Dict[str, object]:
    """Sửa tại chỗ. Trả về báo cáo những gì đã xóa; file không đổi nếu không có gì cần xóa."""
    before = inspect(path) if path.lower().endswith((".xlsx", ".xlsm", ".xlam", ".xltx", ".xltm")) else {}
    if lien_ket_ngoai and before.get("o_cong_thuc_dung_lien_ket"):
        raise ScrubError(f"{path}: có ô công thức dùng liên kết ngoài ({before['o_cong_thuc_dung_lien_ket']}) — "
                         f"không xóa liên kết, cần sửa công thức trước")
    with zipfile.ZipFile(path) as z:
        infos = z.infolist()
        parts = {i.filename: z.read(i.filename) for i in infos}
    report: Dict[str, object] = {"absPath_da_xoa": False}
    wb = parts.get("xl/workbook.xml")
    if wb is not None:
        text = wb.decode("utf-8")
        new = ABS_PATH_BLOCK.sub("", text)
        if ABS_PATH_ANY.search(new):
            raise ScrubError(f"{path}: absPath có dạng lạ, không xóa an toàn được")
        if new != text:
            parts["xl/workbook.xml"] = new.encode("utf-8")
            report["absPath_da_xoa"] = True
    if lien_ket_ngoai and "xl/_rels/workbook.xml.rels" in parts:
        _drop_external_links(parts, report)
    if not report["absPath_da_xoa"] and not report.get("lien_ket_ngoai_da_xoa"):
        return report
    fd, tmp = tempfile.mkstemp(suffix=os.path.splitext(path)[1], dir=os.path.dirname(os.path.abspath(path)))
    os.close(fd)
    try:
        with zipfile.ZipFile(tmp, "w") as out:
            for info in infos:
                if info.filename in parts:
                    out.writestr(info, parts[info.filename], compress_type=info.compress_type)
        shutil.move(tmp, path)
    finally:
        if os.path.exists(tmp):
            os.remove(tmp)
    return report


def main(argv=None) -> int:
    ap = argparse.ArgumentParser(description="Xóa đường dẫn máy cá nhân (absPath) và liên kết ngoài thừa trong file Office")
    ap.add_argument("files", nargs="+")
    ap.add_argument("--lien-ket-ngoai", action="store_true", help="xóa cả liên kết ngoài không ô công thức nào dùng")
    ap.add_argument("--kiem-tra", action="store_true", help="chỉ báo cáo, không sửa")
    a = ap.parse_args(argv)
    loi = 0
    for f in a.files:
        if not f.lower().endswith(OFFICE_SUFFIXES):
            continue
        try:
            r = inspect(f) if a.kiem_tra else scrub(f, a.lien_ket_ngoai)
        except ScrubError as e:
            print(f"[LỖI] {e}")
            loi += 1
            continue
        print(f"{f}: {r}")
    return 1 if loi else 0


if __name__ == "__main__":
    sys.exit(main())
