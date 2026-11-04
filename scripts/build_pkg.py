"""產生 dist/tcr-pkg.xlsx：上傳到 Claude Project「檔案」的程式包。

知識庫的文字檔不會出現在執行環境的磁碟上，試算表會原檔放進
/mnt/user-data/attachments/。所以把程式包成一個合法的 xlsx（最小工作表＋
tcr/docs/ 底下的原檔），Claude 在對話中解壓即可取得一字不差的程式檔。

用法：python3 scripts/build_pkg.py
每次改了 docs/ 底下的程式，重跑本腳本、commit，再把 dist/tcr-pkg.xlsx 換上專案。
"""
import hashlib
import os
import re
import zipfile

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
FILES = ["analyze.py", "make_deck.py", "build.js", "tagtoo_pptx.js", "tagtoo_assets.js",
         "package.json", "deck_data.example.json", "deck_extra.example.json"]
OUT = os.path.join(ROOT, "dist", "tcr-pkg.xlsx")
STAMP = (2026, 1, 1, 0, 0, 0)  # 固定時間戳，同樣內容產出同樣的檔

ver = re.search(r"Version: ([0-9.]+)",
                open(os.path.join(ROOT, "project-instructions.md"), encoding="utf-8").read()).group(1)

XML = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
NS = "http://schemas.openxmlformats.org"
PARTS = {
    "[Content_Types].xml": f'<Types xmlns="{NS}/package/2006/content-types">'
        f'<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
        f'<Default Extension="xml" ContentType="application/xml"/>'
        f'<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>'
        f'<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/></Types>',
    "_rels/.rels": f'<Relationships xmlns="{NS}/package/2006/relationships">'
        f'<Relationship Id="rId1" Type="{NS}/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>',
    "xl/workbook.xml": f'<workbook xmlns="{NS}/spreadsheetml/2006/main" xmlns:r="{NS}/officeDocument/2006/relationships">'
        f'<sheets><sheet name="readme" sheetId="1" r:id="rId1"/></sheets></workbook>',
    "xl/_rels/workbook.xml.rels": f'<Relationships xmlns="{NS}/package/2006/relationships">'
        f'<Relationship Id="rId1" Type="{NS}/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/></Relationships>',
    "xl/worksheets/sheet1.xml": f'<worksheet xmlns="{NS}/spreadsheetml/2006/main"><sheetData><row r="1">'
        f'<c r="A1" t="inlineStr"><is><t>Commerce Report v{ver} package (code files inside, do not edit)</t></is></c>'
        f'</row></sheetData></worksheet>',
}

os.makedirs(os.path.dirname(OUT), exist_ok=True)
with zipfile.ZipFile(OUT, "w", zipfile.ZIP_STORED) as z:
    for name, body in PARTS.items():
        z.writestr(zipfile.ZipInfo(name, STAMP), XML + body)
    # 版本標記：解壓後可用來確認程式包與指示檔是同一版
    z.writestr(zipfile.ZipInfo("tcr/docs/PKG_VERSION", STAMP), ver + "\n")
    for n in FILES:
        z.writestr(zipfile.ZipInfo("tcr/docs/" + n, STAMP), open(os.path.join(ROOT, "docs", n), "rb").read())

print(f"✓ {OUT}（v{ver}，{os.path.getsize(OUT):,} 位元組）")
for n in FILES:
    print(f"  {n}\t{hashlib.sha256(open(os.path.join(ROOT, 'docs', n), 'rb').read()).hexdigest()[:16]}")
