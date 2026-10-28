#!/usr/bin/env python
"""Bangun index.html dari src/ + pustaka SheetJS yang di-inline.

Urutan: head.html (membuka <script>) -> 01-state -> 02-ribbon -> 03-core ->
04-formula -> 05-features -> 06-tools -> 07-io -> pustaka XLSX -> penutup.
Pustaka diletakkan setelah kode aplikasi karena hanya dipakai saat runtime
(klik Buka/Simpan), sehingga deklarasi `const` di kode tidak tertimpa.
"""
from pathlib import Path
import sys

ROOT = Path(__file__).resolve().parent.parent
ORDER = ["01-state.js", "02-ribbon.js", "03-core.js", "04-formula.js",
         "05-features.js", "06-tools.js", "08-extras.js", "09-workbooks.js", "07-io.js"]
LIB = Path(r"C:\Users\alfan.andrianto\Documents\tracking-klaim-ui\vendor\xlsx.full.min.js")
TAIL = "\n</script>\n</body>\n</html>\n"


def build() -> str:
    head = (ROOT / "src" / "head.html").read_text(encoding="utf-8")
    parts = [(ROOT / "src" / n).read_text(encoding="utf-8") for n in ORDER]
    lib = LIB.read_text(encoding="utf-8") if LIB.exists() else "/* lib hilang */"
    if "</script" in lib:
        raise SystemExit("pustaka memuat </script>, perlu di-escape")
    return head + "\n" + "\n".join(parts) + "\n/* ===== SheetJS (inline) ===== */\n" + lib + TAIL


def main() -> int:
    out = build()
    tgt = ROOT / "index.html"
    tgt.write_text(out, encoding="utf-8", newline="")
    print(f"index.html: {len(out):,} karakter".replace(",", "."))
    return 0


if __name__ == "__main__":
    sys.exit(main())
