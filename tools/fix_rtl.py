#!/usr/bin/env python3
"""משלים הגדרות עברית מימין לשמאל שספריית docx-js לא יודעת לייצר.

docx-js לא חושפת הגדרת כיווניות ברמת המקטע (section), ולכן המסמך כולו
נשאר משמאל לימין גם כשכל פסקה בו מסומנת נכון. הסקריפט מזריק ישירות
ל-XML שני אלמנטים:

    <w:bidi/>        כיווניות בסיס של המקטע — ימין לשמאל
    <w:rtlGutter/>   שוליי הכריכה עוברים לצד ימין, כמו בספר עברי

לפי סכמת ECMA-376 שני אלה חייבים לבוא בתוך <w:sectPr> אחרי <w:pgNumType>
ולפני <w:docGrid>. סדר האלמנטים נאכף — הזרקה במקום הלא נכון פוסלת את הקובץ.
"""
import re
import shutil
import sys
import zipfile
from pathlib import Path

DOC = "word/document.xml"


def patch_section(xml: str) -> tuple[str, int]:
    """מוסיף bidi ו-rtlGutter לכל <w:sectPr> שעוד אין בו."""
    count = 0

    def repl(m: re.Match) -> str:
        nonlocal count
        sect = m.group(0)
        if "<w:bidi/>" in sect:
            return sect
        count += 1
        inject = "<w:bidi/><w:rtlGutter/>"
        # docGrid הוא האלמנט שבא מיד אחרי rtlGutter בסכמה
        if "<w:docGrid" in sect:
            return sect.replace("<w:docGrid", inject + "<w:docGrid", 1)
        # אין docGrid — להוסיף ממש לפני סגירת התג
        return sect.replace("</w:sectPr>", inject + "</w:sectPr>", 1)

    return re.sub(r"<w:sectPr[^>]*>.*?</w:sectPr>", repl, xml, flags=re.S), count


def process(path: Path) -> None:
    with zipfile.ZipFile(path) as z:
        names = z.namelist()
        if DOC not in names:
            raise SystemExit(f"{path}: אין {DOC} בארכיון")
        items = {n: z.read(n) for n in names}

    xml = items[DOC].decode("utf-8")
    patched, n = patch_section(xml)
    if n == 0:
        print(f"  {path.name}: כבר מתוקן, אין שינוי")
        return
    items[DOC] = patched.encode("utf-8")

    tmp = path.with_suffix(".tmp.docx")
    with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as z:
        for name in names:  # שמירה על סדר הכניסות המקורי
            z.writestr(name, items[name])
    shutil.move(str(tmp), str(path))
    print(f"  {path.name}: הוזרקו bidi ו-rtlGutter ל-{n} מקטעים")


if __name__ == "__main__":
    targets = [Path(a) for a in sys.argv[1:]] or sorted(Path("analysis").glob("*.docx"))
    if not targets:
        raise SystemExit("לא נמצאו קבצים לעיבוד")
    print("השלמת הגדרות עברית מימין לשמאל:")
    for t in targets:
        process(t)
