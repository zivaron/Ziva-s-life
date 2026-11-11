#!/usr/bin/env python3
"""משלים תיקוני XML שספריית docx-js לא מייצרת נכון בעצמה.

שני תיקונים נפרדים, ושניהם קיימים כי docx-js פשוט לא חושפת שליטה על מה
שהתיקון הזה נוגע בו:

1. כיווניות ברמת המקטע (section) — docx-js לא חושפת הגדרה כזו כלל, ולכן
   המסמך כולו נשאר משמאל לימין גם כשכל פסקה בו מסומנת נכון. מזריקים
   ל-<w:sectPr>:
       <w:bidi/>        כיווניות בסיס של המקטע — ימין לשמאל
       <w:rtlGutter/>   שוליי הכריכה עוברים לצד ימין, כמו בספר עברי
   לפי סכמת ECMA-376 שני אלה חייבים לבוא אחרי <w:pgNumType> ולפני
   <w:docGrid>. סדר האלמנטים נאכף — הזרקה במקום הלא נכון פוסלת את הקובץ.

2. סדר הגבולות בפסקה (<w:pBdr>) — כשמבקשים גבול בארבעה צדדים (למשל
   P(..., {box: color}) ב-lib_docx.js), docx-js תמיד פולטת אותם בסדר
   top, bottom, left, right — קבוע בספרייה, לא משנה באיזה סדר מעבירים
   את האובייקט ב-JS. הסכמה דורשת top, left, bottom, right (ו-between,
   bar אחריהם). התוצאה: כל מסמך עם גבול מלא בפסקה נפסל. מסדרים מחדש.
"""
import re
import shutil
import sys
import zipfile
from pathlib import Path

DOC = "word/document.xml"
PBDR_ORDER = ["top", "left", "bottom", "right", "between", "bar"]


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


def reorder_borders(xml: str) -> tuple[str, int]:
    """מסדרת מחדש את ילדי <w:pBdr> לפי סדר הסכמה. פסקה עם צד אחד בלבד
    (כמו leftBar או RULE) כבר בסדר תקין ולא משתנה — בטוח להריץ על הכול."""
    count = 0

    def repl(m: re.Match) -> str:
        nonlocal count
        inner = m.group(0)
        by_name = {}
        for tag_match in re.finditer(r"<w:(top|left|bottom|right|between|bar)\b[^>]*/>", inner):
            by_name[tag_match.group(1)] = tag_match.group(0)
        if len(by_name) < 2:
            return inner  # צד יחיד — אין מה לסדר
        ordered = "<w:pBdr>" + "".join(by_name[k] for k in PBDR_ORDER if k in by_name) + "</w:pBdr>"
        if ordered != inner:
            count += 1
        return ordered

    return re.sub(r"<w:pBdr>.*?</w:pBdr>", repl, xml, flags=re.S), count


def process(path: Path) -> None:
    with zipfile.ZipFile(path) as z:
        names = z.namelist()
        if DOC not in names:
            raise SystemExit(f"{path}: אין {DOC} בארכיון")
        items = {n: z.read(n) for n in names}

    xml = items[DOC].decode("utf-8")
    xml, n_sect = patch_section(xml)
    xml, n_bdr = reorder_borders(xml)

    if n_sect == 0 and n_bdr == 0:
        print(f"  {path.name}: כבר מתוקן, אין שינוי")
        return
    items[DOC] = xml.encode("utf-8")

    tmp = path.with_suffix(".tmp.docx")
    with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as z:
        for name in names:  # שמירה על סדר הכניסות המקורי
            z.writestr(name, items[name])
    shutil.move(str(tmp), str(path))
    bits = []
    if n_sect:
        bits.append(f"bidi/rtlGutter ל-{n_sect} מקטעים")
    if n_bdr:
        bits.append(f"סדר גבולות תוקן ב-{n_bdr} פסקאות")
    print(f"  {path.name}: " + " · ".join(bits))


if __name__ == "__main__":
    targets = [Path(a) for a in sys.argv[1:]] or sorted(Path("analysis").glob("*.docx"))
    if not targets:
        raise SystemExit("לא נמצאו קבצים לעיבוד")
    print("השלמת הגדרות עברית מימין לשמאל:")
    for t in targets:
        process(t)
