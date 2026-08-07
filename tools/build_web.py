#!/usr/bin/env python3
"""בונה דף אינטרנט מהגרסה האחרונה של הספר, כדי שתהיה גישה אליו מכל מקום.

הטקסט נלקח ישירות מקובץ ה-Word של זיוה — לא מועתק ביד — כדי שלא ייווצרו
הפרשים בין הדף לבין כתב היד.
"""
import html
import json
import re
import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET

W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
SRC = Path("manuscript/סאני_גרסה_אחרונה_של_זיוה.docx")
OUT = Path("web/הספר.html")

# ── קריאת הפסקאות מקובץ ה-Word ────────────────────────────────────────────
root = ET.fromstring(zipfile.ZipFile(SRC).read("word/document.xml"))
paras: list[str] = []
for p in root.iter(W + "p"):
    buf = []
    for n in p.iter():
        tag = n.tag.split("}")[1]
        if tag == "t":
            buf.append(n.text or "")
        elif tag == "br":
            buf.append(" ")           # מפריד שורה בתוך אותה פסקה
    s = "".join(buf).strip()
    if s:
        paras.append(s)

CH_STARTS = [i for i, s in enumerate(paras) if re.fullmatch(r"פרק \d", s)]
GLOSS = next(i for i, s in enumerate(paras) if s.startswith("מילים חדשות"))

CHAPTERS = [
    ("ההבטחה", 0, None), ("תובלי וכוח החברות", 1, None),
    ("יונתן וכוח ההקשבה", 2, None), ("אימרי וכוח האומץ", 3, None),
    ("אופיר וכוח הסקרנות", 3, "אבד עלה החברות"), ("תמרה וכוח התקווה", 4, None),
    ("האור חזר אל היער", 5, None),
]

LEAF_STARTS = ("מִי שֶׁהוֹלֵךְ לְבַד", "מִי שֶׁמַּקְשִׁיב", "אֹמֶץ אֵינוֹ", "רמז אחד קטן", "תִּקְוָה הִיא")
REFRAIN_STARTS = ("עוֹד עֲלֵה זָהָב", "את הָעָלֶה הָאַחֲרוֹן")
INTERLUDE_STARTS = ("רחוק משם, בקרחת היער", "ובקרחת היער, העץ", "ובקרחת היער כפף",
                    "ורחוק משם, במורד הנחל", "והעץ העתיק חיכה")


def classify(s: str) -> str:
    if s.startswith("אֲנַחְנוּ יַחַד"):
        return "song"
    if s.startswith("בלב היער הסמוך"):
        return "scroll"
    if s.startswith(LEAF_STARTS):
        return "leaf"
    if s.startswith(REFRAIN_STARTS):
        return "refrain"
    if s.startswith(INTERLUDE_STARTS):
        return "interlude"
    if s.startswith("כל עוד ימשיכו להאמין"):
        return "closing"
    return "p"


def render(s: str) -> str:
    return html.escape(s).replace(" ", "<br>")


# ── גוף הספר ──────────────────────────────────────────────────────────────
body = []
counts = {}
for n, (title, leaves, lost) in enumerate(CHAPTERS):
    start = CH_STARTS[n] + 2                       # לדלג על "פרק N" ועל הכותרת
    end = CH_STARTS[n + 1] if n + 1 < len(CH_STARTS) else GLOSS
    dots = "".join(
        f'<i class="dot{" on" if k < leaves else ""}"></i>' for k in range(5)
    ) + ('<i class="dot lost" title="אבד בנחל"></i>' if lost else "")
    counts[n] = leaves
    body.append(f"""
<section class="chapter" id="p{n+1}">
  <header class="ch-head">
    <p class="eyebrow">פרק {n+1}</p>
    <h2>{html.escape(title)}</h2>
    <p class="pocket"><span class="dots">{dots}</span>
      <span class="pocket-txt">{'עלי זהב בכיס: ' + str(leaves) if leaves else 'הכיס עדיין ריק'}{' · ' + lost if lost else ''}</span>
    </p>
  </header>""")
    for s in paras[start:end]:
        kind = classify(s)
        tag = "blockquote" if kind in ("scroll", "song", "leaf") else "p"
        body.append(f'  <{tag} class="{kind}">{render(s)}</{tag}>')
    body.append("</section>")

# ── מילון ─────────────────────────────────────────────────────────────────
gl = paras[GLOSS + 3:]                              # לדלג על הכותרת ועל שורות הטבלה
gloss_rows = "".join(
    f"<div class=g-row><dt>{html.escape(gl[i])}</dt><dd>{html.escape(gl[i+1])}</dd></div>"
    for i in range(0, len(gl) - 1, 2)
)

FIXES = [
    ("1", "פרק 1", "תשובתה-  SUNNY LIKE THE SUN", "להסיר את האנגלית — זו המילה הלועזית היחידה בספר, והיא גם שורפת את הסוף"),
    ("2", "פרק 1", "משהו בהיה שונה", "משהו היה שונה"),
    ("3", "פרק 2", "לבד אף אחד לא מוצאים כלום", "לבד אף אחד לא מוצא כלום"),
    ("4", "פרק 4", "מה מסתתר שם? חשש.", "מה מסתתר שם? חשב."),
    ("5", "פרק 4", '"גם אני," וגם אני לחשו אופיר ותמרה.', '"גם אני," לחשה אופיר. "וגם אני," לחשה תמרה.'),
    ("6", "פרק 5", '"איך נתקדם"?', '"איך נתקדם?"'),
    ("7", "פרק 6", "המשימה גדולה מידי", "המשימה גדולה מדי"),
    ("8", "פרק 6", "היא חילקה אותה לשישה", 'היא חילקה אותו — "כריך" הוא זכר'),
    ("9", "פרק 7", "סאני והחברוה הבינו", "סאני וחבריה הבינו"),
    ("10", "לאורך הספר", "רווח לפני נקודה או פסיק — 23 מקומות", '"העצים ." ← "העצים." · "אליה ," ← "אליה,"'),
    ("11", "לאורך הספר", "28 מקפים קצרים מול 7 ארוכים", "לאחד — בעברית ספרותית המקף הארוך הוא הנכון"),
    ("12", "לאורך הספר", 'ניקוד חלקי: "כּל" מול "כָּל" · "יִתגשם"', "לנקד שורה שלמה או לוותר על הניקוד בה"),
]

BIG = [
    ("רגע השם נהרס",
     "העברת את גילוי המשמעות של השם לעמוד הראשון — ובאנגלית — ומחקת את הסצנה בסוף שבה אופיר שואלת "
     "&rdquo;סאני, איך קוראים לך?&ldquo; ועונה &rdquo;שמש&ldquo;. הבדיחה מסופרת עכשיו לפני שהיא נבנתה. "
     "הכוח שלה היה שהקורא מגלה אותה יחד עם הדמויות, בעמוד האחרון, אחרי חמישה פרקים שבהם ראה את סאני מאירה לאחרים.",
     "להשאיר בעמוד הראשון רק את הרמז העברי שכבר כתבת — &rdquo;אמא אומרת שהיא נראית כמו שמש קטנה שיצאה לטייל&ldquo; — "
     "ולהחזיר את סצנת הגילוי לסוף."),
    ("סאני כבר לא מסבירה מה עשתה",
     "פיזרת לאורך הספר חמישה רגעים שבהם סאני מזהה מי צריך לפעול: היא הסתכלה על תובלי, הסתובבה אל יונתן, "
     "שתקה כי היה תורו של אימרי, אמרה שאופיר עוד לא דיברה — ובפרק 6 אף אחד לא הסתכל בחזרה. "
     "חמשת הרגעים האלה הם מנגנון, והוא עובד רק אם בסוף מישהו אומר בקול מה הם היו.",
     "להחזיר, אחרי &rdquo;וכל השאר נהיה מואר וגלוי&ldquo;: &rdquo;זה מה שעשיתי כל הדרך. לא מצאתי אף עלה. "
     "אני רק ידעתי, בכל פעם, על מי להסתכל.&ldquo; ואחריו תובלי: &rdquo;זה בכלל לא כוח.&ldquo; ויונתן: &rdquo;זה הכוח.&ldquo;"),
    ("החרוז בשיר נשבר",
     "שינית את השורה האחרונה מ&rdquo;יִתְעוֹרֵר&ldquo; ל&rdquo;יתגשם&ldquo;. &rdquo;לְוַתֵּר&ldquo; ו&rdquo;יִתְגַּשֵּׁם&ldquo; "
     "לא מתחרזים, והמילה גם ארוכה בהברה. השיר מופיע שלוש פעמים — הוא ההמנון של החבורה, הדבר שילדים אמורים לשיר בקול.",
     "&rdquo;יִתְעוֹרֵר&ldquo; היא גם בדיוק המילה שמתארת את מה שקורה לעץ בסוף. ואם &rdquo;יתגשם&ldquo; חשובה לך, "
     "אפשר לשנות את השורה שלפניה: &rdquo;לֵב אֶל לֵב, וְיַחַד נַחְלֹם — כָּל מַה שֶּׁחָלַמְנוּ יִתְגַּשֵּׁם.&ldquo;"),
]

SMALL = [
    ("המסר הסוגר השתנה", "&rdquo;בכל אחד ואחד מהם טמונים כל חמשת הכוחות&ldquo; מבטל בשורה אחת את החלוקה שהספר בנוי עליה — פרק, כוח ועלה לכל ילד. שקלי: &rdquo;לבד, לאף אחד מהם לא היה מספיק. רק יחד היה להם הכול.&ldquo;"),
    ("ההומור התכווץ", "מחמישה רגעים מצחיקים נשארו שניים וחצי. בכולם הוחלפה שורה עם עוקץ בשורה שמסבירה. הכי כדאי להחזיר: הוויכוח של אימרי בפרק 2, ו&rdquo;פחדת&ldquo; / &rdquo;פחדתי&ldquo; בפרק 4."),
    ("הסבר כפול חזר בפרק 2", "&rdquo;זה אומר שרק כשהחלטנו ללכת ביחד מצאנו את עלה הזהב&ldquo; ומיד אחריו המספר אומר את אותו דבר — ומעליהם הפתגם על העלה. שלוש פעמים לאותו רעיון."),
    ("הצליל בפרק 3", "&rdquo;נמוך. ארוך. מתגלגל&ldquo; ואז מתברר שזו שקית דובשניות. שקית נשמעת גבוהה ומרשרשת. או לשנות את התיאור, או להחזיר את הבטן של תמרה."),
    ("&rdquo;תמרה גם שלי&ldquo;", "התוספת &rdquo;חברה שלי&ldquo; מרככת את המכה. הכוח היה בקיצור — שלוש מילים שאומרות שאדם שווה יותר מעלה זהב, בלי להסביר."),
    ("הפתגם של אופיר", "&rdquo;רמז אחד קטן פותח שער גדול&ldquo; מתאר את הסצנה. הכוח של אופיר הוא לשאול — ולכן &rdquo;שאלה אחת קטנה פותחת דלת גדולה&ldquo; קשור אליה ישירות יותר."),
    ("הפזמון הסוגר", "&rdquo;את הָעָלֶה הָאַחֲרוֹן&ldquo; — ה&rdquo;את&ldquo; שובר את המשקל של הפזמון שחוזר ארבע פעמים לפניו."),
    ("הראה במקום לספר", "בפרק 3, &rdquo;ניתן היה לראות את ההתרגשות בעיניו&ldquo; מספר לנו. &rdquo;האוזניים שלו הסמיקו&ldquo; מראה — וגם מצחיק וגם נשאר בזיכרון."),
]

GOOD = [
    ("&rdquo;תודה לסאני ולחבורת חמשת הכוחות&ldquo;", "טוב יותר ממה שהצעתי. העץ מבחין בין סאני לחבורה — וזו בדיוק ההבחנה שהספר עושה. שורה אחת שעושה עבודה של פסקה."),
    ("&rdquo;עלה זהב קטן ועליו כתוב — תודה&ldquo;", "גם זה טוב יותר. עץ הגינה מדבר באותה שפה של עלי הזהב, והמעגל נסגר בלי מילה מיותרת."),
]

rows = "".join(
    f"<tr><td class=num>{n}</td><td class=where>{html.escape(w)}</td>"
    f"<td class=was>{html.escape(a)}</td><td class=now>{b if '&' in b else html.escape(b)}</td></tr>"
    for n, w, a, b in FIXES)
big = "".join(
    f"<article class=big><h3>{t}</h3><p>{d}</p><p class=fix><span>ההצעה</span> {f}</p></article>"
    for t, d, f in BIG)
small = "".join(f"<div class=note><h4>{t}</h4><p>{d}</p></div>" for t, d in SMALL)
good = "".join(f"<div class=note good><h4>{t}</h4><p>{d}</p></div>" for t, d in GOOD)
nav = "".join(
    f'<a href="#p{i+1}"><span class=n>{i+1}</span>{html.escape(t)}</a>'
    for i, (t, _, _) in enumerate(CHAPTERS))

words = sum(len(p.split()) for p in paras)

CSS = """
:root{
  --ground:#F5F7F2; --surface:#FFFFFF; --raise:#EDF1E8;
  --ink:#1C231C; --muted:#5B665A; --faint:#8A957F;
  --gold:#8A6A12; --gold-soft:#F0E7CC; --rule:#DDE3D8; --shadow:rgba(28,35,28,.07);
  --serif:'Frank Ruhl Libre','David Libre',David,'Times New Roman',Georgia,serif;
  --sans:'Assistant',Heebo,'Arial Hebrew',Arial,'Segoe UI',system-ui,sans-serif;
}
@media (prefers-color-scheme:dark){:root:not([data-theme="light"]){
  --ground:#131711; --surface:#1B2018; --raise:#232A20;
  --ink:#E8EDE4; --muted:#9EAA9A; --faint:#78846F;
  --gold:#D9B44A; --gold-soft:#332B15; --rule:#2C332A; --shadow:rgba(0,0,0,.4);
}}
:root[data-theme="dark"]{
  --ground:#131711; --surface:#1B2018; --raise:#232A20;
  --ink:#E8EDE4; --muted:#9EAA9A; --faint:#78846F;
  --gold:#D9B44A; --gold-soft:#332B15; --rule:#2C332A; --shadow:rgba(0,0,0,.4);
}
*{box-sizing:border-box}
body{margin:0;background:var(--ground);color:var(--ink);font-family:var(--sans);
  direction:rtl;text-align:right;line-height:1.7;-webkit-text-size-adjust:100%}
.wrap{max-width:44rem;margin:0 auto;padding:0 1.25rem}

/* ראש הדף */
.masthead{padding:4.5rem 0 2.5rem;border-bottom:1px solid var(--rule);margin-bottom:2.5rem}
.masthead h1{font-family:var(--serif);font-size:clamp(2.1rem,6vw,3.1rem);line-height:1.15;
  margin:0 0 .5rem;font-weight:700;text-wrap:balance}
.masthead .sub{color:var(--muted);font-size:1.05rem;margin:0 0 1.5rem}
.meta{display:flex;flex-wrap:wrap;gap:.5rem;list-style:none;padding:0;margin:0;
  font-size:.82rem;color:var(--faint)}
.meta li{background:var(--raise);border-radius:999px;padding:.3rem .8rem;
  font-variant-numeric:tabular-nums}
.meta li.now{color:var(--gold);background:var(--gold-soft);font-weight:600}

/* ניווט */
nav.toc{display:grid;grid-template-columns:repeat(auto-fit,minmax(13rem,1fr));gap:.4rem;
  margin:0 0 3.5rem}
nav.toc a{display:flex;align-items:baseline;gap:.6rem;padding:.6rem .8rem;border-radius:.4rem;
  text-decoration:none;color:var(--ink);font-size:.95rem;transition:background .15s}
nav.toc a:hover,nav.toc a:focus-visible{background:var(--raise)}
nav.toc .n{font-family:var(--serif);color:var(--gold);font-size:.95rem;
  font-variant-numeric:tabular-nums;min-width:1ch}

/* פרק */
.chapter{margin:0 0 4.5rem}
.ch-head{margin:0 0 2rem;padding-top:1.5rem;border-top:1px solid var(--rule)}
.eyebrow{font-size:.75rem;letter-spacing:.14em;color:var(--faint);margin:0 0 .35rem}
.ch-head h2{font-family:var(--serif);font-size:clamp(1.5rem,4vw,2rem);margin:0 0 .7rem;
  font-weight:700;text-wrap:balance}
.pocket{display:flex;align-items:center;gap:.6rem;margin:0;font-size:.8rem;color:var(--faint)}
.dots{display:inline-flex;gap:.28rem}
.dot{width:.5rem;height:.5rem;border-radius:50%;border:1px solid var(--rule);display:block}
.dot.on{background:var(--gold);border-color:var(--gold)}
.dot.lost{border-color:var(--gold);border-style:dashed;background:none;opacity:.6}

/* גוף הסיפור */
.chapter p,.chapter blockquote{font-family:var(--serif);font-size:1.14rem;line-height:1.85}
.chapter p.p{margin:0 0 1.15rem}
blockquote{margin:2rem 0;padding:0;text-align:center}
blockquote.leaf{color:var(--gold);font-weight:600;font-size:1.2rem;
  padding:1.4rem 1rem;background:var(--gold-soft);border-radius:.5rem}
blockquote.song{color:var(--gold);font-size:1.15rem;line-height:2}
blockquote.scroll{color:var(--muted);font-style:italic;border-inline-start:2px solid var(--gold);
  padding-inline-start:1.2rem;text-align:start;margin:2rem 0}
p.refrain{text-align:center;color:var(--gold);margin:2.2rem 0;line-height:2;font-size:1.1rem}
p.interlude{margin:2.5rem auto 0;max-width:30rem;color:var(--muted);font-style:italic;
  font-size:1rem;text-align:center;padding-top:1.5rem;border-top:1px solid var(--rule)}
p.closing{text-align:center;color:var(--gold);font-size:1.2rem;line-height:2;margin:2.5rem 0 0}

/* מילון */
.gloss{border-top:1px solid var(--rule);padding-top:2.5rem;margin-bottom:4.5rem}
.g-row{display:grid;grid-template-columns:9rem 1fr;gap:1rem;padding:.7rem 0;
  border-bottom:1px solid var(--rule)}
.g-row dt{font-family:var(--serif);font-weight:700;color:var(--gold);margin:0}
.g-row dd{margin:0;color:var(--muted);font-size:.95rem}
@media(max-width:34rem){.g-row{grid-template-columns:1fr;gap:.2rem}}

/* הערות */
.notes{background:var(--surface);border-top:1px solid var(--rule);padding:3.5rem 0 4rem;
  margin-top:2rem}
.notes h2{font-family:var(--serif);font-size:1.9rem;margin:0 0 .5rem;text-wrap:balance}
.notes .lede{color:var(--muted);margin:0 0 2.5rem}
h3.sec{font-size:.78rem;letter-spacing:.14em;color:var(--faint);margin:3rem 0 1rem;
  font-weight:600}
.tbl{overflow-x:auto;border:1px solid var(--rule);border-radius:.5rem}
table{border-collapse:collapse;width:100%;font-size:.88rem;min-width:34rem}
th,td{text-align:right;padding:.6rem .75rem;border-bottom:1px solid var(--rule);vertical-align:top}
thead th{background:var(--raise);font-size:.76rem;letter-spacing:.06em;color:var(--muted);
  font-weight:600;position:sticky;top:0}
tr:last-child td{border-bottom:0}
td.num{color:var(--faint);font-variant-numeric:tabular-nums;width:2.5rem}
td.where{color:var(--muted);white-space:nowrap;font-size:.82rem}
td.was{color:var(--muted);text-decoration:line-through;text-decoration-color:var(--faint)}
td.now{color:var(--ink)}
.big{border-inline-start:2px solid var(--gold);padding:.2rem 1.2rem;margin:0 0 2rem}
.big h3{font-family:var(--serif);font-size:1.3rem;margin:0 0 .5rem}
.big p{margin:0 0 .8rem;color:var(--muted);font-size:.98rem}
.fix{background:var(--gold-soft);border-radius:.45rem;padding:.85rem 1rem;color:var(--ink)!important}
.fix span{display:block;font-size:.72rem;letter-spacing:.12em;color:var(--gold);
  font-weight:700;margin-bottom:.3rem}
.grid{display:grid;grid-template-columns:repeat(auto-fit,minmax(16rem,1fr));gap:1rem}
.note{background:var(--raise);border-radius:.5rem;padding:1rem 1.1rem}
.note h4{margin:0 0 .4rem;font-size:.98rem}
.note p{margin:0;font-size:.9rem;color:var(--muted);line-height:1.65}
.note.good{background:var(--gold-soft)}
.note.good h4{color:var(--gold)}
footer{color:var(--faint);font-size:.85rem;padding:2.5rem 0;text-align:center}
@media(prefers-reduced-motion:reduce){*{transition:none!important}}
"""

HTML = f"""<title>סאני וחבורת חמשת הכוחות — הגרסה האחרונה</title>
<style>{CSS}</style>
<div class="wrap">
  <header class="masthead">
    <h1>סאני וחבורת חמשת הכוחות</h1>
    <p class="sub">סיפור על חברות, אומץ ועל הכוחות שיש בכל אחד מאיתנו · מאת זיוה רון</p>
    <ul class="meta">
      <li class="now">הגרסה האחרונה</li>
      <li>{words:,} מילים</li><li>7 פרקים</li><li>6 ילדים</li>
      <li>12 תיקוני חובה</li>
    </ul>
  </header>
  <nav class="toc">{nav}<a href="#hearot"><span class="n">◆</span>הערות ותיקונים</a></nav>
{''.join(body)}
  <section class="gloss">
    <h2 style="font-family:var(--serif);font-size:1.6rem;margin:0 0 1.2rem">מילים חדשות שפגשנו ביער</h2>
    <dl style="margin:0">{gloss_rows}</dl>
  </section>
</div>

<section class="notes" id="hearot">
  <div class="wrap">
    <h2>הערות לגרסה הזו</h2>
    <p class="lede">סריקה מלאה של הנוסח, בהשוואה לנוסח הקודם. הגרסה הזו טובה משמעותית
      מכתב היד המקורי — שמרת על השלד ועל הדרמה ושינית עשרות ניסוחים לקול שלך.
      שלושה דברים כאן מוחקים את הרגעים הכי חזקים בספר, ולכן הם ראשונים.</p>

    <h3 class="sec">שלושה דברים ששווה להחזיר</h3>
    {big}

    <h3 class="sec">תיקוני חובה — לא ענייני טעם</h3>
    <div class="tbl"><table>
      <thead><tr><th>#</th><th>איפה</th><th>במקור</th><th>התיקון</th></tr></thead>
      <tbody>{rows}</tbody>
    </table></div>

    <h3 class="sec">הערות תוכן נוספות</h3>
    <div class="grid">{small}</div>

    <h3 class="sec">ושתי תוספות שלך ששיפרו את הספר</h3>
    <div class="grid">{good}</div>

    <footer>הדף נוצר מקובץ ה־Word של הגרסה האחרונה · הטקסט הוא של זיוה רון</footer>
  </div>
</section>
"""

OUT.parent.mkdir(exist_ok=True)
OUT.write_text(HTML, encoding="utf-8")
print(f"נכתב {OUT} — {len(HTML)/1024:.1f} KB · {words} מילים · {len(CH_STARTS)} פרקים")
