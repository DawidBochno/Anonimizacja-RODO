#!/usr/bin/env python3
"""Anonimizacja RODO. Wyszukuje dane osobowe w DOCX i PDF (PESEL, NIP, REGON,
nr dowodu, IBAN, e-mail, telefon + wlasna lista slow) i trwale je usuwa:
w PDF tekst jest wycinany z pliku (nie tylko zakrywany), w DOCX zamieniany
na gwiazdki. Czysci tez metadane (autor itp.).

Uruchomienie: python anonimizacja.py            (GUI)
              python anonimizacja.py --selftest (test logiki)
"""
import os
import re
import sys
import threading
import traceback
from collections import Counter

if getattr(sys, "frozen", False):
    APP_DIR = os.path.dirname(sys.executable)
else:
    APP_DIR = os.path.dirname(os.path.abspath(__file__))

# ------------------------------------------------------------ walidatory ----


def _digits(s):
    return [int(c) for c in s if c.isdigit()]


def _weighted(d, weights):
    return sum(a * b for a, b in zip(d, weights))


def pesel_ok(s):
    d = _digits(s)
    return len(d) == 11 and (10 - _weighted(d[:10], [1, 3, 7, 9] * 3) % 10) % 10 == d[10]


def nip_ok(s):
    d = _digits(s)
    return len(d) == 10 and _weighted(d, [6, 5, 7, 2, 3, 4, 5, 6, 7]) % 11 == d[9]


def regon_ok(s):
    d = _digits(s)
    if len(d) == 9:
        w = [8, 9, 2, 3, 4, 5, 6, 7]
    elif len(d) == 14:
        w = [2, 4, 8, 5, 0, 9, 7, 3, 6, 1, 2, 4, 8]
    else:
        return False
    return _weighted(d, w) % 11 % 10 == d[-1]


def dowod_ok(s):
    s = s.replace(" ", "").upper()
    vals = [ord(c) - 55 if c.isalpha() else int(c) for c in s]
    return len(vals) == 9 and _weighted(vals, [7, 3, 1, 9, 7, 3, 1, 7, 3]) % 10 == 0


def iban_ok(s):
    s = re.sub(r"\s", "", s).upper()
    if not s.startswith("PL"):
        s = "PL" + s
    if len(s) != 28:
        return False
    num = "".join(str(int(c, 36)) for c in s[4:] + s[:4])
    return int(num) % 97 == 1


# Kolejnosc = priorytet: dluzsze/pewniejsze wzorce zajmuja tekst pierwsze.
# (nazwa, regex, walidator lub None)
DETECTORS = [
    ("IBAN", r"(?<![\w])(?:PL ?)?\d{2}(?: ?\d{4}){6}(?!\d)", iban_ok),
    ("REGON", r"(?<![\d-])\d{14}(?![\d-])", regon_ok),
    ("PESEL", r"(?<![\d-])\d{11}(?![\d-])", pesel_ok),
    ("NIP", r"(?<![\w-])(?:PL ?)?(?:\d{3}-?\d{3}-?\d{2}-?\d{2}|\d{3}-\d{2}-\d{2}-\d{3})(?![\d-])",
     nip_ok),
    ("REGON", r"(?<![\d-])\d{9}(?![\d-])", regon_ok),
    ("Dowód", r"(?<![\w])[A-Z]{3} ?\d{6}(?!\d)", dowod_ok),
    ("E-mail", r"[\w.+-]+@[\w-]+(?:\.[\w-]+)+", None),
    ("Telefon", r"(?<![\d+])(?:(?:\+48|0048) ?)?(?:\d{3}[ -]\d{3}[ -]\d{3}|\(?\d{2}\)?[ -]\d{3}[ -]\d{2}[ -]\d{2})(?!\d)"
     r"|(?:\+48|0048) ?\d{9}(?!\d)", None),
]
CATEGORIES = list(dict.fromkeys(name for name, _, _ in DETECTORS))


def find_spans(text, enabled=None, words=()):
    """Zwraca [(start, koniec, kategoria)] bez nakladania, posortowane."""
    taken = []

    def free(a, b):
        return all(b <= x or a >= y for x, y, _ in taken)

    for name, rx, check in DETECTORS:
        if enabled is not None and name not in enabled:
            continue
        for m in re.finditer(rx, text):
            if (check is None or check(m.group())) and free(m.start(), m.end()):
                taken.append((m.start(), m.end(), name))
    for w in sorted({w.strip() for w in words if w.strip()}, key=len, reverse=True):
        for m in re.finditer(r"(?<!\w)%s(?!\w)" % re.escape(w), text, re.I):
            if free(m.start(), m.end()):
                taken.append((m.start(), m.end(), "Lista słów"))
    return sorted(taken)


def mask(text, spans):
    """Zamienia znaki w zakresach na '*' (dlugosc tekstu bez zmian)."""
    chars = list(text)
    for a, b, _ in spans:
        for i in range(a, b):
            if not chars[i].isspace():
                chars[i] = "*"
    return "".join(chars)


# ------------------------------------------------------------------ DOCX ----


def anon_docx(src, dst, enabled, words):
    import docx
    from docx.oxml.ns import qn

    doc = docx.Document(src)
    parts = [doc.part] + [r.target_part for r in doc.part.rels.values()
                          if r.reltype.endswith(("/header", "/footer",
                                                 "/footnotes", "/endnotes"))]
    stats = Counter()
    tags = (qn("w:t"), qn("w:delText"))  # delText = usuniete w trybie sledzenia zmian
    for part in parts:
        for p in part.element.iter(qn("w:p")):
            nodes = [n for n in p.iter(*tags)]
            text = "".join(n.text or "" for n in nodes)
            spans = find_spans(text, enabled, words)
            if not spans:
                continue
            stats.update(s[2] for s in spans)
            masked, pos = mask(text, spans), 0
            # tekst akapitu bywa pociety na wiele "runow" - maska ma ta sama
            # dlugosc, wiec kazdy fragment dostaje swoj wycinek
            for n in nodes:
                ln = len(n.text or "")
                n.text = masked[pos:pos + ln]
                pos += ln
    cp = doc.core_properties
    for attr in ("author", "last_modified_by", "comments", "title",
                 "subject", "keywords", "category"):
        setattr(cp, attr, "")
    doc.save(dst)
    return stats, 0


# ------------------------------------------------------------------- PDF ----


def anon_pdf(src, dst, enabled, words):
    import pymupdf

    doc = pymupdf.open(src)
    stats, missed, has_text = Counter(), 0, False
    for page in doc:
        text = page.get_text()
        has_text = has_text or bool(text.strip())
        found = False
        for a, b, cat in find_spans(text, enabled, words):
            frag = text[a:b]
            rects = []
            # fragment przelamany miedzy liniami szukamy linia po linii
            for piece in frag.split("\n"):
                if piece.strip():
                    rects += page.search_for(piece)
            if rects:
                for r in rects:
                    page.add_redact_annot(r, fill=(0, 0, 0))
                stats[cat] += 1
                found = True
            else:
                missed += 1
        if found:
            page.apply_redactions()  # wycina tekst (i piksele obrazow) spod prostokatow
    if not has_text:
        doc.close()
        raise ValueError("PDF nie ma warstwy tekstowej (skan) - najpierw przepusc go przez OCR")
    doc.set_metadata({})
    doc.del_xml_metadata()
    doc.save(dst, garbage=4, deflate=True)
    doc.close()
    return stats, missed


# ------------------------------------------------------------------ wsad ----

HANDLERS = {".docx": anon_docx, ".pdf": anon_pdf}


def run_batch(inp, out_dir, enabled, words, log=print):
    if os.path.isfile(inp):
        files = [inp]
    else:
        files = sorted(os.path.join(inp, f) for f in os.listdir(inp)
                       if os.path.splitext(f)[1].lower() in HANDLERS)
    if not files:
        log("Brak plikow DOCX/PDF w: %s" % inp)
        return
    os.makedirs(out_dir, exist_ok=True)
    for f in files:
        name, ext = os.path.splitext(os.path.basename(f))
        dst = os.path.join(out_dir, name + "_anonim" + ext.lower())
        log("Przetwarzam: %s" % os.path.basename(f))
        try:
            stats, missed = HANDLERS[ext.lower()](f, dst, enabled, words)
            summary = ", ".join("%s: %d" % kv for kv in sorted(stats.items())) or "nic nie znaleziono"
            log("  %s" % summary)
            if missed:
                log("  UWAGA: %d znalezionych danych nie udalo sie zlokalizowac na stronie"
                    " - sprawdz plik recznie!" % missed)
            log("  OK -> %s" % dst)
        except Exception:
            log("BLAD: %s\n%s" % (os.path.basename(f), traceback.format_exc()))
    log("Zakonczono (%d plikow). Zawsze przejrzyj wynik przed publikacja." % len(files))


# -------------------------------------------------------------------- GUI ----


def gui():
    import tkinter as tk
    from tkinter import filedialog, ttk, scrolledtext

    root = tk.Tk()
    root.title("Anonimizacja RODO (DOCX / PDF)")
    root.geometry("780x620")
    pad = dict(padx=6, pady=3)

    v_in = tk.StringVar(value=os.path.join(APP_DIR, "INPUT"))
    v_out = tk.StringVar(value=os.path.join(APP_DIR, "OUTPUT"))
    v_cats = {c: tk.BooleanVar(value=True) for c in CATEGORIES}

    f = ttk.Frame(root)
    f.pack(fill="x", **pad)
    ttk.Label(f, text="Plik lub folder:").grid(row=0, column=0, sticky="w", **pad)
    ttk.Entry(f, textvariable=v_in, width=60).grid(row=0, column=1, **pad)
    ttk.Button(f, text="Plik...", command=lambda: v_in.set(filedialog.askopenfilename(
        filetypes=[("DOCX / PDF", "*.docx *.pdf")]) or v_in.get())).grid(row=0, column=2, **pad)
    ttk.Button(f, text="Folder...", command=lambda: v_in.set(
        filedialog.askdirectory() or v_in.get())).grid(row=0, column=3, **pad)
    ttk.Label(f, text="Folder wyjsciowy:").grid(row=1, column=0, sticky="w", **pad)
    ttk.Entry(f, textvariable=v_out, width=60).grid(row=1, column=1, **pad)
    ttk.Button(f, text="Wybierz...", command=lambda: v_out.set(
        filedialog.askdirectory() or v_out.get())).grid(row=1, column=2, **pad)

    g = ttk.LabelFrame(root, text="Co usuwac")
    g.pack(fill="x", **pad)
    for i, c in enumerate(CATEGORIES):
        ttk.Checkbutton(g, text=c, variable=v_cats[c]).grid(row=0, column=i, sticky="w", **pad)

    h = ttk.LabelFrame(root, text="Dodatkowe slowa do usuniecia (imiona, nazwiska, adresy) - jedno w linii")
    h.pack(fill="x", **pad)
    words_box = tk.Text(h, height=5)
    words_box.pack(fill="x", **pad)

    log_box = scrolledtext.ScrolledText(root, height=14)
    log_box.pack(fill="both", expand=True, **pad)

    def log(msg):
        def put():
            log_box.insert("end", str(msg) + "\n")
            log_box.see("end")
        root.after(0, put)

    btn = ttk.Button(root, text="Anonimizuj")
    btn.pack(pady=6)

    def start():
        inp, out = v_in.get().strip('" '), v_out.get().strip('" ')
        if not os.path.exists(inp):
            return log("Wskaz istniejacy plik lub folder.")
        if not out:
            return log("Wskaz folder wyjsciowy.")
        enabled = {c for c, v in v_cats.items() if v.get()}
        words = words_box.get("1.0", "end").splitlines()
        btn.config(state="disabled")
        log_box.delete("1.0", "end")

        def work():
            try:
                run_batch(inp, out, enabled, words, log)
            finally:
                root.after(0, lambda: btn.config(state="normal"))

        threading.Thread(target=work, daemon=True).start()

    btn.config(command=start)
    if "--selftest" in sys.argv:
        root.after(200, root.destroy)
    root.mainloop()


# --------------------------------------------------------------- selftest ----


def selftest():
    import tempfile
    import docx
    import pymupdf

    assert pesel_ok("44051401359") and not pesel_ok("44051401358")
    assert nip_ok("123-456-32-18") and not nip_ok("1234563219")
    assert regon_ok("123456785") and regon_ok("12345678512347") and not regon_ok("123456789")
    assert dowod_ok("ABA300000") and not dowod_ok("ABA300001")
    assert iban_ok("PL61 1090 1014 0000 0712 1981 2874")
    assert iban_ok("61109010140000071219812874") and not iban_ok("61109010140000071219812875")

    t = ("PESEL 44051401359, NIP 123-456-32-18, kwota 12345678901, "
         "dowod ABA 300000, tel. +48 600 700 800, mail jan.kowalski@urzad.gov.pl, "
         "konto PL61 1090 1014 0000 0712 1981 2874, Jan Kowalski.")
    spans = find_spans(t, None, ["Jan Kowalski"])
    cats = Counter(s[2] for s in spans)
    assert cats == Counter({"PESEL": 1, "NIP": 1, "Dowód": 1, "Telefon": 1,
                            "E-mail": 1, "IBAN": 1, "Lista słów": 1}), cats
    m = mask(t, spans)
    assert len(m) == len(t) and "12345678901" in m and "44051401359" not in m
    assert "Kowalski." not in m and "kowalski@" not in m
    assert [s[2] for s in find_spans(t, {"E-mail"}, ())] == ["E-mail"]

    tmp = tempfile.mkdtemp()
    # DOCX: PESEL pociety na dwa runy + stopka
    d = docx.Document()
    p = d.add_paragraph("PESEL: 4405")
    p.add_run("1401359 koniec")
    d.sections[0].footer.paragraphs[0].text = "kontakt: a@b.pl"
    d.core_properties.author = "Jan Kowalski"
    src = os.path.join(tmp, "t.docx")
    d.save(src)
    anon_docx(src, os.path.join(tmp, "o.docx"), None, ())
    o = docx.Document(os.path.join(tmp, "o.docx"))
    assert o.paragraphs[0].text == "PESEL: *********** koniec", o.paragraphs[0].text
    assert "a@b.pl" not in o.sections[0].footer.paragraphs[0].text
    assert o.core_properties.author == ""

    # PDF: tekst ma faktycznie zniknac z pliku
    pdoc = pymupdf.open()
    pdoc.new_page().insert_text((72, 72), "Wnioskodawca PESEL 44051401359 zostaje")
    pdoc.set_metadata({"author": "Jan Kowalski"})
    pdoc.save(os.path.join(tmp, "t.pdf"))
    stats, missed = anon_pdf(os.path.join(tmp, "t.pdf"), os.path.join(tmp, "o.pdf"), None, ())
    assert stats["PESEL"] == 1 and missed == 0
    out = pymupdf.open(os.path.join(tmp, "o.pdf"))
    txt = out[0].get_text()
    assert "44051401359" not in txt and "Wnioskodawca" in txt, txt
    assert not out.metadata.get("author")
    out.close()
    print("selftest OK")


if __name__ == "__main__":
    if "--selftest" in sys.argv:
        selftest()
        if "--gui" in sys.argv:
            gui()
    else:
        gui()
