# -*- coding: utf-8 -*-
"""Ekstrakcja ARCYBAZY: tekst + podswietlenia (zielone = poprawna odpowiedz)."""
import ctypes, json, re
import pypdfium2 as pdfium
import pypdfium2.raw as pr

PDF = 'arcy/ARCYBAZA FARMA Z ODP.pdf'
COLORS = {(0, 255, 0): 'G', (255, 255, 0): 'Y', (0, 255, 255): 'C'}


def page_rects(pg):
    out = []
    for i in range(pr.FPDFPage_CountObjects(pg.raw)):
        o = pr.FPDFPage_GetObject(pg.raw, i)
        if pr.FPDFPageObj_GetType(o) != 2:
            continue
        r = ctypes.c_uint(); g = ctypes.c_uint(); b = ctypes.c_uint(); a = ctypes.c_uint()
        if not pr.FPDFPageObj_GetFillColor(o, *map(ctypes.byref, (r, g, b, a))):
            continue
        tag = COLORS.get((r.value, g.value, b.value))
        if not tag:
            continue
        L = ctypes.c_float(); B = ctypes.c_float(); R = ctypes.c_float(); T = ctypes.c_float()
        if pr.FPDFPageObj_GetBounds(o, *map(ctypes.byref, (L, B, R, T))):
            out.append((tag, L.value, B.value, R.value, T.value))
    return out


def run():
    pdf = pdfium.PdfDocument(PDF)
    lines = []
    for p in range(len(pdf)):
        pg = pdf[p]
        rects = page_rects(pg)
        tp = pg.get_textpage()
        n = pr.FPDFText_CountChars(tp.raw)
        cur, curtags, prev_y = [], [], None
        for i in range(n):
            ch = tp.get_text_range(i, 1)
            L = ctypes.c_double(); R = ctypes.c_double()
            B = ctypes.c_double(); T = ctypes.c_double()
            pr.FPDFText_GetCharBox(tp.raw, i, *map(ctypes.byref, (L, R, B, T)))
            cx, cy = (L.value + R.value) / 2, (B.value + T.value) / 2
            tag = ''
            for t, l, b, r, tt in rects:
                if l - 1 <= cx <= r + 1 and b - 1 <= cy <= tt + 1:
                    tag = t
                    break
            if ch == '\r':
                continue
            if ch == '\n':
                if cur:
                    lines.append((p + 1, ''.join(cur), ''.join(curtags)))
                cur, curtags = [], []
                continue
            cur.append(ch)
            curtags.append(tag if ch.strip() else ' ')
        if cur:
            lines.append((p + 1, ''.join(cur), ''.join(curtags)))
    return lines


if __name__ == '__main__':
    lines = run()
    with open('arcy_lines.json', 'w') as f:
        json.dump(lines, f, ensure_ascii=False)
    # podglad: linia + dominujacy tag
    with open('arcy_tagged.txt', 'w') as f:
        for p, txt, tags in lines:
            g = tags.count('G'); y = tags.count('Y'); c = tags.count('C')
            m = ''
            if g > max(y, c) and g > 2: m = '[ZIELONY] '
            elif y > max(g, c) and y > 2: m = '[ZOLTY] '
            elif c > 2: m = '[CYAN] '
            f.write(f'{p:4d} {m}{txt}\n')
    print('linii:', len(lines))
