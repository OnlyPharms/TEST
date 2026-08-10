# -*- coding: utf-8 -*-
"""
Buduje plik eksportowy Only Pharms z dokumentu "ARCYBAZA FARMA Z ODP".

Zrodlo: 367-stronicowa skladanka egzaminow i kolokwiow z farmakologii
(kierunek lekarski / IWL), wersja 29.05.2020. Poprawne odpowiedzi sa
w PDF oznaczone KOLOREM PODSWIETLENIA, nie tekstem - odczytane z obiektow
graficznych PDF (zielony 0,255,0 oraz zolty 255,255,0).

Pipeline (skrypty robocze w scratchpadzie):
  extract_arcy.py  - tekst + podswietlenia per znak
  parse2.py        - sekcje / pytania / opcje (dwa style: a.b.c.d oraz 1.2.3.4)
  classify.py      - automatyczne Section/Category ze slownikow JAK_KATEGORYZOWAC

Uklad kolumn: Section, Category, Poziom, Type, Question, True 1..5, False 1..5
plus dwie kolumny informacyjne (Zrodlo, Pewnosc), ktore importer ignoruje.
"""
import csv, json, os, re, sys, unicodedata
from collections import Counter
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, PatternFill
from openpyxl.utils import get_column_letter

SCRATCH = ('/tmp/claude-0/-home-user-TEST/'
           '818078ff-c57a-51d2-987e-4f2584083ea6/scratchpad')
sys.path.insert(0, SCRATCH)
import classify as C  # noqa: E402

MAX_T, MAX_F = 5, 5
HEADER = (["Section", "Category", "Poziom", "Type", "Question"]
          + [f"True {i}" for i in range(1, MAX_T + 1)]
          + [f"False {i}" for i in range(1, MAX_F + 1)]
          + ["Zrodlo", "Pewnosc"])



# --- czyszczenie notatek studenckich doklejonych do tresci opcji ---
CUT = [
    re.compile(r'\s*[•▪]\s.*$', re.S),                        # lista wypunktowana
    re.compile(r'\s*[-–]\s*(nie |tak |chyba |bo |raczej |imo|moim zdaniem|wg mnie)\b.*$',
               re.I | re.S),
    re.compile(r'\s*\?\s*[-–].*$', re.S),
    re.compile(r'\s+TO\s+[A-ZĄĆĘŁŃÓŚŹŻ]{4,}.*$', re.S),
    # notatka WERSALIKAMI po spacji  ("... iv TYLKO NA TE")
    re.compile(r'\s+(?=(?:[A-ZĄĆĘŁŃÓŚŹŻ]{3,}[\s,:;.!?–-]+){1,}[A-ZĄĆĘŁŃÓŚŹŻ]{3,}).*$', re.S),
    # notatka WERSALIKAMI sklejona bez spacji ("...skurczowejKORZYSTNA ROKOWNICZO")
    re.compile(r'(?<=[a-ząćęłńóśźż])(?=[A-ZĄĆĘŁŃÓŚŹŻ]{3,}(?:[\s,:;.!?–-]+[A-ZĄĆĘŁŃÓŚŹŻ]{3,}|[A-ZĄĆĘŁŃÓŚŹŻ]{3,})).*$', re.S),
]

# wzorce, ktorych obecnosc dyskwalifikuje wiersz (zostaje odrzucony, nie publikowany)
DIRTY = re.compile(
    r'[a-e]\s*[–-]\s*praw(idłow|dziw)|(?:[A-ZĄĆĘŁŃÓŚŹŻ]{4,}[\s,:;–-]+){2,}|!!|\?\?'
    r'|[•▪]|-\s*>|–\s*>'
    r'|\b(chyba|imo|nwm|moim zdaniem|wg mnie|nie pamiętam|prezka|slajd|sketchy)\b')


# punktowe poprawki tekstu wykryte podczas weryfikacji wzrokowej
TEXT_FIXES = [
    # numer strony wkleil sie w tresc opcji (str. 60 PDF, pyt. 65)
    (re.compile(r'(barierę krew- mózg)\s+22$'), r'\1'),
]


def strip_note(t):
    for rx in CUT:
        t = rx.sub('', t)
    for rx, rep in TEXT_FIXES:
        t = rx.sub(rep, t)
    return re.sub(r'\s+', ' ', t).strip(' .;,–-')


def norm(s):
    return re.sub(r'[^a-z0-9]', '',
                  unicodedata.normalize('NFKD', s.lower())
                  .encode('ascii', 'ignore').decode())



# --- poprawki merytoryczne naniesione po weryfikacji krzyzowej ---
# (tresc pytania -> opcje, ktore musza byc oznaczone jako poprawne)
CORRECTIONS = {
    # Student podswietlil tylko "tiazydy". Diuretyki petlowe wywoluja zasadowice
    # hipochloremiczna - potwierdza to oficjalny klucz egzaminu 2016/2017.
    'zasadowicemetabolicznamogapowodowac': ['diuretyki pętlowe'],
    # Liraglutyd (agonista GLP-1) jest podawany podskornie w cukrzycy typu 2 -
    # inny wariant tego pytania w ARCYBAZIE zaznacza go jako poprawny.
    # Potwierdzone przez uzytkowniczke.
    'terapiacukrzycytypu2podskornie': ['Liraglutyd'],
}


def apply_corrections(q):
    want = CORRECTIONS.get(norm(q['text']))
    if not want:
        return 0
    hit = 0
    for o in q['opts']:
        if any(norm(w) == norm(o['text']) for w in want) and o['tag'] not in ('G', 'Y'):
            o['tag'] = 'G'
            hit += 1
    return hit


def collect():
    secs = json.load(open(os.path.join(SCRATCH, 'arcy_parsed2.json')))
    cand, stats = [], Counter()
    for s in secs:
        for q in s['qs']:
            stats['wszystkie'] += 1
            if not any(o['tag'] in ('G', 'Y') for o in q['opts']):
                stats['bez klucza'] += 1
                continue
            for o in q['opts']:
                o['text'] = strip_note(o['text'])
            q['opts'] = [o for o in q['opts'] if o['text']]
            if not q['opts'] or not any(o['tag'] in ('G', 'Y') for o in q['opts']):
                stats['bez klucza'] += 1
                continue
            q['src'] = s['name']
            q['annot'] = s['name'].startswith('EGZAMIN IWL')
            cand.append(q)

    # Sekcje IWL to robocza, adnotowana kopia studencka: notatki sa wklejone
    # w tresc opcji i pytan, czesto mala litera, bez zadnego stalego znacznika.
    # Nie da sie zagwarantowac ich czystosci filtrami, wiec odrzucamy je w calosci.
    kept = []
    for q in cand:
        if q['annot']:
            stats['kopia adnotowana IWL'] += 1
            continue
        kept.append(q)

    # Ten sam stem bywa w dokumencie w kilku wariantach: jeden czysty, drugi
    # z notatka doklejona do opcji. ALE ten sam stem ("Digoksyna:", "Wskaz
    # prawdziwe polaczenie lek-wskazanie:") miewaja tez ROZNE pytania z roznych
    # lat. Scalamy wiec tylko wtedy, gdy zestawy opcji faktycznie sie pokrywaja
    # (podobienstwo Jaccarda >= 0.5); inaczej zostawiamy oba pytania.
    by_stem = {}
    for q in kept:
        by_stem.setdefault(norm(q['text'])[:80], []).append(q)

    def same_question(a, b):
        """Warianty tego samego pytania roznia sie tylko doklejona notatka:
        tyle samo opcji, a kazda para to identyczne teksty albo jeden jest
        prefiksem drugiego. Rozne pytania o tym samym stemie NIE sa scalane."""
        if len(a['opts']) != len(b['opts']):
            return False
        used = set()
        for oa in a['opts']:
            ta = norm(oa['text'])
            for j, ob in enumerate(b['opts']):
                if j in used:
                    continue
                tb = norm(ob['text'])
                if ta == tb or tb.startswith(ta) or ta.startswith(tb):
                    used.add(j)
                    break
            else:
                return False
        return True

    finals = []
    for group in by_stem.values():
        clusters = []
        for q in group:
            for cl in clusters:
                if same_question(q, cl[0]):
                    cl.append(q)
                    break
            else:
                clusters.append([q])
        for cl in clusters:
            cl.sort(key=lambda q: sum(len(o['text']) for o in q['opts']))
            finals.append(cl[0])
            stats['wariant z notatka'] += len(cl) - 1

    seen, out = {}, []
    for q in finals:
        key = norm(q['text']) + '|' + '|'.join(
            sorted(norm(o['text']) for o in q['opts']))
        if key in seen:
            stats['duplikaty'] += 1
            continue
        seen[key] = 1
        stats['poprawki'] += apply_corrections(q)
        out.append(q)
    return out, stats


def build():
    qs, stats = collect()
    rows, dropped = [], []
    noF = 0
    for q in qs:
        if DIRTY.search(q['text']) or any(DIRTY.search(o['text']) for o in q['opts']):
            stats['odrzucone jako zanieczyszczone'] += 1
            continue
        trues = [o['text'] for o in q['opts'] if o['tag'] in ('G', 'Y')]
        falses = [o['text'] for o in q['opts'] if o['tag'] not in ('G', 'Y')]
        sec, cat, score = C.classify(q['text'], [o['text'] for o in q['opts']])
        if sec is None:
            dropped.append(q)
            continue
        if not trues:
            dropped.append(q)
            continue
        if not falses:
            noF += 1
        trues, falses = trues[:MAX_T], falses[:MAX_F]
        conf = 'wysoka' if score >= 3 else ('srednia' if score == 2 else 'niska')
        rows.append([sec, cat, 'lekarski', 'testowe', q['text']]
                    + trues + [''] * (MAX_T - len(trues))
                    + falses + [''] * (MAX_F - len(falses))
                    + [q['src'], conf])
    return rows, stats, dropped, noF


def main():
    rows, stats, dropped, noF = build()
    base = 'farmakologia_ARCYBAZA_OnlyPharms'

    with open(base + '.csv', 'w', newline='', encoding='utf-8-sig') as f:
        w = csv.writer(f)
        w.writerow(HEADER)
        w.writerows(rows)

    wb = Workbook()
    ws = wb.active
    ws.title = 'Pytania'
    ws.append(HEADER)
    for r in rows:
        ws.append(r)
    for c in ws[1]:
        c.font = Font(bold=True, color='FFFFFF')
        c.fill = PatternFill('solid', fgColor='2F5597')
        c.alignment = Alignment(vertical='center')
    ws.freeze_panes = 'F2'
    ws.auto_filter.ref = f'A1:{get_column_letter(len(HEADER))}{len(rows) + 1}'
    for col, wdt in {'A': 30, 'B': 21, 'C': 9, 'D': 9, 'E': 60}.items():
        ws.column_dimensions[col].width = wdt
    for i in range(6, len(HEADER) - 1):
        ws.column_dimensions[get_column_letter(i)].width = 32
    ws.column_dimensions[get_column_letter(len(HEADER) - 1)].width = 26
    ws.column_dimensions[get_column_letter(len(HEADER))].width = 10
    for row in ws.iter_rows(min_row=2):
        for c in row:
            c.alignment = Alignment(vertical='top', wrap_text=True)
    wb.save(base + '.xlsx')

    print('sparsowane pytania:      ', stats['wszystkie'])
    print('  bez klucza (odrzucone):', stats['bez klucza'])
    print('  kopia adnotowana IWL (odrzucone):', stats['kopia adnotowana IWL'])
    print('  duplikaty (odrzucone): ', stats['duplikaty'])
    print('  warianty z notatka (odrzucone):', stats['wariant z notatka'])
    print('  bez dzialu (odrzucone):', len(dropped))
    print('  zanieczyszczone (odrzucone):', stats['odrzucone jako zanieczyszczone'])
    print('  poprawki merytoryczne:', stats['poprawki'])
    print('WYEKSPORTOWANO:          ', len(rows))
    print('  w tym bez dystraktorow (same poprawne):', noF)
    print()
    print('pewnosc klasyfikacji:', dict(Counter(r[-1] for r in rows)))
    print()
    for k, v in Counter(r[0] for r in rows).most_common():
        print(f'  {v:5d}  {k}')


if __name__ == '__main__':
    main()
