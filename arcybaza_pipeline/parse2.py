# -*- coding: utf-8 -*-
"""Parser ARCYBAZY v2 - obsluguje oba style opcji (a.b.c.d oraz 1.2.3.4)."""
import json, re

LINES = json.load(open('arcy_lines.json'))

QN = re.compile(r'^(\d{1,3})\s*[.)]\s*(.*)$')
OPTL = re.compile(r'^([a-hA-H])\s*[.)]\s*(.*)$')
SEC_TOC = re.compile(r'^(EGZ|KOLOKWIUM)\s+\d', re.I)

# smieci: numery stron, stopki, blok instrukcji, karta odpowiedzi
JUNK = re.compile(
    r'^(str\.\s*\d+|\d+|Wersja\s*\d+|POWODZENIA!?|KOD ODPOWIEDZI|'
    r'[a-e]\s*[–-]\s*(prawidłow|wszystkie)|Instrukcja|'
    r'Warszawski Uniwersytet Medyczny|Katedra i Zakład Farmakologii.*|'
    r'EGZAMIN Z FARMAKOLOGII|IWL \d ROK|Własnoręczny podpis|Numer indeksu|'
    r'[…\.\s]+|imię i nazwisko|Zaznacz prawidłową odpowiedź.*|'
    r'Poniżej znajdują się pytania.*|Prawidłowe (1|2|tylko|wszystkie).*)$', re.I)

INSTR = re.compile(r'^\d\.\s*(Kartę odpowiedzi|Numer wersji|Uwzględniane|'
                   r'Prawidłowe zaznaczenie|Zamalowany obszar|Dla każdego pytania|'
                   r'Odpowiedzi można|W przypadku błędnego|Nie należy)', re.I)

# przypisy studenckie doklejone do tresci opcji
NOTE = re.compile(r'\s*(-\s*>|–\s*>|->).*$')


def dom(tags):
    g, y, c = tags.count('G'), tags.count('Y'), tags.count('C')
    if g >= 2 and g > max(y, c):
        return 'G'
    if y >= 2 and y > max(g, c):
        return 'Y'
    if c >= 2 and c > max(g, y):
        return 'C'
    return ''


def clean(s):
    s = NOTE.sub('', s)
    s = re.sub(r'&gt;', '>', s)
    s = re.sub(r'&lt;', '<', s)
    s = re.sub(r'\s+', ' ', s).strip(' .;,')
    return s


def parse():
    secs, cur, q, opt = [], None, None, None
    exam_seen = 0

    def flush_opt():
        nonlocal opt
        if opt is not None and q is not None:
            opt['text'] = clean(opt['text'])
            if opt['text']:
                q['opts'].append(opt)
        opt = None

    def flush_q():
        nonlocal q
        flush_opt()
        if q is not None and cur is not None:
            q['text'] = clean(q['text'])
            if q['text'] and q['opts']:
                cur['qs'].append(q)
        q = None

    def newsec(name, page):
        nonlocal cur
        flush_q()
        cur = {'name': name, 'page': page, 'qs': []}
        secs.append(cur)

    for idx, (page, txt, tags) in enumerate(LINES):
        s = txt.strip()
        if page <= 2 or not s:
            continue
        if SEC_TOC.match(s) and len(s) < 60 and '...' not in s:
            newsec(s, page)
            continue
        if re.match(r'^EGZAMIN Z FARMAKOLOGII', s, re.I):
            exam_seen += 1
            # data jest 2-3 linie nizej
            nm = 'EGZAMIN IWL'
            for j in range(idx, min(idx + 5, len(LINES))):
                d = re.match(r'^\s*(\d{1,2} \w+ (20\d\d))\s*$', LINES[j][1].strip())
                if d:
                    nm = f'EGZAMIN IWL {d.group(1)}'
                    break
            newsec(nm, page)
            continue
        if cur is None or JUNK.match(s) or INSTR.match(s):
            continue

        mo = OPTL.match(s)
        mq = QN.match(s)

        if mo:
            if q is None:
                continue
            flush_opt()
            opt = {'label': mo.group(1).lower(), 'text': mo.group(2), 'tag': dom(tags)}
            continue

        if mq:
            n, rest = int(mq.group(1)), mq.group(2)
            nopts = len(q['opts']) + (1 if opt else 0) if q else 0
            has_letter = bool(q and (q['opts'] or opt) and
                              ((q['opts'][0]['label'] if q['opts'] else opt['label']).isalpha()))
            is_opt = (q is not None and 1 <= n <= 5 and n == nopts + 1
                      and nopts < 5 and not has_letter)
            if is_opt:
                flush_opt()
                opt = {'label': str(n), 'text': rest, 'tag': dom(tags)}
                continue
            flush_q()
            q = {'nr': n, 'page': page, 'text': rest, 'opts': [],
                 'style': 'num'}
            continue

        # kontynuacja
        if opt is not None:
            opt['text'] += ' ' + s
            if not opt['tag']:
                opt['tag'] = dom(tags)
        elif q is not None:
            q['text'] += ' ' + s

    flush_q()
    # ustal styl po fakcie
    for sec in secs:
        for x in sec['qs']:
            x['style'] = 'letter' if x['opts'] and x['opts'][0]['label'].isalpha() else 'num'
    return secs


if __name__ == '__main__':
    secs = parse()
    json.dump(secs, open('arcy_parsed2.json', 'w'), ensure_ascii=False)
    print(f'{"sekcja":<38}{"pyt":>5}{"4opt":>6}{"klucz":>7}{"styl":>7}')
    T = K = 0
    for s in secs:
        qs = s['qs']
        four = sum(1 for x in qs if 3 <= len(x['opts']) <= 6)
        key = sum(1 for x in qs if any(o['tag'] in ('G','Y') for o in x['opts']))
        st = 'litery' if qs and qs[0]['style'] == 'letter' else 'cyfry'
        T += len(qs); K += key
        print(f'{s["name"][:37]:<38}{len(qs):>5}{four:>6}{key:>7}{st:>7}')
    print(f'\nRAZEM pytan: {T}   z kluczem: {K}')
