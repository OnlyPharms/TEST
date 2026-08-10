# „Baza – Egzamin – 2022"

Eksport 69 pytań do formatu Only Pharms.

| Plik | Do czego |
|---|---|
| `farmakologia_egzamin_2022_OnlyPharms.xlsx` | do przeglądania / poprawiania |
| `farmakologia_egzamin_2022_OnlyPharms.csv` | do wgrania w panelu administratora |
| `build_export_2022.py` | skrypt generujący oba pliki |

## To źródło jest innego rodzaju niż trzy poprzednie

Poprzednie pliki były oficjalnymi książeczkami egzaminacyjnymi. Ten jest
**studencką rekonstrukcją pisaną wspólnie po egzaminie** — i to widać:

- **brak klucza odpowiedzi**,
- 14 × „chyba", 34 × „?", 7 × „+1", 2 × „nie pamiętam", 1 × „nwm",
- 16 pustych opcji (same numery bez treści),
- wtrącone dyskusje między studentami wewnątrz treści pytań
  („ktoś coś?", „moim zdaniem nie", „dzisiaj byłam na wglądzie i Panie
  powiedziały…"),
- pomieszane dwie wersje egzaminu, miejscami sprzeczne co do kolejności
  i treści odpowiedzi,
- na początku żartobliwe pytanie („czy paź ma dobre serduszko?").

**Wniosek praktyczny:** ten plik ma niższą wiarygodność niż trzy poprzednie
i warto przejrzeć go przed wgraniem. Cała treść pytań została oczyszczona
z komentarzy studenckich, a wszystkie odpowiedzi ustalone merytorycznie.

## Co pominięto i dlaczego

Z 78 ponumerowanych pozycji (plus żartobliwa na początku) wyeksportowano
**69**. Pominięto 9:

| Pyt. | Powód |
|---|---|
| 12 | opcja 3 nieczytelna („E…peg alfa") |
| 13 | powtórzenie pyt. 7; opcja 2 sporna (papaweryna czy racekadotryl) |
| 28 | brak opcji 3 i 4 |
| 64 | brak opcji 2 i 3 |
| 66 | brak opcji 1 i 3 |
| 69 | brak opcji 2, a opcje 3 i 4 merytorycznie sporne |
| 74 | brak opcji 1 i 3, a obie znane opcje są fałszywe — zero `True` |
| 75 | znana tylko opcja 1 |
| 78 | brak opcji 4; pytanie spoza farmakologii (EBM / prawo farmaceutyczne) |

Pominięto też powtórzenia z końca dokumentu („2 wersja" pyt. 22, pyt. 79
i nienumerowane powtórzenie pytania o IPP) — to inne zapisy pytań już
uwzględnionych.

Cztery pytania (**54, 68, 71, 73**) mają w źródle brakującą jedną opcję,
ale pozostałe trzy są kompletne i jednoznaczne — zostały wyeksportowane
z trzema opcjami.

## Nowy dział do dopisania

58 z 69 pytań zaimportuje się od razu. Dochodzi **jeden nowy dział** —
pierwszy od trzech plików:

| Section | Klucz do `SECTION_ALIASES` | Pytania |
|---|---|---|
| LEKI PRZECIWNOWOTWOROWE | `"LEKI PRZECIWNOWOTWOROWE"` | 2, 57 |

Pozostałe 9 pytań używa działów zgłoszonych wcześniej: `CHOROBA
NIEDOKRWIENNA SERCA` (9, 19, 38), `LEKI HIPOLIPEMIZUJĄCE` (22, 40, 65),
`ZABURZENIA RYTMU SERCA` (21, 72), `LEKI PRZECIWWIRUSOWE` (36).

Ten plik jako pierwszy wykorzystuje istniejące działy `OSTEOPOROZA`
(pyt. 39) i `OTĘPIENIE` (pyt. 68) oraz kategorię `Monitorowanie` (pyt. 51).

## Pytania, w których źródło samo się spiera

Studenci nie doszli do zgody — moje rozstrzygnięcia:

- **5** — bupropion oznaczony jako fałsz; to inhibitor wychwytu
  noradrenaliny i dopaminy, nie serotoniny.
- **6** — atropina we wstrząsie anafilaktycznym oznaczona jako fałsz
  (lekiem jest adrenalina). Zgodne z wnioskiem części grupy („odp. 1,3").
- **8** — przedawkowanie oksykodonu oznaczone jako fałsz: opioidy dają
  miozę i bradykardię, ale nie ślinotok, który wskazuje na mechanizm
  muskarynowy.
- **14** — użyto pierwszej zapisanej wersji opcji (furosemid,
  sulfasalazyna, adapalen, fluorochinolony); druga wersja podawała
  torasemid i wyciąg z dziurawca. Wszystkie sześć są fototoksyczne.
- **19** — przyjęto, że mechanizm w miażdżycowej ChNS polega na
  zmniejszeniu naprężenia mięśniówki komór; do tego wniosku doszła też
  grupa po dyskusji.
- **37** — połączenie wankomycyna + gentamycyna uznane za interakcję
  (sumowanie nefrotoksyczności to interakcja farmakodynamiczna).
- **44** — digoksyna **nie** działa hipotensyjnie; zgodne z odpowiedzią,
  którą studentka uzyskała na wglądzie do prac.
- **56** — wzrost masy ciała oznaczony jako fałsz: odpowiadają za niego
  głównie receptory H1 i 5-HT2C, nie D2.

## Pytania wątpliwe mimo rozstrzygnięcia

Warte rzutu okiem — tu nie mam pewności:

- **26** (profilaktyka okołooperacyjna) — jako prawdziwe zostawiono tylko
  „dopasowana indywidualnie"; opcje o czasie podania i dawce są dyskusyjne.
- **43** — makrolid i doksycyklina oznaczone jako fałsz przy etiologii
  paciorkowcowej; zależy to od przyjętych rekomendacji.
- **51** — agonista GPIIb/IIIa oznaczony jako fałsz (monitoruje się przy
  nim płytki, ale nie samą terapię). Źródło odnotowuje, że pytanie zostało
  zgłoszone przez studentów jako wadliwe.
- **52** — opcja 2 miała w źródle dwa różne brzmienia; użyto pierwszego.
- **59** — żaden z leków nie spełnia wszystkich trzech warunków naraz;
  wybrano oseltamiwir i metronidazol jako najbliższe.

## Powtórzenia treści z wcześniejszymi plikami

Cztery pytania mają stem identyczny z wcześniejszymi („Wskaż prawidłowe
połączenie lek – wskazanie / mechanizm działania / działanie niepożądane",
„Działanie przeciwkaszlowe wykazuje"), ale **inne** odpowiedzi — to osobne
pytania. Pyt. 70 jest bardzo zbliżone do pyt. 17 z 2016 (butamirat zamiast
pentoksyweryny) — warto zostawić oba albo scalić.

## Stan całości

| Plik | Pytań |
|---|---|
| 2014/2015 | 100 |
| 2016 | 99 |
| 2016/2017 | 100 |
| 2022 | 69 |
| **razem** | **368** |

Działy do dopisania w aplikacji (łącznie dla wszystkich czterech plików):
`CHOROBA NIEDOKRWIENNA SERCA`, `ZABURZENIA RYTMU SERCA`,
`LEKI PRZECIWWIRUSOWE`, `LEKI HIPOLIPEMIZUJACE`, `JASKRA`,
`LEKI PRZECIWNOWOTWOROWE`.

Źródło: dokument „Baza – Egzamin – 2022", opracowanie studenckie.
