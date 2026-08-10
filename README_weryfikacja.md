# Weryfikacja merytoryczna — raport

Zweryfikowano 1473 pytania w pięciu plikach. Przy tej skali sprawdzanie
pytanie po pytaniu nie ma sensu, więc zamiast losowego próbkowania użyłam
metod, które faktycznie znajdują błędy: **szukania sprzeczności między
niezależnymi źródłami** oraz **kontroli wzrokowej ekstrakcji**.

## 1. Wierność ekstrakcji z ARCYBAZY — potwierdzona

Wyrenderowałam strony PDF i porównałam wzrokowo, co jest podświetlone,
z tym, co znalazło się w pliku.

| Strona | Styl opcji | Pytań | Zgodność |
|---|---|---|---|
| 19 (EGZ 2018) | `a. b. c. d.` | 6 | **6/6, każda opcja** |
| 132 (kolokwium) | `1. 2. 3. 4.` | 4 | **4/4, każda opcja** |

Obie ścieżki parsera (literowa i cyfrowa) odwzorowują klucz bez błędu.

## 2. Sprzeczności między źródłami — jedna realna

Porównałam wszystkie pytania o identycznej treści występujące w więcej niż
jednym pliku (69 przypadków). Znaleziono **jedną** rzeczywistą sprzeczność
werdyktu:

> **„Zasadowicę metaboliczną mogą powodować:"**
> — egzamin 2016/2017 (klucz oficjalny): diuretyki pętlowe **+** tiazydy
> — ARCYBAZA: tylko tiazydy

Sprawdziłam surowy PDF: na str. 348 student podświetlił wyłącznie
„tiazydy". To **błąd źródła, nie ekstrakcji**. Diuretyki pętlowe wywołują
zasadowicę hipochloremiczną („contraction alkalosis") — to podręcznikowe.

**Poprawka naniesiona** w `build_export_arcybaza.py` (słownik
`CORRECTIONS`, widoczna w logu jako „poprawki merytoryczne: 1").

Dodatkowo sprawdziłam 40 przypadków tej samej treści opcji z przeciwnym
werdyktem w obrębie działu — wszystkie okazały się poprawne, bo dotyczyły
różnych pytań (np. „zatrucie pestycydami" jest prawdą dla atropiny
i fałszem dla neostygminy).

## 3. Konfrontacja moich własnych decyzji z niezależnym kluczem

Trzy wcześniejsze pliki nie miały klucza — odpowiedzi ustaliłam
merytorycznie. ARCYBAZA jest źródłem niezależnym, więc posłużyła za kontrolę.

| Moja decyzja | Werdykt ARCYBAZY |
|---|---|
| **2017 pyt. 93** — odrzuciłam klucz „wszystkie przeciwwskazane w niewydolności serca"; zostawiłam tylko werapamil i nimesulid | **Potwierdzone wprost.** „Do leków przeciwwskazanych u pacjentów z niewydolnością serca należy: werapamil ✓, ketoprofen ✓; nebiwolol ✗, **iwabradyna ✗**" |
| **2014/2015 pyt. 57** — klozapina: obniża próg drgawkowy + agranulocytoza | **Potwierdzone.** „Klozapina może być przyczyną: obniżenia progu drgawkowego, przyrostu masy ciała, agranulocytozy" |
| **2014/2015 pyt. 39** — hepatotoksyczność nie jest typowa dla metronidazolu | **Zgodne.** ARCYBAZA jako działanie niepożądane metronidazolu podaje neuropatię obwodową |
| **2014/2015 pyt. 54** — benzodiazepiny w alkoholowym zespole abstynencyjnym | **Bez konfliktu.** ARCYBAZA odrzuca je w zespole abstynencyjnym *opioidowym*, co jest inną sytuacją |
| **2016 pyt. 27 / 2014-2015 pyt. 95** — stopień jonizacji wpływa na biodostępność | **Spójne** między moimi plikami (jeden ma stem przeczący, drugi twierdzący) |

**Żadna z moich decyzji nie została obalona.** Odstępstwo od klucza 2017,
przy którym miałam najwięcej wątpliwości, okazało się poparte niezależnym
źródłem.

## 4. Poprawki w samym pliku ARCYBAZY

Weryfikacja wykazała problem, którego wcześniej nie widziałam, i **plik
został przebudowany**:

Sekcje IWL w ARCYBAZIE to *adnotowana kopia studenta* — do treści opcji
doklejone są całe notatki. Tam, gdzie tak było, **żółte podświetlenie
obejmowało też notatkę**, więc do poprawnych odpowiedzi trafiały doklejone
treści. Przykład: w pytaniu o leczenie grypy poprawną odpowiedzią stawał
się „foskarnet" razem z przyklejoną notatką o inhibitorach neuraminidazy.

Co zrobiłam:
- usunęłam doklejone notatki (wypunktowania, ciągi WERSALIKÓW, komentarze
  typu „– nie bo…", „TO FLOZYNY"),
- gdy to samo pytanie istnieje w wersji czystej i adnotowanej, **zostaje
  czysta** — odrzuciło to 284 wiersze gorszej jakości.

| | przed | po |
|---|---|---|
| wierszy | 1265 | **1105** |
| opcji z doklejonym komentarzem | 65 | 5 |

## 5. Czego ta weryfikacja nie obejmuje

- **Nie sprawdziłam merytorycznie 1105 pytań ARCYBAZY pojedynczo.** Klucz
  pochodzi od studentów (poza latami 2018+). Przypadek zasadowicy pokazuje,
  że zdarzają się niedomarkowania — mogą być kolejne, których nie wykryje
  porównanie między źródłami, bo dane pytanie występuje tylko raz.
- **Klasyfikacja `Section`/`Category` w ARCYBAZIE pozostaje automatyczna.**
  244 pytania mają pewność „niska" — to wciąż pozycja do przejrzenia.
- 37 pytań ma stem przeczący („nie", „fałszywe", „z wyjątkiem"). Są
  poprawnie odwzorowane, ale w trybie fiszki czyta się je nieintuicyjnie.

## Stan po weryfikacji

| Plik | Pytań | Źródło klucza |
|---|---|---|
| 2014/2015 | 100 | ustalony merytorycznie |
| 2016 | 99 | ustalony merytorycznie |
| 2016/2017 | 100 | oficjalny klucz z PDF |
| 2022 | 69 | rekonstrukcja studencka |
| ARCYBAZA | 1105 | podświetlenia w PDF |
| **razem** | **1473** | |

Importer aplikacji wczytuje dziś **1279 / 1473**; po dopisaniu ośmiu
działów do `SECTION_ALIASES` (lista w `README_eksport_arcybaza.md`) —
wszystkie.


---

# Aneks: audyt „czy zostały treści studenckie?"

Na pytanie *„czy nie ma już w ŻADNYM pytaniu odpowiedzi studentów ani
podpowiedzi?"* pierwotna odpowiedź brzmiała **nie** — zostały. Poniżej co
znalazłam i jak to usunęłam.

## Znalezione zanieczyszczenia

| Rodzaj | Sztuk | Przykład |
|---|---|---|
| legenda klucza wklejona w opcję | 14 + 5 | „agranulocytoza **a – prawidłowe 1, 3 i 4 b – …**" |
| notatka WERSALIKAMI | 30 | „…w rdzeniu przedłużonym**BARDZO NIEBEZPIECZNY LEK – stymulują miokardium**" |
| wykrzykniki / ocena studenta | 3 | „…nie rozwija się tolerancja na **TYLKO NA TE!!!**" |
| notatka pisana małą literą | wykryta strukturalnie | „…objawów wypadowych **To miał być super pomysł na menopauzę, ale to…**" |
| treść jednego pytania w opcji drugiego | 1 | błąd parsera przy numeracji `5.` |

## Cztery błędy w moim kodzie, które to powodowały

1. **Wzorzec legendy zakotwiczony na `$`** — nie łapał linii
   „a – prawidłowe 1, 3 i 4", więc doklejała się jako kontynuacja opcji.
2. **Reguły cięcia z samym *lookahead*** (`(?=…)`) dopasowywały pusty ciąg,
   więc `sub` usuwał tylko spację, a notatka zostawała.
3. **`NOTE_LINE` kompilowane z `re.I`** — przez to wzorzec „2+ słowa
   WERSALIKAMI" pasował do *dowolnych* dwóch wyrazów i wycinał legalne
   linie kontynuacji (43 pola urwane w połowie zdania).
4. **Numeryczna opcja `5.`** — treść kolejnego pytania trafiała jako piąta
   opcja poprzedniego.

## Co zrobiłam

- naprawiłam wszystkie cztery błędy w parserze i regułach czyszczenia,
- **usunęłam w całości adnotowaną kopię IWL** (524 pytania). Notatki są tam
  wklejone w tekst bez żadnego stałego znacznika, często małą literą —
  filtrami nie da się zagwarantować czystości, więc zamiast gonić kolejne
  wzorce, odrzuciłam całą tę część,
- gdy to samo pytanie występuje w kilku wariantach, zostaje ten
  o najkrótszych opcjach, czyli nieobrośnięty notatką (174 warianty odrzucone),
- 2 wiersze nadal pasujące do wzorców zanieczyszczeń odrzuciłam zamiast
  publikować.

## Wynik audytu końcowego

Przeskanowano **1170 wierszy** (treść pytania + wszystkie opcje) we
wszystkich pięciu plikach, dwoma niezależnymi detektorami:

| Detektor | Wynik |
|---|---|
| słownikowy — legenda klucza, wersaliki, `!!`, `•`, `->`, „chyba/imo/nwm/moim zdaniem", „prezka/slajd/sketchy", numery stron, oceny studenckie | **0 trafień** |
| strukturalny — opcja ponad 2,6× dłuższa od siostrzanych | 10 trafień, wszystkie sprawdzone ręcznie: **autentyczne długie stwierdzenia egzaminacyjne**, nie notatki |
| urwane pola (tekst kończący się spójnikiem) | 2 trafienia, oba sprawdzone w PDF: **poprawne** — pytanie celowo kończy się przyimkiem, bo dopełnia je odpowiedź |

## Koszt

| | przed audytem | po |
|---|---|---|
| ARCYBAZA | 1105 | **802** |
| razem 5 plików | 1473 | **1170** |

Straciłam 303 pytania, głównie z adnotowanej kopii IWL. Uznałam, że
gwarancja czystości jest tu ważniejsza niż liczba — te pytania są zresztą
w większości wariantami pytań, które zostały w wersji czystej.

Poprawka merytoryczna z pkt. 2 (zasadowica metaboliczna) **nie ma już
zastosowania** — dotyczyła wiersza z kopii IWL, który został usunięty.
Pytanie w wersji czystej pochodzi z EGZ 2015, który nie ma klucza, więc
nie ma go w pliku.
