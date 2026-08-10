# ARCYBAZA FARMA — 868 pytań

> **Zaktualizowane po weryfikacji** (patrz `README_weryfikacja.md`): usunięto
> całą adnotowaną kopię studencką IWL i warianty pytań z doklejonymi
> notatkami. Szczegóły i wynik audytu: `README_weryfikacja.md`.

| Plik | Do czego |
|---|---|
| `farmakologia_ARCYBAZA_OnlyPharms.xlsx` | do przeglądania / poprawiania |
| `farmakologia_ARCYBAZA_OnlyPharms.csv` | do wgrania w panelu administratora |
| `build_export_arcybaza.py` | skrypt generujący oba pliki |

Źródło: 367-stronicowa składanka egzaminów i kolokwiów z farmakologii,
kierunek lekarski (IWL), wersja 29.05.2020. **`Poziom = lekarski`** dla
wszystkich wierszy.

## Klucz odpowiedzi był zapisany kolorem, nie tekstem

Okładka mówi: *„Poprawne odpowiedzi zaznaczamy na zielono"*. Tego nie widać
w zwykłej ekstrakcji tekstu — podświetlenia to obiekty graficzne PDF.
Odczytałam je bezpośrednio: zielony `(0,255,0)` w sekcjach z odpowiedziami
`a/b/c/d` i żółty `(255,255,0)` w sekcjach typu K, a następnie zmapowałam
prostokąty podświetleń na ramki poszczególnych znaków tekstu.

Dzięki temu **klucz pochodzi z dokumentu, a nie z mojej oceny** — tak jak
w pliku 2016/2017, a inaczej niż w plikach 2014/2015, 2016 i 2022.

## Format pasuje tu wyjątkowo dobrze

Pytania mają 4 opcje, z których poprawna może być dowolna liczba (1–4).
To dokładnie model Only Pharms: `True` = opcje podświetlone, `False` =
pozostałe. Rozkład: 1 poprawna – 158, 2 – 336, 3 – 213, 4 – 161.

## Bilans

| | |
|---|---|
| pytań sparsowanych | 1617 |
| odrzucone: brak klucza | 126 |
| odrzucone: adnotowana kopia IWL | 524 |
| odrzucone: warianty z notatką | 104 |
| odrzucone: zanieczyszczone | 2 |
| odrzucone: brak działu | 4 |
| odrzucone: sprzeczne warianty | 4 |
| **wyeksportowane** | **868** |

Cały egzamin **EGZ 2015 I termin (95 pytań) nie ma zaznaczonego klucza** —
zgodne z notą na okładce, że oficjalny klucz jest dopiero od 2018. To
większość odrzuconych ze 126.

Duplikaty liczone po treści pytania **razem z kompletem opcji**, więc
pytania powtórzone między latami zostały scalone do jednego wiersza.

**161 pytań ma wszystkie opcje poprawne**, czyli zero dystraktorów. Będą
działać jako fiszki i pytania na wpisywanie, ale nie w trybie ABCD
(`canAbcd` wymaga co najmniej jednego `False`). Nie dopisywałam do nich
dystraktorów — przy tej skali byłoby to 161 zmyślonych zestawów.

## Klasyfikacja jest automatyczna — to główna różnica

W poprzednich czterech plikach `Section` i `Category` przypisywałam ręcznie,
pytanie po pytaniu. Przy 868 pytaniach zrobił to klasyfikator słownikowy
(ok. 900 wzorców leków i pojęć, ograniczony do nazw dozwolonych
w `JAK_KATEGORYZOWAC.txt`).

Kolumna **`Pewnosc`** mówi, na ilu niezależnych trafieniach słownikowych
oparte jest przypisanie działu:

| Pewność | Pytań | Co to znaczy |
|---|---|---|
| wysoka | 342 | 3+ trafienia — przypisanie praktycznie pewne |
| średnia | 270 | 2 trafienia |
| niska | 256 | 1 trafienie — **te warto przejrzeć** |

Rozkład kategorii wyszedł zbliżony do plików robionych ręcznie (Wskazania
34%, Klasyfikacja 25%, Działania niepożądane 12%), ale „Klasyfikacja" jest
tu workiem na pytania o samej nazwie leku w treści („Cyklosporyna:"),
gdzie kategoria jest z natury niejednoznaczna.

Kolumny `Zrodlo` (z którego egzaminu pochodzi pytanie) i `Pewnosc` są
**ignorowane przez importer** — sprawdziłam to, uruchamiając na pliku
prawdziwą funkcję `parse_question_rows` z `server.py`. Można je zostawić
albo usunąć.

## Dwa nowe działy

Poza pięcioma zgłoszonymi wcześniej doszły dwa, bo ARCYBAZA obejmuje
tematy nieobecne w egzaminach stomatologicznych:

- **LEKI IMMUNOSUPRESYJNE** (20 pytań) — cyklosporyna, takrolimus,
  mykofenolan, przeciwciała monoklonalne, łuszczyca, RZS
- **FARMAKOLOGIA OGÓLNA I EBM** (15 pytań) — badania kliniczne, QALY,
  AOTMiT, ChPL, farmakowigilancja, terapia genowa, doping

## Do dopisania w `server.py`

Uruchomiłam prawdziwy importer na wszystkich pięciu plikach naraz.
Z 1235 wierszy wczytuje się dziś **1080**; pozostałe **155** wpada
w `if not dzial: continue` i znika bez komunikatu. Wystarczy dopisać
do `SECTION_ALIASES`:

```python
"ZABURZENIA RYTMU SERCA": "arytmie",          # 41 pytań
"LEKI HIPOLIPEMIZUJACE": "lipid",             # 41
"LEKI PRZECIWWIRUSOWE": "wirus",              # 32
"LEKI PRZECIWNOWOTWOROWE": "onko",            # 30
"CHOROBA NIEDOKRWIENNA SERCA": "chns",        # 27
"LEKI IMMUNOSUPRESYJNE": "immuno",            # 20
"FARMAKOLOGIA OGOLNA I EBM": "ogolna",        # 15
"JASKRA": "jaskra",                           # 2
```

(Klucze bez polskich znaków — `norm_section` usuwa diakrytyki. Wartości to
propozycje ID; każdy nowy dział trzeba też dodać do listy działów
w aplikacji.)

Po dopisaniu wczytuje się **1235 / 1235 pytań**.

## Czego nie zrobiłam

- Kontrola krzyżowa jest opisana w `README_weryfikacja.md`. Nie sprawdzałam
  jednak merytorycznie każdego z 868 pytań osobno — klucz pochodzi od
  studentów (poza latami 2018+, gdzie okładka deklaruje klucz oficjalny)
  i miejscami może być błędny.
- Nie usuwałam pytań powtarzających się między ARCYBAZĄ a czterema
  wcześniejszymi plikami — to inne kierunki (lekarski vs stomatologia)
  i inne brzmienia, ale część tematów się pokrywa.
- Nie dopisywałam dystraktorów do 161 pytań, które ich nie mają.

## Stan całości

| Plik | Pytań | Klucz |
|---|---|---|
| 2014/2015 | 100 | ustalony merytorycznie |
| 2016 | 98 | ustalony merytorycznie |
| 2016/2017 | 100 | oficjalny klucz z PDF |
| 2022 | 69 | rekonstrukcja studencka |
| ARCYBAZA | 868 | z podświetleń w PDF |
| **razem** | **1235** | |
