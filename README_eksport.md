# Farmakologia — egzamin WUM 2014/2015 (Stomatologia, rok III, wersja 1)

Eksport 100 pytań do formatu Only Pharms.

## Pliki

| Plik | Do czego |
|---|---|
| `farmakologia_egzamin_2014_2015_OnlyPharms.xlsx` | plik do przeglądania / poprawiania |
| `farmakologia_egzamin_2014_2015_OnlyPharms.csv` | plik do wgrania w panelu administratora |
| `build_export.py` | skrypt, który generuje oba pliki z danych źródłowych |

## Układ kolumn

Zgodny z `JAK_KATEGORYZOWAC.txt` i z plikami `questions/*.csv`:

```
Section | Category | Poziom | Type | Question | True 1..5 | False 1..5
```

`Poziom` jest pusty dla wszystkich pytań — egzamin był dla kierunku
stomatologicznego, więc pytania są dostępne dla wszystkich.

`Type` = `testowe` dla wszystkich 100 pytań.

## Jak odwzorowano pytania typu K (pyt. 1–90)

Egzamin ma 90 pytań typu K (treść + 4 ponumerowane stwierdzenia + klucz
A–E: „Prawidłowe 1, 3 i 4" itd.) oraz 10 pytań jednokrotnego wyboru
(pyt. 91–100).

Pytania typu K rozłożono na format Only Pharms bez zmiany treści:

- **Question** = treść pytania (stem) w oryginalnym brzmieniu,
- **True 1..n** = te ponumerowane stwierdzenia, które są prawdziwe,
- **False 1..n** = te ponumerowane stwierdzenia, które są fałszywe
  (czyli oryginalne dystraktory z egzaminu).

Litera klucza (A–E) nie jest potrzebna: przy `Type=testowe` i więcej niż
jednym `True` aplikacja sama włącza tryb wielokrotnego wyboru i pokazuje
wszystkie stwierdzenia jako osobne opcje — student zaznacza prawdziwe.
Odwzorowuje to oryginalne pytanie 1:1.

Pytania 91–100: **True 1** = poprawna odpowiedź, **False 1..4** =
pozostałe opcje. Opcje meta („żadna z powyższych", „prawidłowe A i B",
„wszystkie powyższe") rozwinięto na konkretne treści.

## Dystraktory dopisane

Dla 29 pytań, w których **wszystkie podane stwierdzenia są prawdziwe**
(klucz E / „wszystkie powyższe"), egzamin nie dostarcza żadnego dystraktora. Bez `False`
aplikacja nie może zbudować pytania ABCD (`index.html`, `canAbcd`),
więc do tych pytań dopisano po 3 wiarygodne farmakologicznie dystraktory.
Wszystkie pozostałe `False` pochodzą wprost z egzaminu.

## Nowe działy do dopisania w aplikacji

93 ze 100 pytań zaimportuje się od razu. Pozostałe 7 używa działów,
których nie ma jeszcze w `SECTION_ALIASES` w `server.py` — bez
dopisania ich wiersze zostaną **po cichu pominięte** przy imporcie:

| Section w pliku | Klucz do `SECTION_ALIASES` | Pytania |
|---|---|---|
| CHOROBA NIEDOKRWIENNA SERCA | `"CHOROBA NIEDOKRWIENNA SERCA"` | 10, 88 |
| ZABURZENIA RYTMU SERCA | `"ZABURZENIA RYTMU SERCA"` | 52, 56 |
| LEKI PRZECIWWIRUSOWE | `"LEKI PRZECIWWIRUSOWE"` | 72 |
| LEKI HIPOLIPEMIZUJĄCE | `"LEKI HIPOLIPEMIZUJACE"` | 74 |
| JASKRA | `"JASKRA"` | 79 |

(Klucz jest bez polskich znaków — `norm_section` usuwa diakrytyki.)

Alternatywnie można te 7 pytań przepiąć do istniejących działów.

## Poprawki literówek z oryginału

- pyt. 42: „glibenuryd" → **glibenklamid**, „pioglitazom" → **pioglitazon**
- pyt. 45: „aztreonom" → **aztreonam** (pyt. 1)
- pyt. 33: „Clostridium difficilae" → **Clostridium difficile**
- pyt. 41: „gwajafenzyna" → **gwajafenezyna**
- pyt. 50: „beta-adrenoliyków" → **beta-adrenolityków**
- pyt. 80: „Neostygomina" → **Neostygmina**
- pyt. 81: „zespół Stevena-Johnsona" → **zespół Stevensa-Johnsona**
- pyt. 48: w oryginale dwie pozycje miały numer 3 (makrogol, fosforan sodu)

## Pytania do weryfikacji merytorycznej

Egzamin nie zawiera klucza odpowiedzi — poprawne odpowiedzi ustalono
merytorycznie. Poniższe wymagają rzutu okiem przed wgraniem:

- **22** (efekt pierwszego przejścia) — przyjęto, że zależy od pasażu,
  eliminacji i metabolizmu wątrobowego, a nie od szybkości wchłaniania.
- **34** — „kotrimoksazol – zapalenie prostaty" oraz „fosfomycyna – ZUM"
  oba prawdziwe; klucz A–E nie przewiduje kombinacji „3 i 4".
- **39** — „uszkodzenie wątroby" oznaczono jako fałsz (dla metronidazolu
  charakterystyczne są metaliczny smak i reakcja disulfiramowa).
- **57** (klozapina) — prawdziwe są stwierdzenia 3 i 4; klucz A–E nie
  przewiduje kombinacji „3 i 4".
- **69** — jako selektywny inhibitor COX-2 uznano wyłącznie celekoksyb;
  nimesulid potraktowano jako inhibitor preferencyjny.
- **73** — pirymetaminę uznano za niestosowaną w profilaktyce malarii.
- **91** — osteoporoza i łysienie opisywane są dla heparyny
  niefrakcjonowanej; jako odpowiedź przyjęto krwawienie.
- **96** (oparzenia) — brak w opcjach sulfadiazyny srebra; przyjęto
  jodowany poliwidon.

Źródło: Warszawski Uniwersytet Medyczny, Katedra i Zakład Farmakologii
Doświadczalnej i Klinicznej, rok akademicki 2014/2015.
