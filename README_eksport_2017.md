# Farmakologia — egzamin WUM 2016/2017 (Stomatologia, rok III, wersja 1)

Eksport 100 pytań do formatu Only Pharms.

| Plik | Do czego |
|---|---|
| `farmakologia_egzamin_2017_OnlyPharms.xlsx` | do przeglądania / poprawiania |
| `farmakologia_egzamin_2017_OnlyPharms.csv` | do wgrania w panelu administratora |
| `build_export_2017.py` | skrypt generujący oba pliki |

Układ kolumn, `Type=testowe` i sposób rozłożenia pytań typu K — identyczne
jak w poprzednich plikach (opis w `README_eksport.md`).

Struktura: pyt. **1–72** typu K, pyt. **73–100** jednokrotnego wyboru.

## Ten egzamin ma klucz odpowiedzi

W przeciwieństwie do dwóch poprzednich plików PDF zawiera klucz („Odp. A"
itd.). Przypisanie True/False wynika **z klucza**, nie z mojej oceny —
poza trzema wyjątkami opisanymi niżej. Dokładność tego pliku jest więc
wyraźnie wyższa niż poprzednich.

## PDF jest wersją roboczą — rozplątane ślady redakcji

W pliku widać nieprzyjęte poprawki: stare i nowe brzmienie sklejone
w jeden ciąg, plus znaczniki `[E1]`–`[E6]`. Wszędzie przyjęto brzmienie
**nowsze** (wstawione), zgodnie z konwencją „skreślone → wstawione":

| Pyt. | W PDF | Przyjęto |
|---|---|---|
| 4 | „zwężenie źrenicy**bradykardia**" | bradykardia |
| 6 | „ketoprofen**metamizol**" | metamizol |
| 8 | „leki adsorpcyjne**leki absorpcyjne**`[E1]`" | leki adsorpcyjne |
| 10 | „**o**Olej rycynowy", „**p**Parafina płynna" | olej rycynowy, parafina płynna |
| 11 | „niklosamid **albendazol**" | albendazol |
| 15 | „rak sutka**piersi**" | rak piersi |
| 44 | „palmisetron" | palonosetron |
| 53 | „moczenie nocne**nycturia**`[E2]`" | nykturia |
| 55 | „zwiększenie liczby płytek**nadpłytkowość**`[E3]`" | nadpłytkowość |
| 58 | dwie wersje treści + `[E4]` | „profilaktyka wtórna incydentu sercowo-naczyniowego" |
| 73 | „N-acetylocysteinę **cysteinę**" | N-acetylocysteina |
| 74 | opcje B i C zawierają po dwa warianty + `[E5]` | rozdzielone |
| 100 | „tyraminę zawierających `[E6]`" | usunięto powtórzenie |

Pyt. 7 ma w kluczu zapisane „Odp. **BA**". Poprawna jest **B** (loperamid
i difenoksylat) — odpowiedź A obejmowałaby preparat z senesu, który jest
lekiem przeczyszczającym, nie zapierającym.

Poprawiono też literówki: „piroksyikamu", „chinapriylu", „praziykwantel",
„cefukroksymu", „koksycykliny", „ezetymibd", „myiastenia", „l-dL-dopaę",
„Propionibacterium acne", „Actinomyces israeli", „Pneumocystis jiroveci".

## Trzy odstępstwa od klucza egzaminu

To jedyne miejsca, gdzie **nie** poszłam za kluczem. Każde jest do
cofnięcia jedną zmianą w pliku.

- **74 (pantoprazol)** — klucz podaje „Odp. B", ale opcja B to
  „zwiększenie stężenia prostaglandyny E w żołądku". Pantoprazol hamuje
  pompę protonową = **opcja D**. To pozostałość po redakcji tego pytania
  (opcje B i C zawierają w PDF po dwa warianty tekstu).
- **92 (wydzielanie insuliny stymuluje)** — komisja wpisała „Brak
  odpowiedzi prawidłowej". Liraglutyd, jako agonista GLP-1, **pobudza
  glukozozależne wydzielanie insuliny**, więc wpisano go jako poprawną
  odpowiedź. Jeśli wolisz wierność kluczowi — usuń ten wiersz.
- **93 (przeciwwskazane w niewydolności serca)** — klucz podaje E
  („wszystkie powyższe"). Iwabradyna jest **wskazana** w niewydolności
  serca z obniżoną frakcją wyrzutową (badanie SHIFT, 2010), a amiodaron
  jest lekiem antyarytmicznym z wyboru w tej grupie chorych. Jako
  przeciwwskazane zapisano tylko **werapamil i nimesulid**. Klucz jest tu
  po prostu nieaktualny, a wgranie go uczyłoby błędu klinicznego.

## Duplikat względem wcześniejszych plików

Pytanie **2** („Zjawisko samoindukcji") to to samo pytanie co **66**
z egzaminu 2014/2015 — te same cztery stwierdzenia, różnica tylko
redakcyjna („zachodzi" / „rozwija się", „jest przyczyną" / „może być
przyczyną"). Warto zostawić jedną wersję.

Pozostałe powtórzenia treści pytań (złośliwy zespół neuroleptyczny,
hipokaliemia, „lek – mechanizm działania") mają **inne** stwierdzenia —
to osobne pytania, warto zachować oba.

## Co ten klucz potwierdza w poprzednich plikach

- **pyt. 5** potwierdza moją decyzję z 2014/2015 pyt. 69: selektywnym
  inhibitorem COX-2 jest wyłącznie celekoksyb (klucz: „Odp. D").
- **pyt. 12** potwierdza 2014/2015 pyt. 96: w oparzeniach — jodowany
  poliwidon.
- **pyt. 23** wskazuje, że w 2014/2015 pyt. 91 warto zmienić odpowiedź:
  tutaj łysienie, osteoporoza i trombocytopenia są **poprawnymi**
  działaniami niepożądanymi heparyny. Tamto pytanie dotyczyło heparyny
  drobnocząsteczkowej i miało jedną odpowiedź, ale flagowałam je jako
  wątpliwe — ten klucz to potwierdza.

## Działy

95 ze 100 pytań zaimportuje się od razu. Pozostałe 5 używa działów
zgłoszonych już wcześniej — **nie dochodzi żaden nowy**:

| Section | Pytania 2017 |
|---|---|
| CHOROBA NIEDOKRWIENNA SERCA | 58, 77 |
| LEKI HIPOLIPEMIZUJĄCE | 94, 99 |
| LEKI PRZECIWWIRUSOWE | 79 |

## Dystraktory dopisane

Dla 20 pytań, w których klucz wskazuje „Prawidłowe wszystkie (1, 2, 3 i 4)",
egzamin nie dostarcza dystraktora — dopisano po 3 wiarygodne
farmakologicznie. Wszystkie pozostałe `False` pochodzą wprost z egzaminu.

## Pytania wadliwie skonstruowane (klucz zachowany)

Poszłam za kluczem, ale warto o nich wiedzieć:

- **43** — klucz wyklucza klometiazol z leczenia alkoholowego zespołu
  abstynencyjnego, choć jest to jego klasyczne wskazanie.
- **84** — klucz uznaje za fałszywe tylko „zespół parkinsonowski", ale
  karbamazepina raczej zwiększa niż zmniejsza masę ciała, więc opcja E
  też jest dyskusyjna (pyt. 47 przyjmuje jednak to samo założenie, więc
  egzamin jest tu wewnętrznie spójny).
- **88** — poprawny linezolid, ale tygecyklina również działa na MRSA.
- **98** — poprawny imipenem, ale aztreonam także hamuje syntezę ściany
  komórkowej.
- **38** — „stosowany w krwawieniach spowodowanych niedoborem fibrynogenu"
  klucz uznaje za prawdę; to stwierdzenie budzi wątpliwości.

Źródło: Warszawski Uniwersytet Medyczny, Katedra i Zakład Farmakologii
Doświadczalnej i Klinicznej, rok akademicki 2016/2017, wersja 1.
