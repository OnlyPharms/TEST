# Farmakologia — egzamin 2016 (Stomatologia, lato)

> **Aktualizacja:** pyt. 61 usunięte jako dokładny duplikat pyt. 61
> z egzaminu 2014/2015. Plik ma 98 pytań.

Eksport 98 pytań do formatu Only Pharms.

## Pliki

| Plik | Do czego |
|---|---|
| `farmakologia_egzamin_2016_OnlyPharms.xlsx` | plik do przeglądania / poprawiania |
| `farmakologia_egzamin_2016_OnlyPharms.csv` | plik do wgrania w panelu administratora |
| `build_export_2016.py` | skrypt, który generuje oba pliki |

Układ kolumn, `Type=testowe` i sposób rozłożenia pytań typu K — identyczne
jak w pliku 2014/2015 (opis w `README_eksport.md`).

## Dlaczego 98, a nie 100

W oryginalnym PDF **nie ma pytania 11** — numeracja przeskakuje z 10 na 12.
Sprawdzone na poziomie układu strony, to nie błąd ekstrakcji. Dodatkowo
usunięto pyt. 61 jako duplikat (patrz niżej).

## Struktura egzaminu

- **pyt. 1–70** (bez 11) — typ K: treść + 4 ponumerowane stwierdzenia,
- **pyt. 71–100** — jednokrotny wybór A–E, w tym 8 opisów przypadków
  klinicznych (80, 82, 83, 89, 90, 91, 95, 99).

W tym PDF nie wydrukowano legendy klucza A–E, ale nie jest ona potrzebna:
w formacie Only Pharms zapisujemy prawdziwość każdego stwierdzenia z osobna.

## Działy

94 z 98 pytań zaimportuje się od razu. Pozostałe 4 używają działów
zgłoszonych już przy pliku 2014/2015 — **nie dochodzi żaden nowy**:

| Section | Pytania 2016 |
|---|---|
| CHOROBA NIEDOKRWIENNA SERCA | 35, 37 |
| ZABURZENIA RYTMU SERCA | 42, 43 |

## Duplikat względem pliku 2014/2015

Pytanie **61** („Glikokortykosteroidy zmniejszają liczbę:") było
**identyczne** z pytaniem 61 z egzaminu 2014/2015 — zostało z tego pliku
usunięte.

Pytania 50 (2016) i 81 (2014/2015) mają identyczną treść pytania
(„Wskaż prawidłowe połączenie lek – działanie niepożądane:"), ale
**różne** odpowiedzi — to dwa różne pytania, oba warto zostawić.

## Dystraktory dopisane

Dla 18 pytań, w których wszystkie stwierdzenia są prawdziwe, egzamin nie
dostarcza dystraktora — dopisano po 3 wiarygodne farmakologicznie.
Wszystkie pozostałe `False` pochodzą wprost z egzaminu.

## Pytania przeformułowane

- **86** (mechanizm działania litu) — oryginalna poprawna odpowiedź brzmiała
  „żadnym z powyższych mechanizmów", co jako odpowiedź na fiszce nic nie uczy.
  Wpisano konkretny mechanizm: *hamowanie przemian fosfatydyloinozytoli
  w komórce*; cztery oryginalne opcje zostały dystraktorami.
- **1** to pytanie typu „z wyjątkiem" — jako `True` zapisano stwierdzenie
  będące wyjątkiem (czyli poprawną odpowiedzią), zgodnie z treścią polecenia.
- **23** i **95** (2014/2015) również są pytaniami przeczącymi („nie
  występuje", „nie zależy") — zapisane analogicznie.

## Poprawki literówek z oryginału

- pyt. 7: „cisaprid" → **cyzapryd**
- pyt. 8: „nerki analegetycznej" → **nerki analgetycznej**
- pyt. 15: „stanu astamtycznego" → **stanu astmatycznego**
- pyt. 25: „atenololol" → **atenolol**
- pyt. 26: „lewetiracetam" → **lewetyracetam**
- pyt. 28: „ketakonazol" → **ketokonazol**
- pyt. 32: „hipotonią otrostatyczną" → **hipotonią ortostatyczną**
- pyt. 39: „intropowe" → **inotropowe**
- pyt. 45: „aksetyl cefurosymu" → **aksetyl cefuroksymu**, „ko-trymoksazol"
  → **kotrimoksazol**
- pyt. 46: „spironololaktonem" → **spironolaktonem**
- pyt. 59: „cefriakson" → **ceftriakson**
- pyt. 62: „sekrecje" → **sekrecję**
- pyt. 73: „monokloalnym" → **monoklonalnym**
- pyt. 99: „nitrogiceryny" → **nitrogliceryna**
- pyt. 100: „ksylometazolina … antagonistą" — zachowano (to dystraktor)

## Pytania do weryfikacji merytorycznej

Egzamin nie zawiera klucza — poprawne odpowiedzi ustalono merytorycznie.
Warte rzutu okiem przed wgraniem:

- **13** (metformina) — „wskazana zwłaszcza u osób w podeszłym wieku
  z otyłością" oznaczono jako fałsz (ryzyko kwasicy mleczanowej).
- **16** (budezonid donosowo) — przyjęto, że wszystkie cztery działania
  są możliwe; „grzybica jamy ustnej" i „chrypka" są bardziej typowe dla
  postaci wziewnej do oskrzeli.
- **25** — atenolol oznaczono jako fałsz; beta-adrenolityki bywają
  stosowane pomocniczo w zespole abstynencyjnym, ale leczeniem są
  benzodiazepiny.
- **43** — ściśle biorąc ani metoprolol, ani digoksyna nie przywracają
  rytmu zatokowego (kontrolują częstość); to najbliższa poprawna para
  spośród podanych.
- **65** — „duża zawartość wody w organizmie" oznaczono jako fałsz
  (duża Vd wskazuje na kumulację tkankową i lipofilność).
- **75** — jako antagonistę 5-HT2 wskazano trazodon; mirtazapina również
  blokuje 5-HT2, ale klasyfikacyjnie to trazodon (SARI).
- **83** — wybrano reduktazę dihydrofolianową (trimetoprim); opcja
  „kwasu p-aminobenzoesowego" odnosi się do składowej sulfonamidowej.
- **91** — wskazano flukonazol (dobra penetracja do PMR); w indukcji
  stosuje się amfoterycynę B z flucytozyną.
- **92** — wskazano meflochinę; doksycyklina jest równie poprawną
  opcją profilaktyki w rejonach oporności na chlorochinę.

Źródło: egzamin z farmakologii, kierunek lekarsko-dentystyczny, lato 2016.
