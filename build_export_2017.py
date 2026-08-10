# -*- coding: utf-8 -*-
"""
Buduje plik eksportowy Only Pharms z egzaminu:
Farmakologia, Stomatologia rok III, WUM 2016/2017, wersja 1 (100 pytan).

Ten egzamin ZAWIERA klucz odpowiedzi ("Odp. A" itd.) - przypisanie True/False
wynika z klucza, nie z wlasnej oceny (wyjatki opisane w README).

PDF jest wersja robocza ze sladami redakcji: stare i nowe brzmienie sklejone
w jeden ciag (np. "ketoprofenmetamizol") oraz znaczniki [E1]..[E6].
Wszedzie przyjeto brzmienie NOWSZE (wstawione), zgodnie z konwencja
"skreslone -> wstawione".

Uklad kolumn: Section, Category, Poziom, Type, Question, True 1..5, False 1..5
"""
import csv
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, PatternFill
from openpyxl.utils import get_column_letter

# (nr, Section, Category, Question, [True...], [False...], klucz_egzaminu)
Q = [
(1, "FARMAKOKINETYKA", "Farmakokinetyka",
 "Stan stacjonarny:",
 ["ustala się na podstawie znajomości T1/2",
  "osiągany jest w przybliżeniu po 4 okresach półtrwania",
  "to stan, w którym kolejne podanie leku nie zmienia średniego stężenia leku we krwi"],
 ["to parametr nieistotny w praktyce klinicznej"], "A"),

(2, "FARMAKOKINETYKA", "Farmakokinetyka",
 "Zjawisko samoindukcji:",
 ["to zdolność niektórych leków do przyspieszania własnego metabolizmu",
  "rozwija się podczas długotrwałego stosowania niektórych leków",
  "może być przyczyną powstawania tolerancji na lek",
  "może zwiększać skuteczność terapeutyczną leku, gdy formą aktywną jest metabolit"],
 ["polega na hamowaniu własnego metabolizmu przez lek",
  "występuje wyłącznie po podaniu dożylnym",
  "dotyczy wyłącznie leków wydalanych w postaci niezmienionej"], "E"),

(3, "LEKI PRZECIWHISTAMINOWE", "Wskazania",
 "Lekami przeciwwymiotnymi najczęściej stosowanymi w chorobie lokomocyjnej są:",
 ["difenhydramina", "dimenhydrynat"],
 ["loratadyna", "aprepitant"], "C"),

(4, "LEKI PRZECIWHISTAMINOWE", "Działania niepożądane",
 "Do działań niepożądanych antagonistów receptora histaminowego typu 1 "
 "I generacji należą:",
 ["zaburzenia akomodacji",
  "trudności w oddawaniu moczu u osób z przerostem gruczołu krokowego",
  "senność"],
 ["bradykardia"], "A"),

(5, "NIEOPIOIDOWE LEKI PRZECIWBÓLOWE", "Mechanizm działania",
 "Lekiem selektywnie hamującym cyklooksygenazę typu 2 jest:",
 ["celekoksyb"],
 ["ibuprofen", "kwas acetylosalicylowy", "ketoprofen"], "D"),

(6, "NIEOPIOIDOWE LEKI PRZECIWBÓLOWE", "Wskazania",
 "U dzieci poniżej 12 r.ż. jako lek przeciwgorączkowy stosowany jest:",
 ["ibuprofen", "paracetamol"],
 ["kwas acetylosalicylowy", "metamizol"], "B"),

(7, "UKŁAD POKARMOWY", "Klasyfikacja",
 "Jako lek zapierający zastosujesz:",
 ["loperamid", "difenoksylat"],
 ["siarczan magnezu", "preparat z ziela senesu"], "B"),

(8, "UKŁAD POKARMOWY", "Wskazania",
 "Wskaż prawdziwe stwierdzenia dotyczące leczenia biegunek:",
 ["leki zapierające mogą wydłużać leczenie biegunek zakaźnych",
  "w terapii biegunek można podawać leki adsorpcyjne",
  "postępowanie w biegunkach obejmuje nawodnienie i podaż elektrolitów"],
 ["w biegunkach bakteryjnych zawsze wskazane jest podanie antybiotyku"], "A"),

(9, "UKŁAD ODDECHOWY", "Wskazania",
 "W przypadku kaszlu skuteczne może być podanie:",
 ["pentoksyweryna", "dekstrometorfan"],
 ["jodek potasu", "gwajakolosulfonian"], "C"),

(10, "UKŁAD POKARMOWY", "Klasyfikacja",
 "Osmotycznym lekiem przeczyszczającym jest:",
 ["laktuloza", "siarczan sodu"],
 ["olej rycynowy", "parafina płynna"], "C"),

(11, "LEKI PRZECIWPASOŻYTNICZE", "Wskazania",
 "W zakażeniu tasiemcem możesz zastosować:",
 ["albendazol", "prazykwantel"],
 ["metronidazol", "pentamidyna"], "B"),

(12, "LEKI PRZECIWBAKTERYJNE", "Wskazania",
 "W leczeniu oparzeń zastosujesz:",
 ["azotan srebra", "jodowany poliwidon"],
 ["boran sodu", "nadmanganian potasu"], "C"),

(13, "LEKI PRZECIWBAKTERYJNE", "Interakcje",
 "Reakcja disulfiramowa może zachodzić podczas jednoczesnego stosowania "
 "alkoholu etylowego oraz:",
 ["metronidazol"],
 ["fosfomycyna", "amikacyna", "erytromycyna"], "D"),

(14, "LEKI ROŚLINNE I SUPLEMENTY", "Klasyfikacja",
 "Saponiny występują w:",
 ["korzeń pierwiosnka", "korzeń lukrecji"],
 ["liście dębu", "kłącze rzewienia dłoniastego"], "B"),

(15, "ANTYKONCEPCJA", "Przeciwwskazania",
 "Przeciwwskazaniem do stosowania doustnych środków antykoncepcyjnych "
 "zawierających wyłącznie gestageny jest:",
 ["świeżo przebyta choroba zakrzepowo-zatorowa", "rak piersi"],
 ["nadciśnienie tętnicze", "niewydolność serca"], "C"),

(16, "LEKI PRZECIWLĘKOWE", "Mechanizm działania",
 "Benzodiazepiny:",
 ["obniżają ciśnienie tętnicze", "nasilają przewodnictwo chlorkowe",
  "hamują ośrodek oddechowy", "powodują niepamięć wsteczną"],
 ["blokują receptory GABA-A", "pobudzają ośrodek oddechowy",
  "podwyższają ciśnienie tętnicze"], "E"),

(17, "FARMAKOKINETYKA", "Farmakokinetyka",
 "Metabolity leków:",
 ["są lepiej rozpuszczalne w wodzie niż pierwotne związki"],
 ["są lepiej rozpuszczalne w lipidach niż pierwotny związek",
  "są określane jako „proleki”",
  "są eliminowane wolniej z organizmu niż pierwotne związki"], "D"),

(18, "ZNIECZULENIA OGÓLNE", "Działania niepożądane",
 "Podczas stosowania tiopentalu może wystąpić:",
 ["nasilenie objawów porfirii wskutek wzrostu aktywności syntazy kwasu "
  "delta-aminolewulinowego", "skurcz krtani"],
 ["żółtaczka cholestatyczna", "agranulocytoza"], "C"),

(19, "NIEOPIOIDOWE LEKI PRZECIWBÓLOWE", "Działania niepożądane",
 "Do działań niepożądanych kwasu acetylosalicylowego należy:",
 ["zaburzenie czynności nerek", "skurcz oskrzeli",
  "zwiększenie stężenia transaminaz w surowicy", "krwawienie z żołądka"],
 ["nadkrzepliwość", "hipertermia złośliwa", "przerost dziąseł"], "E"),

(20, "ZNIECZULENIA OGÓLNE", "Działania niepożądane",
 "Ketamina:",
 ["zwiększa zużycie tlenu przez tkankę mózgową", "powoduje halucynacje",
  "pobudza czynność serca", "powoduje koszmary senne"],
 ["obniża ciśnienie tętnicze", "zwalnia czynność serca",
  "zmniejsza przepływ mózgowy"], "E"),

(21, "ZNIECZULENIA OGÓLNE", "Działania niepożądane",
 "Podtlenek azotu:",
 ["powoduje niedotlenienie tkanek"],
 ["jest nefrotoksyczny", "podlega metabolizmowi wątrobowemu",
  "jest hepatotoksyczny"], "D"),

(22, "LEKI PRZECIWDEPRESYJNE", "Działania niepożądane",
 "Działaniem niepożądanym leków hamujących wychwyt zwrotny serotoniny jest:",
 ["nudności", "biegunka", "zawroty głowy", "dysfunkcje seksualne"],
 ["agranulocytoza", "hipertermia złośliwa", "przerost dziąseł"], "E"),

(23, "HEMOSTAZA", "Działania niepożądane",
 "Działaniem niepożądanym heparyny jest:",
 ["łysienie", "trombocytopenia", "osteoporoza", "krwawienie"],
 ["hiperaldosteronizm", "nadkrzepliwość pierwotna", "methemoglobinemia"], "E"),

(24, "NIEOPIOIDOWE LEKI PRZECIWBÓLOWE", "Działania niepożądane",
 "Stosowanie piroksykamu może prowadzić do:",
 ["wystąpienia obrzęków", "gorszej kontroli ciśnienia tętniczego"],
 ["obniżonej tolerancji glukozy", "moczówki prostej"], "B"),

(25, "LEKI PRZECIWBAKTERYJNE", "Farmakokinetyka",
 "Aminoglikozydy:",
 ["kumulują się w korze nerek",
  "są wydzielane w formie niezmienionej przez nerki"],
 ["dobrze przechodzą przez barierę krew-mózg",
  "dobrze wchłaniają się z przewodu pokarmowego"], "C"),

(26, "LEKI PRZECIWBAKTERYJNE", "Wskazania",
 "Klindamycyna jest stosowana w zakażeniu:",
 ["Bacteroides fragilis", "Propionibacterium acnes", "Staphylococcus aureus"],
 ["Actinomyces israelii"], "A"),

(27, "LEKI PRZECIWPASOŻYTNICZE", "Mechanizm działania",
 "Embonian pyrantelu:",
 ["pobudza receptory nikotynowe",
  "jest stosowany w zakażeniu Ascaris lumbricoides"],
 ["potęguje aktywność GABA", "blokuje pobieranie glukozy"], "C"),

(28, "UKŁAD PRZYWSPÓŁCZULNY", "Wskazania",
 "Neostygmina może być stosowana w:",
 ["zatrzymaniu moczu", "przedawkowaniu d-tubokuraryny",
  "porażennej niedrożności jelita", "terapii jaskry"],
 ["zatruciu związkami fosforoorganicznymi", "astmie oskrzelowej",
  "bradykardii zatokowej"], "E"),

(29, "LEKI PRZECIWGRZYBICZE", "Wskazania",
 "Flukonazol jest stosowany w zakażeniu:",
 ["Cryptococcus", "Histoplasma", "Coccidioides", "Candida"],
 ["Aspergillus", "Mucor", "Epidermophyton"], "E"),

(30, "LEKI PRZECIWGRZYBICZE", "Mechanizm działania",
 "Terbinafina:",
 ["jest inhibitorem aktywności epoksydazy skwalenowej",
  "może być podawana miejscowo i doustnie"],
 ["uszkadza chrząstkę stawową", "działa fototoksycznie"], "C"),

(31, "LEKI PRZECIWBAKTERYJNE", "Wskazania",
 "Metronidazol jest stosowany w zakażeniu:",
 ["Giardia lamblia", "Bacteroides fragilis", "Entamoeba histolytica",
  "Clostridium difficile"],
 ["Pseudomonas aeruginosa", "Candida albicans", "Mycobacterium tuberculosis"], "E"),

(32, "LEKI PRZECIWPSYCHOTYCZNE", "Działania niepożądane",
 "Złośliwy zespół neuroleptyczny objawia się:",
 ["wzrostem aktywności kinazy fosfokreatynowej", "sztywnością mięśni",
  "dysfagią"],
 ["hipotermią"], "A"),

(33, "HEMOSTAZA", "Mechanizm działania",
 "Klopidogrel:",
 ["hamuje wiązanie ADP z receptorem płytkowym P2Y12",
  "może być podawany łącznie z kwasem acetylosalicylowym u chorych "
  "po założeniu stentu naczyniowego"],
 ["powoduje hiperkaliemię", "zwiększa ryzyko wystąpienia zespołu Reye’a"], "B"),

(34, "NADCIŚNIENIE TĘTNICZE", "Wskazania",
 "W stanie nagłym w celu obniżenia ciśnienia tętniczego zastosujesz:",
 ["urapidyl", "labetalol", "nitroprusydek sodu"],
 ["guanetydyna"], "A"),

(35, "UKŁAD PRZYWSPÓŁCZULNY", "Mechanizm działania",
 "Pilokarpina:",
 ["przechodzi przez barierę krew-mózg", "powoduje skurcz mięśnia rzęskowego",
  "ma powinowactwo do receptorów muskarynowych"],
 ["jest rozkładana przez acetylocholinoesterazę"], "A"),

(36, "LEKI PRZECIWPSYCHOTYCZNE", "Wskazania",
 "Neuroleptyki mogą być stosowane:",
 ["w leczeniu uporczywej czkawki", "jako leki przeciwwymiotne",
  "w leczeniu świądu", "w leczeniu psychoz"],
 ["w leczeniu choroby Parkinsona", "w leczeniu jaskry",
  "w leczeniu nużliwości mięśni"], "E"),

(37, "LEKI PRZECIWBAKTERYJNE", "Wskazania",
 "Połączenie trimetoprim/sulfametoksazol jest stosowane:",
 ["w zakażeniach przewodu pokarmowego",
  "w bakteryjnym zakażeniu gruczołu krokowego",
  "w zakażeniu dróg moczowych",
  "w zapaleniu płuc wywołanym przez Pneumocystis jirovecii"],
 ["w zakażeniu Pseudomonas aeruginosa", "w gruźlicy",
  "w rzekomobłoniastym zapaleniu jelit"], "E"),

(38, "HEMOSTAZA", "Mechanizm działania",
 "Kwas epsilon-aminokapronowy:",
 ["hamuje rozpuszczanie skrzepu",
  "jest stosowany w krwawieniach spowodowanych niedoborem fibrynogenu",
  "może wywołać zakrzepowe zapalenie żył"],
 ["nasila działanie streptokinazy"], "A"),

(39, "LEKI PRZECIWGRZYBICZE", "Wskazania",
 "Miejscowo w grzybicy jamy ustnej zastosujesz:",
 ["nystatyna"],
 ["klotrimazol", "amfoterycyna B", "amorolfina"], "D"),

(40, "LEKI PRZECIWDEPRESYJNE", "Wskazania",
 "W profilaktyce choroby afektywnej dwubiegunowej możesz zastosować:",
 ["arypiprazol", "kwas walproinowy", "węglan litu"],
 ["droperidol"], "A"),

(41, "LEKI PRZECIWDEPRESYJNE", "Działania niepożądane",
 "W przebiegu stosowania węglanu litu może wystąpić:",
 ["hiperglikemia", "drżenie włókienkowe", "niedoczynność tarczycy",
  "wielomocz"],
 ["nadczynność tarczycy", "skąpomocz", "agranulocytoza"], "E"),

(42, "LEKI PRZECIWPSYCHOTYCZNE", "Działania niepożądane",
 "Podczas stosowania kwetiapiny może wystąpić:",
 ["przyrost masy ciała", "zwiększenie stężenia cholesterolu we krwi",
  "suchość w ustach"],
 ["skłonność do hazardu"], "A"),

(43, "LEKI PRZECIWLĘKOWE", "Wskazania",
 "W przebiegu zespołu abstynencyjnego w chorobie alkoholowej zastosujesz:",
 ["propranolol", "diazepam"],
 ["citalopram", "klometiazol"], "C"),

(44, "UKŁAD POKARMOWY", "Mechanizm działania",
 "Wskaż prawidłowe połączenie lek – mechanizm działania:",
 ["skopolamina – antagonista receptora muskarynowego",
  "palonosetron – antagonista receptora 5-HT3",
  "metoklopramid – antagonista receptora D2",
  "aprepitant – antagonista receptora neurokininy 1"],
 ["ondansetron – agonista receptora 5-HT3",
  "domperidon – agonista receptora D2",
  "dimenhydrynat – agonista receptora H1"], "E"),

(45, "LEKI NASENNE", "Wskazania",
 "W krótkotrwałym leczeniu bezsenności zastosujesz:",
 ["zopiklon"],
 ["ramelteon", "hydroksyzyna", "tiopental"], "D"),

(46, "FARMAKOKINETYKA", "Interakcje",
 "Enzymy mikrosomalne wątroby indukuje:",
 ["fenytoina", "karbamazepina", "fenobarbital"],
 ["kwas walproinowy"], "A"),

(47, "PADACZKA", "Działania niepożądane",
 "Zmniejszenie masy ciała może wystąpić w przebiegu stosowania:",
 ["karbamazepina"],
 ["kwas walproinowy", "chloropromazyna", "betametazon"], "D"),

(48, "BÓLE GŁOWY", "Wskazania",
 "W zapobieganiu migrenie zastosujesz:",
 ["topiramat", "metoprolol", "walproinian sodu", "toksyna botulinowa"],
 ["sumatryptan", "ergotamina", "paracetamol"], "E"),

(49, "NADCIŚNIENIE TĘTNICZE", "Farmakokinetyka",
 "Korzystny współczynnik T/P (trough-to-peak) wykazuje:",
 ["lacydypina", "peryndopryl", "felodypina CR", "telmisartan"],
 ["kaptopryl", "nifedypina o natychmiastowym uwalnianiu", "prazosyna"], "E"),

(50, "LEKI PRZECIWPSYCHOTYCZNE", "Działania niepożądane",
 "Zaburzenia pozapiramidowe mogą wystąpić w przebiegu stosowania:",
 ["chloropromazyna", "metoklopramid"],
 ["lorazepam", "rywastygmina"], "B"),

(51, "CUKRZYCA", "Działania niepożądane",
 "W przebiegu leczenia insuliną może wystąpić:",
 ["przyrost masy ciała", "lipodystrofia", "hipoglikemia"],
 ["kwasica ketonowa"], "A"),

(52, "TARCZYCA", "Mechanizm działania",
 "Obwodową konwersję T4 w T3 hamuje:",
 ["amiodaron", "propranolol", "jodowe środki kontrastowe",
  "propylotiouracyl"],
 ["tiamazol", "lewotyroksyna", "karbimazol"], "E"),

(53, "NIEOPIOIDOWE LEKI PRZECIWBÓLOWE", "Działania niepożądane",
 "Do działań niepożądanych metamizolu należy:",
 ["wypadanie włosów", "agranulocytoza"],
 ["zwiększenie ciśnienia śródgałkowego", "nykturia"], "C"),

(54, "LEKI PRZECIWLĘKOWE", "Wskazania",
 "W terapii lęku zastosujesz:",
 ["diazepam", "escitalopram"],
 ["rysperydon", "tiamytal"], "C"),

(55, "GLIKOKORTYKOSTEROIDY", "Działania niepożądane",
 "W przebiegu stosowania prednizonu może wystąpić:",
 ["nadpłytkowość", "hipokaliemia", "zaćma", "wzrost ciśnienia tętniczego"],
 ["hiperkaliemia", "małopłytkowość", "hipotonia ortostatyczna"], "E"),

(56, "UKŁAD WSPÓŁCZULNY", "Wskazania",
 "W leczeniu nieżytu nosa możesz zastosować:",
 ["pseudoefedryna", "fenylefryna", "tetryzolina"],
 ["okseladyna"], "A"),

(57, "LEKI PRZECIWBAKTERYJNE", "Wskazania",
 "Podchloryn sodu może być stosowany:",
 ["do dezynfekcji kanałów korzeniowych w leczeniu endodontycznym",
  "do amputacji przyżyciowej miazgi w zębach mlecznych"],
 ["do dezynfekcji skóry chorego przed operacją", "w leczeniu czyraczności"], "B"),

(58, "CHOROBA NIEDOKRWIENNA SERCA", "Wskazania",
 "Celem profilaktyki wtórnej incydentu sercowo-naczyniowego należy stosować:",
 ["rozuwastatyna", "kwas acetylosalicylowy"],
 ["bisoprolol", "nitrogliceryna"], "C"),

(59, "LEKI MOCZOPĘDNE", "Wskazania",
 "Furosemid zastosujesz w:",
 ["niewydolności serca", "ostrej niewydolności nerek", "obrzęku płuc",
  "nadciśnieniu tętniczym"],
 ["hipowolemii", "hiponatremii", "zatruciu glikozydami naparstnicy"], "E"),

(60, "LEKI MOCZOPĘDNE", "Działania niepożądane",
 "Hipokaliemia może wystąpić w przebiegu stosowania:",
 ["salbutamol", "prednizon", "indapamid"],
 ["chinapryl"], "A"),

(61, "LEKI PRZECIWBAKTERYJNE", "Wskazania",
 "W trądziku młodzieńczym zastosujesz:",
 ["kwas fusydowy", "erytromycyna", "tetracyklina"],
 ["mupirocyna"], "A"),

(62, "LEKI PRZECIWBAKTERYJNE", "Wskazania",
 "Gramicydynę zastosujesz w:",
 ["zakażeniu błony naczyniowej", "zakażeniach przewodu słuchowego zewnętrznego"],
 ["zakażeniu protez stawowych", "zanokcicy"], "B"),

(63, "LEKI MOCZOPĘDNE", "Działania niepożądane",
 "Zasadowicę metaboliczną mogą powodować:",
 ["diuretyki pętlowe", "tiazydy"],
 ["inhibitory anhydrazy węglanowej", "diuretyki oszczędzające potas"], "C"),

(64, "STANY NAGŁE", "Działania niepożądane",
 "Hydroksyetyloskrobia:",
 ["może być stosowana w przebiegu wstrząsu septycznego",
  "powoduje świąd skóry oporny na działanie leków przeciwhistaminowych",
  "może powodować zaburzenia krzepnięcia krwi", "uszkadza nerki"],
 ["jest hepatotoksyczna", "powoduje agranulocytozę",
  "jest przeciwwskazana w hipowolemii"], "E"),

(65, "ZNIECZULENIA MIEJSCOWE", "Wskazania",
 "Lidokaina może być stosowana w:",
 ["neuralgii popółpaścowej", "bólach fantomowych", "bolesnej polineuropatii",
  "znieczuleniu skóry przed cewnikowaniem żył"],
 ["znieczuleniu ogólnym wziewnym", "premedykacji doustnej",
  "leczeniu methemoglobinemii"], "E"),

(66, "PADACZKA", "Wskazania",
 "W stanie padaczkowym celowe jest zastosowanie:",
 ["lorazepam", "fenytoina", "fenobarbital"],
 ["gabapentyna"], "A"),

(67, "PADACZKA", "Wskazania",
 "Wskaż prawidłowe połączenie lek – wskazanie:",
 ["drżenie samoistne – propranolol",
  "padaczka z napadami uogólnionymi toniczno-klonicznymi – kwas walproinowy"],
 ["zespół niespokojnych nóg – sertralina",
  "poród przedwczesny – dihydroergotamina"], "B"),

(68, "GLIKOKORTYKOSTEROIDY", "Sposób podania",
 "Wziewnie stosuje się:",
 ["mometazon", "cyklezonid"],
 ["deksametazon", "hydrokortyzon"], "C"),

(69, "LEKI PRZECIWDEPRESYJNE", "Wskazania",
 "Trójpierścieniowe leki przeciwdepresyjne możesz zastosować w leczeniu:",
 ["moczenia nocnego", "zespołu natręctw", "ciężkiej postaci depresji"],
 ["zespołu neuroleptycznego"], "A"),

(70, "LEKI PRZECIWBAKTERYJNE", "Wskazania",
 "W zakażeniu kości zastosujesz:",
 ["teikoplanina", "klindamycyna"],
 ["limecyklina", "kolistyna"], "C"),

(71, "CHOROBA PARKINSONA", "Wskazania",
 "W leczeniu choroby Parkinsona możesz zastosować:",
 ["pramipeksol", "selegilina", "L-dopa"],
 ["moklobemid"], "A"),

(72, "LEKI MOCZOPĘDNE", "Działania niepożądane",
 "Zmniejszenie stężenia wapnia w surowicy może wystąpić podczas terapii:",
 ["torasemid", "metyloprednizolon"],
 ["fenoterol", "diltiazem"], "B"),

# --- pytania jednokrotnego wyboru (73-100) ---

(73, "UKŁAD ODDECHOWY", "Sposób podania",
 "Wskaż odpowiedź fałszywą dotyczącą leków działających na drogi oddechowe:",
 ["ostatnią dawkę acetylocysteiny należy przyjąć tuż przed snem"],
 ["przy podawaniu leków wykrztuśnych i sekretolitycznych należy zadbać "
  "o odpowiednią podaż płynów",
  "ambroksol jest metabolitem bromheksyny",
  "kodeina dobrze wchłania się z przewodu pokarmowego",
  "N-acetylocysteinę podaje się przy przedawkowaniu paracetamolu"], "B"),

(74, "UKŁAD POKARMOWY", "Mechanizm działania",
 "Mechanizm działania pantoprazolu polega na:",
 ["hamowaniu aktywności pompy protonowej"],
 ["blokowaniu receptora histaminowego typu 2",
  "zwiększeniu stężenia prostaglandyny E w żołądku",
  "specyficznym blokowaniu receptorów muskarynowych typu 1",
  "działaniu neutralizującym"], "D*"),

(75, "LEKI PRZECIWBAKTERYJNE", "Interakcje",
 "Suplementacja pirydoksyną jest zalecana podczas stosowania:",
 ["izoniazyd"],
 ["streptomycyna", "pirazynamid", "etambutol", "etionamid"], "D"),

(76, "LEKI MOCZOPĘDNE", "Mechanizm działania",
 "Antagonistą aldosteronu jest:",
 ["eplerenon"],
 ["triamcynolon", "ketokonazol", "gancyklowir", "sumatryptan"], "B"),

(77, "CHOROBA NIEDOKRWIENNA SERCA", "Interakcje",
 "Podczas stosowania diazotanu izosorbidu nie możesz zastosować:",
 ["sildenafil"],
 ["chlortalidon", "fenytoina", "deksametazon", "spironolakton"], "D"),

(78, "LEKI ROŚLINNE I SUPLEMENTY", "Mechanizm działania",
 "Nasiona Foeniculum vulgare wykazują działanie:",
 ["estrogenne"],
 ["nefrotoksyczne", "hepatotoksyczne", "karcinogenne", "zapierające"], "C"),

(79, "LEKI PRZECIWWIRUSOWE", "Wskazania",
 "W leczeniu grypy zastosujesz:",
 ["oseltamiwir"],
 ["acyklowir", "kapsaicyna", "abakawir", "lamiwudyna"], "E"),

(80, "CUKRZYCA", "Działania niepożądane",
 "Niedokrwistość megaloblastyczna może wystąpić podczas terapii:",
 ["metformina"],
 ["gliklazyd", "nateglinid", "sitagliptyna", "pioglitazon"], "B"),

(81, "NADCIŚNIENIE TĘTNICZE", "Działania niepożądane",
 "Stężenie kwasu moczowego w surowicy obniża:",
 ["walsartan"],
 ["klonidyna", "hydrochlorotiazyd", "metoprolol", "terazosyna"], "C"),

(82, "NADCIŚNIENIE TĘTNICZE", "Mechanizm działania",
 "Inhibitorem reniny jest:",
 ["aliskiren"],
 ["rezerpina", "urapidil", "benazepryl", "moksonidyna"], "D"),

(83, "HEMOSTAZA", "Mechanizm działania",
 "Bezpośrednim inhibitorem czynnika Xa jest:",
 ["rywaroksaban"],
 ["dabigatran", "fondaparynuks", "eptifibatyd", "prasugrel"], "A"),

(84, "PADACZKA", "Działania niepożądane",
 "Wskaż fałszywe twierdzenie dotyczące karbamazepiny:",
 ["może powodować zespół parkinsonowski"],
 ["stosowana jest w neuralgii nerwu trójdzielnego", "działa przeciwdrgawkowo",
  "powoduje hiponatremię", "powoduje zmniejszenie masy ciała"], "B"),

(85, "NADCIŚNIENIE TĘTNICZE", "Wskazania",
 "Amlodypina stosowana jest w:",
 ["nadciśnieniu tętniczym"],
 ["obrzękach pochodzenia nerkowego", "zaburzeniach rytmu serca",
  "kardiomiopatii przerostowej", "nefropatii cukrzycowej"], "D"),

(86, "LEKI PRZECIWDEPRESYJNE", "Mechanizm działania",
 "Wskaż fałszywe twierdzenie dotyczące amitryptyliny:",
 ["selektywnie hamuje wychwyt zwrotny serotoniny"],
 ["działa cholinolitycznie", "stosowana jest w bólu neuropatycznym",
  "obniża próg drgawkowy", "działa przeciwdepresyjnie"], "D"),

(87, "UKŁAD PRZYWSPÓŁCZULNY", "Wskazania",
 "W terapii myasthenia gravis zastosujesz:",
 ["pirydostygmina"],
 ["propylotiouracyl", "selegilina", "lewetyracetam", "oksybutynina"], "C"),

(88, "LEKI PRZECIWBAKTERYJNE", "Wskazania",
 "W zakażeniach metycylinoopornymi szczepami Staphylococcus aureus (MRSA) "
 "zastosujesz:",
 ["linezolid"],
 ["kolistyna", "kloksacylina", "spektynomycyna", "tygecyklina"], "A"),

(89, "LEKI PRZECIWBAKTERYJNE", "Działania niepożądane",
 "Zaburzenia słuchu mogą wystąpić w przebiegu stosowania:",
 ["wankomycyna"],
 ["cefuroksym", "dalfoprystyna/chinuprystyna", "azytromycyna",
  "doksycyklina"], "E"),

(90, "OTĘPIENIE", "Wskazania",
 "W terapii choroby Alzheimera możesz zastosować:",
 ["donepezil"],
 ["sulpiryd", "pergolid", "lamotrygina", "sumatryptan"], "B"),

(91, "LEKI PRZECIWBAKTERYJNE", "Farmakokinetyka",
 "Zaznacz twierdzenie fałszywe dotyczące amikacyny:",
 ["dobrze wchłania się z przewodu pokarmowego"],
 ["ma efekt poantybiotykowy",
  "wykazuje synergizm z beta-laktamami w stosunku do enterokoków",
  "może spowodować ostrą niewydolność nerek", "nie działa na beztlenowce"], "B"),

(92, "CUKRZYCA", "Mechanizm działania",
 "Wydzielanie insuliny stymuluje:",
 ["liraglutyd"],
 ["miglitol", "metformina", "pioglitazon", "pramlintyd"], "brak*"),

(93, "NIEWYDOLNOŚĆ SERCA", "Przeciwwskazania",
 "W niewydolności serca przeciwwskazane jest podawanie:",
 ["werapamil", "nimesulid"],
 ["iwabradyna", "amiodaron"], "E*"),

(94, "LEKI HIPOLIPEMIZUJĄCE", "Mechanizm działania",
 "Blaszki miażdżycowe w naczyniach wieńcowych stabilizuje:",
 ["simwastatyna"],
 ["enoksaparyna", "dabigatran", "digoksyna", "nebiwolol"], "D"),

(95, "UKŁAD ODDECHOWY", "Wskazania",
 "Do przerywania napadu duszności w astmie zalecisz:",
 ["bromek ipratropium w nebulizacji", "salbutamol dożylnie"],
 ["teofilina w syropie", "salmeterol w aerozolu"], "E"),

(96, "HEMOSTAZA", "Wskazania",
 "Do hamowania krwawień służy ogólnoustrojowe podanie:",
 ["kwas traneksamowy"],
 ["abciksimab", "kolagen", "dekstran", "mannitol"], "C"),

(97, "NADCIŚNIENIE TĘTNICZE", "Wskazania",
 "W nefropatii cukrzycowej celowe jest zastosowanie:",
 ["lizynopryl"],
 ["nitrendypina", "nitrogliceryna", "klopamid", "prazosyna"], "D"),

(98, "LEKI PRZECIWBAKTERYJNE", "Mechanizm działania",
 "Syntezę ściany komórkowej hamuje:",
 ["imipenem"],
 ["ryfampicyna", "lewofloksacyna", "netilmycyna"], "D"),

(99, "LEKI HIPOLIPEMIZUJĄCE", "Mechanizm działania",
 "Bezpośredni wpływ na zahamowanie wchłaniania cholesterolu z przewodu "
 "pokarmowego wykazuje:",
 ["ezetymib"],
 ["kolestyramina", "atorwastatyna", "fenofibrat", "niacyna"], "A"),

(100, "LEKI PRZECIWBAKTERYJNE", "Interakcje",
 "Ograniczenie pokarmów zawierających tyraminę wskazane jest podczas terapii:",
 ["żadnym z wymienionych antybiotyków"],
 ["klarytromycyna", "cyprofloksacyna", "tetracyklina", "amoksycylina"], "E"),
]

MAX_T, MAX_F = 5, 5
HEADER = (["Section", "Category", "Poziom", "Type", "Question"]
          + [f"True {i}" for i in range(1, MAX_T + 1)]
          + [f"False {i}" for i in range(1, MAX_F + 1)])


def rows():
    for nr, sec, cat, q, trues, falses, _key in Q:
        assert trues and falses, f"pytanie {nr}: brak True lub False"
        assert len(trues) <= MAX_T and len(falses) <= MAX_F, f"pytanie {nr}"
        yield ([sec, cat, "", "testowe", q]
               + trues + [""] * (MAX_T - len(trues))
               + falses + [""] * (MAX_F - len(falses)))


def main():
    data = list(rows())
    nums = [x[0] for x in Q]
    assert nums == list(range(1, 101)), "numeracja 1-100"
    assert len(data) == 100

    base = "farmakologia_egzamin_2017_OnlyPharms"
    with open(base + ".csv", "w", newline="", encoding="utf-8-sig") as f:
        w = csv.writer(f)
        w.writerow(HEADER)
        w.writerows(data)

    wb = Workbook()
    ws = wb.active
    ws.title = "Pytania"
    ws.append(HEADER)
    for r in data:
        ws.append(r)

    for c in ws[1]:
        c.font = Font(bold=True, color="FFFFFF")
        c.fill = PatternFill("solid", fgColor="2F5597")
        c.alignment = Alignment(vertical="center")
    ws.freeze_panes = "F2"
    ws.auto_filter.ref = f"A1:{get_column_letter(len(HEADER))}{len(data) + 1}"
    for col, wdt in {"A": 32, "B": 22, "C": 9, "D": 10, "E": 62}.items():
        ws.column_dimensions[col].width = wdt
    for i in range(6, len(HEADER) + 1):
        ws.column_dimensions[get_column_letter(i)].width = 34
    for row in ws.iter_rows(min_row=2):
        for c in row:
            c.alignment = Alignment(vertical="top", wrap_text=True)
    wb.save(base + ".xlsx")

    from collections import Counter
    print("wierszy:", len(data))
    print("odstepstwa od klucza egzaminu:",
          [nr for nr, *_r, k in Q if k.endswith("*")])
    print("\nSection:")
    for k, v in sorted(Counter(x[1] for x in Q).items()):
        print(f"  {v:3d}  {k}")
    print("\nCategory:")
    for k, v in sorted(Counter(x[2] for x in Q).items()):
        print(f"  {v:3d}  {k}")


if __name__ == "__main__":
    main()
