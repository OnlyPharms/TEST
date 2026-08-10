# -*- coding: utf-8 -*-
"""
Buduje plik eksportowy Only Pharms z dokumentu "Baza - Egzamin - 2022".

UWAGA CO DO ZRODLA: to NIE jest oficjalna ksiazeczka egzaminacyjna, tylko
studencka rekonstrukcja pisana wspolnie po egzaminie. Nie ma klucza
odpowiedzi, czesc opcji jest niekompletna, a w tresci sa wtracone dyskusje
("chyba", "+1", "nie pamietam", "?"). Wszystkie odpowiedzi ustalone
merytorycznie; komentarze studenckie usuniete z tresci pytan.

Z 78 ponumerowanych pozycji wyeksportowano 69. Pominieto 9 - lista
w README_eksport_2022.md.

Uklad kolumn: Section, Category, Poziom, Type, Question, True 1..5, False 1..5
"""
import csv
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, PatternFill
from openpyxl.utils import get_column_letter

# (nr, Section, Category, Question, [True...], [False...])
Q = [
(1, "UKŁAD POKARMOWY", "Farmakokinetyka",
 "Wskaż prawdziwe stwierdzenia dotyczące inhibitorów pompy protonowej:",
 ["nie różnią się siłą działania hamującego wydzielanie HCl w zakresie dawek "
  "równoważnych",
  "zwiększają prawdopodobieństwo zakażenia Clostridium difficile"],
 ["są podawane w postaci aktywnej, dlatego działają już w żołądku",
  "nie są metabolizowane przez CYP450, dlatego nie wchodzą w interakcje"]),

(2, "LEKI PRZECIWNOWOTWOROWE", "Wskazania",
 "Wspomagająco w leczeniu cytotoksycznym stosuje się:",
 ["palonosetron", "filgrastym"],
 ["talidomid", "letrozol"]),

(3, "LEKI PRZECIWHISTAMINOWE", "Wskazania",
 "W leczeniu późnej fazy odpowiedzi alergicznej z zatkaniem nosa skuteczny "
 "może być:",
 ["mometazon"],
 ["klemastyna", "nafazolina", "antazolina"]),

(4, "NADCIŚNIENIE TĘTNICZE", "Interakcje",
 "Wskaż korzystne (rekomendowane) skojarzenie leków:",
 ["indapamid + peryndopryl + rozuwastatyna",
  "bromek glikopironium + indakaterol + mometazon"],
 ["hydrochlorotiazyd + enalapryl + telmisartan",
  "lizynopryl + diltiazem + propranolol"]),

(5, "LEKI PRZECIWDEPRESYJNE", "Mechanizm działania",
 "Wychwyt zwrotny serotoniny hamuje:",
 ["fluoksetyna", "tramadol", "nefopam"],
 ["bupropion"]),

(6, "UKŁAD PRZYWSPÓŁCZULNY", "Wskazania",
 "Wskazaniem do podania atropiny może być:",
 ["premedykacja", "bradykardia zatokowa"],
 ["zatrucie bieluniem dziędzierzawą", "wstrząs anafilaktyczny"]),

(7, "FARMAKOKINETYKA", "Mechanizm działania",
 "Hamowanie aktywności enzymatycznej to podstawowy mechanizm działania:",
 ["amoksycylina", "racekadotryl", "kaptopryl", "selegilina"],
 ["formoterol", "losartan", "amlodypina"]),

(8, "UKŁAD PRZYWSPÓŁCZULNY", "Antidotum",
 "Pacjent ze zwężonymi źrenicami, ślinotokiem i bradykardią – wskaż "
 "prawdopodobną przyczynę:",
 ["zatrucie związkami fosforoorganicznymi"],
 ["zatrucie lulkiem czarnym", "przedawkowanie oksykodonu",
  "przedawkowanie sertraliny"]),

(9, "CHOROBA NIEDOKRWIENNA SERCA", "Wskazania",
 "Udowodniony korzystny wpływ sercowo-naczyniowy ma:",
 ["atorwastatyna", "kanagliflozyna", "kwas acetylosalicylowy"],
 ["saksagliptyna"]),

(10, "UKŁAD ODDECHOWY", "Działania niepożądane",
 "Skurcz oskrzeli może wywołać:",
 ["propafenon", "kwas acetylosalicylowy"],
 ["atropina", "drotaweryna"]),

(11, "FARMAKOKINETYKA", "Mechanizm działania",
 "Na receptor błonowy działa:",
 ["glikwidon", "dulaglutyd"],
 ["etynyloestradiol", "deksametazon"]),

(14, "LEKI PRZECIWBAKTERYJNE", "Działania niepożądane",
 "Wskaż leki mogące powodować reakcje fototoksyczne:",
 ["furosemid", "sulfasalazyna", "adapalen", "fluorochinolony"],
 ["amoksycylina", "ranitydyna", "insulina"]),

(15, "NADCIŚNIENIE TĘTNICZE", "Przeciwwskazania",
 "Wskaż prawidłowe połączenie lek – przeciwwskazanie:",
 ["atenolol – cukrzyca", "trandolapryl – ciąża"],
 ["amlodypina – blok przedsionkowo-komorowy", "doksazosyna – hiperkaliemia"]),

(16, "UKŁAD PRZYWSPÓŁCZULNY", "Przeciwwskazania",
 "W jaskrze z wąskim kątem przesączania przeciwwskazane są:",
 ["solifenacyna", "ipratropium", "oksybutynina"],
 ["mirabegron"]),

(17, "FARMAKOKINETYKA", "Mechanizm działania",
 "Miejsce wiązania zlokalizowane wewnątrzkomórkowo dotyczy:",
 ["lewofloksacyna", "atorwastatyna"],
 ["formoterol", "lamotrygina"]),

(18, "LEKI PRZECIWBAKTERYJNE", "Wskazania",
 "Wskaż prawidłowe połączenie lek – wskazanie:",
 ["spiramycyna – toksoplazmoza u kobiet w ciąży",
  "metronidazol – ameboza", "ceftriakson – rzeżączka"],
 ["fidaksomycyna – biegunka podróżnych"]),

(19, "CHOROBA NIEDOKRWIENNA SERCA", "Mechanizm działania",
 "Azotany organiczne:",
 ["są używane w celu obniżenia ciśnienia w sytuacjach nagłych",
  "ich zasadniczy mechanizm w chorobie niedokrwiennej serca na podłożu "
  "miażdżycy polega na zmniejszeniu naprężenia mięśniówki komór",
  "zwiększają stężenie cGMP w mięśniówce naczyń",
  "mogą powodować methemoglobinemię"],
 ["zmniejszają stężenie cGMP w mięśniówce naczyń",
  "mogą być bezpiecznie łączone z inhibitorami fosfodiesterazy typu 5",
  "działają przez blokowanie kanałów wapniowych"]),

(20, "LEKI PRZECIWBAKTERYJNE", "Wskazania",
 "Wskaż sytuację, w której chemioprofilaktyka antybiotykowa jest zgodna "
 "z aktualną wiedzą:",
 ["pacjentka z wysokim ryzykiem infekcyjnego zapalenia wsierdzia "
  "z planowanym leczeniem stomatologicznym"],
 ["nieprzytomny chory unieruchomiony na OIOM",
  "80-letnia pacjentka z nietrzymaniem moczu i bezobjawową bakteriurią",
  "starszy pacjent z COVID-19 i niewydolnością oddechową"]),

(21, "ZABURZENIA RYTMU SERCA", "Interakcje",
 "Wskaż kombinację leków o przeciwnych efektach chronotropowych:",
 ["diltiazem – atropina", "amiodaron – dobutamina"],
 ["iwabradyna – digoksyna", "efedryna – izoprenalina"]),

(22, "LEKI HIPOLIPEMIZUJĄCE", "Działania niepożądane",
 "Ryzyko miopatii i rabdomiolizy przy stosowaniu statyny wzrasta gdy:",
 ["jednoczesne stosowanie amlodypiny", "niedoczynność tarczycy",
  "stosowanie u kobiety powyżej 85 r.ż.", "nadużywanie alkoholu"],
 ["jednoczesne stosowanie ezetymibu", "nadczynność tarczycy",
  "stosowanie u młodego mężczyzny bez chorób współistniejących"]),

(23, "UKŁAD POKARMOWY", "Wskazania",
 "Wskaż prawdziwe stwierdzenia dotyczące leków działających na przewód "
 "pokarmowy:",
 ["pankreatyna nie jest wskazana w ostrym zapaleniu trzustki",
  "metoklopramid może powodować objawy pozapiramidowe, dlatego stosowanie go "
  "jako leku przeciwwymiotnego powinno mieć charakter doraźny"],
 ["fosforan sodu jako osmotyczny lek przeczyszczający nie powoduje zaburzeń "
  "elektrolitowych",
  "famotydyna jest elementem schematu eradykacji Helicobacter pylori"]),

(24, "NIEWYDOLNOŚĆ SERCA", "Wskazania",
 "U pacjenta z niewydolnością serca z obniżoną frakcją wyrzutową zastosujesz:",
 ["ramipryl", "eplerenon"],
 ["labetalol", "kanagliflozyna"]),

(25, "HEMOSTAZA", "Sposób podania",
 "Wskaż prawdziwe stwierdzenia dotyczące leczenia przeciwkrzepliwego:",
 ["efekt antagonistów witaminy K zależy od czasu półtrwania czynników "
  "krzepnięcia, dlatego w początkowej fazie leczenia stosuje się heparynę",
  "NOAC nie są zamiennikiem antagonistów witaminy K u pacjenta ze sztuczną "
  "zastawką",
  "heparynę drobnocząsteczkową stosuje się dożylnie lub podskórnie, "
  "ale nie domięśniowo"],
 ["heparyna niefrakcjonowana nie posiada swoistego antidotum, dlatego stosuje "
  "się leczenie objawowe"]),

(26, "LEKI PRZECIWBAKTERYJNE", "Wskazania",
 "Profilaktyka zakażenia miejsca operowanego:",
 ["jest dopasowana indywidualnie do potencjalnego czynnika zakaźnego"],
 ["jest każdorazowo konieczna, niezależnie od czystości pola operacyjnego",
  "jest zawsze jednorazowa, 30–60 minut przed nacięciem skóry",
  "jest zawsze podawana w dawce standardowej"]),

(27, "FARMAKOKINETYKA", "Farmakokinetyka",
 "Lek nie wchłania się z przewodu pokarmowego, wiąże się z białkami w 98%, "
 "nie jest metabolizowany, jest wydalany z moczem w postaci czynnej. "
 "Wskaż prawidłowe stwierdzenie:",
 ["przy niewydolności nerek wzrośnie jego okres półtrwania"],
 ["przy stosowaniu worykonazolu wzrośnie jego czas półtrwania",
  "przy stosowaniu aktywatorów glikoproteiny P zmniejszy się jego "
  "biodostępność po podaniu doustnym",
  "przy znacznej hipoalbuminemii zmaleje jego efekt kliniczny"]),

(29, "NARKOTYCZNE LEKI PRZECIWBÓLOWE", "Działania niepożądane",
 "W zatruciu opioidami występuje:",
 ["hipotermia", "niedociśnienie", "sztywność mięśniowa"],
 ["tachypnoe"]),

(30, "LEKI PRZECIWHISTAMINOWE", "Wskazania",
 "W leczeniu zawrotów głowy pochodzenia błędnikowego stosuje się:",
 ["betahistyna", "dimenhydrynat", "tietylperazyna"],
 ["feksofenadyna"]),

(31, "LEKI PRZECIWBAKTERYJNE", "Wskazania",
 "Fosfomycynę można zastosować w zakażeniu szczepami:",
 ["MDR", "CPE", "VRE", "MRSA"],
 ["Candida albicans", "Aspergillus fumigatus",
  "Mycobacterium tuberculosis"]),

(32, "FARMAKOKINETYKA", "Farmakokinetyka",
 "Przy genetycznie zmniejszonym metabolizmie leku w wątrobie:",
 ["okres półtrwania tego leku będzie dłuższy",
  "trzeba podawać mniejszą dawkę leku niż standardowo"],
 ["większy będzie jego klirens",
  "istnieje konieczność wdrożenia induktora tego enzymu"]),

(33, "LEKI PRZECIWHISTAMINOWE", "Farmakokinetyka",
 "Desloratadyna:",
 ["może być stosowana w alergicznym nieżycie nosa",
  "działa selektywnie na receptor H1"],
 ["istotnie wpływa na aktywność cytochromu P450",
  "może być stosowana pozajelitowo"]),

(34, "HEMOSTAZA", "Wskazania",
 "Alteplaza:",
 ["jest podawana w udarze niedokrwiennym mózgu",
  "należy ją podać jak najszybciej po wystąpieniu epizodu niedokrwiennego"],
 ["jest bezpośrednim inhibitorem plazminy",
  "jest lekiem I rzutu w ostrym zespole wieńcowym"]),

(35, "LEKI PRZECIWDEPRESYJNE", "Działania niepożądane",
 "Escitalopram:",
 ["przy przedawkowaniu może powodować rozszerzenie źrenic",
  "jest lekiem z grupy SSRI",
  "w początkowym okresie leczenia może nasilać lęk",
  "w początkowym okresie leczenia może zwiększać ryzyko samobójstwa"],
 ["jest trójpierścieniowym lekiem przeciwdepresyjnym",
  "przy przedawkowaniu powoduje zwężenie źrenic",
  "działa przeciwlękowo już od pierwszej dawki"]),

(36, "LEKI PRZECIWWIRUSOWE", "Wskazania",
 "W leczeniu COVID-19 stosuje się:",
 ["tocilizumab", "molnupirawir", "remdesywir"],
 ["iwermektyna"]),

(37, "FARMAKOKINETYKA", "Interakcje",
 "Interakcje wykazują zestawienia leków:",
 ["itopryd + skopolamina", "ondansetron + tramadol",
  "omeprazol + klopidogrel", "wankomycyna + gentamycyna"],
 ["paracetamol + witamina C", "amoksycylina + loratadyna",
  "insulina + wapń"]),

(38, "CHOROBA NIEDOKRWIENNA SERCA", "Mechanizm działania",
 "Wskaż prawidłowe połączenie lek – mechanizm działania:",
 ["ranolazyna – hamowanie późnego prądu sodowego",
  "amiodaron – antagonista kanałów potasowych"],
 ["fentolamina – wybiórczy antagonista receptora alfa-1",
  "metyldopa – antagonista receptora alfa-2"]),

(39, "OSTEOPOROZA", "Działania niepożądane",
 "Ryzedronian może powodować:",
 ["owrzodzenia przełyku", "martwica kości żuchwy"],
 ["powikłania sercowo-naczyniowe", "hamowanie mineralizacji kości"]),

(40, "LEKI HIPOLIPEMIZUJĄCE", "Mechanizm działania",
 "Wskaż prawdziwe zdanie o lekach wpływających na profil lipidowy:",
 ["ewolokumab hamuje PCSK-9",
  "ezetymib jest lekiem drugiego rzutu w leczeniu powikłań "
  "sercowo-naczyniowych"],
 ["fibraty znacząco podwyższają frakcję HDL",
  "inklisiran może wywołać ostre zapalenie trzustki"]),

(41, "LEKI PRZECIWBAKTERYJNE", "Działania niepożądane",
 "W leczeniu gruźlicy u pacjenta z niewydolnością wątroby nasilenia jej "
 "niewydolności można spodziewać się przy zastosowaniu:",
 ["izoniazyd", "pirazynamid", "ryfampicyna"],
 ["etambutol"]),

(42, "UKŁAD WSPÓŁCZULNY", "Mechanizm działania",
 "Wskaż efekty kliniczne związane z aktywacją receptorów beta-2:",
 ["rozkurcz mięśniówki gładkiej oskrzeli"],
 ["ślinotok", "hipoglikemia", "bradykardia"]),

(43, "LEKI PRZECIWBAKTERYJNE", "Wskazania",
 "W leczeniu pozaszpitalnego zapalenia płuc o etiologii paciorkowcowej "
 "zgodne z aktualnymi wskazaniami jest zastosowanie:",
 ["amoksycylina bez inhibitora beta-laktamaz przez 7 dni",
  "lewofloksacyna w przypadku nadwrażliwości na beta-laktamy przez 7 dni"],
 ["klarytromycyna przez 7 dni",
  "doksycyklina w przypadku nadwrażliwości na beta-laktamy przez 7 dni"]),

(44, "NIEWYDOLNOŚĆ SERCA", "Mechanizm działania",
 "Digoksyna działa:",
 ["zwiększa aktywność układu przywspółczulnego", "dromotropowo ujemnie",
  "inotropowo dodatnio"],
 ["hipotensyjnie"]),

(45, "CUKRZYCA", "Sposób podania",
 "W hipoglikemii można podać:",
 ["glukoza doustnie", "glukagon donosowo", "glukoza dożylnie",
  "glukagon podskórnie"],
 ["insulina podskórnie", "metformina doustnie", "akarboza doustnie"]),

(46, "STANY NAGŁE", "Antidotum",
 "Wskaż prawidłowe połączenie trucizna – odtrutka:",
 ["metanol – fomepizol", "cyjanek – hydroksykobalamina",
  "diazepam – flumazenil"],
 ["rtęć – deferoksamina"]),

(47, "LEKI PRZECIWBAKTERYJNE", "Wskazania",
 "Wskaż prawidłowo podane spektrum antybiotyku:",
 ["lewofloksacyna – Streptococcus pneumoniae, Pseudomonas aeruginosa"],
 ["cefotaksym – Pseudomonas aeruginosa, Klebsiella pneumoniae",
  "azytromycyna – Streptococcus pneumoniae, Pseudomonas aeruginosa",
  "klindamycyna – Proteus mirabilis"]),

(48, "CUKRZYCA", "Mechanizm działania",
 "Wskaż prawdziwe stwierdzenia dotyczące leków przeciwcukrzycowych:",
 ["liraglutyd zwiększa wydzielanie insuliny z komórek beta trzustki",
  "kanagliflozyna zmniejsza wchłanianie zwrotne glukozy w kanaliku "
  "proksymalnym",
  "metformina zmniejsza insulinooporność"],
 ["insulina działa przez receptory jądrowe"]),

(49, "CUKRZYCA", "Wskazania",
 "Uzasadnione jest zastosowanie:",
 ["semaglutyd w terapii cukrzycy typu 2", "metformina w leczeniu hirsutyzmu",
  "dapagliflozyna w przewlekłej chorobie nerek",
  "insulina w terapii cukrzycy ciężarnych"],
 ["metformina w cukrzycy typu 1 jako jedyny lek",
  "akarboza w kwasicy ketonowej", "glibenklamid w ciąży"]),

(50, "FARMAKOKINETYKA", "Farmakokinetyka",
 "Dłużej niż wskazywałby na to jego okres półtrwania działa:",
 ["omeprazol", "kwas acetylosalicylowy"],
 ["paracetamol", "dekstrometorfan"]),

(51, "HEMOSTAZA", "Monitorowanie",
 "Monitorowania terapii wymaga stosowanie:",
 ["inhibitora czynnika Xa w niewydolności nerek",
  "heparyny niefrakcjonowanej u osoby otyłej", "acenokumarolu"],
 ["agonisty GPIIb/IIIa"]),

(52, "FARMAKOKINETYKA", "Farmakokinetyka",
 "Prolek:",
 ["może być aktywowany w osoczu"],
 ["nie może być podawany dożylnie",
  "zawsze wymaga prawidłowej funkcji wątroby",
  "ma dłuższy okres półtrwania niż leki, które nie są prolekami"]),

(53, "GLIKOKORTYKOSTEROIDY", "Wskazania",
 "Glikokortykosteroidy w leczeniu astmy oskrzelowej:",
 ["mogą być stosowane zarówno przewlekle, jak i doraźnie",
  "mogą być stosowane z montelukastem",
  "mają potwierdzone zmniejszenie śmiertelności w terapii astmy",
  "powodują up-regulację receptorów beta"],
 ["powodują down-regulację receptorów beta",
  "są przeciwwskazane w skojarzeniu z beta-2-mimetykami",
  "działają wyłącznie doraźnie w napadzie duszności"]),

(54, "LEKI ROŚLINNE I SUPLEMENTY", "Klasyfikacja",
 "Wskaż prawdziwe stwierdzenia dotyczące kategorii produktów:",
 ["insulina aspart jest lekiem biotechnologicznym",
  "20% roztwór albumin jest produktem krwiopochodnym"],
 ["suplement diety jest produktem leczniczym"]),

(55, "UKŁAD ODDECHOWY", "Mechanizm działania",
 "Lekiem, którego głównym mechanizmem działania jest hamowanie "
 "fosfodiesterazy, jest:",
 ["roflumilast", "teofilina", "drotaweryna", "tadalafil"],
 ["montelukast", "salbutamol", "ipratropium"]),

(56, "LEKI PRZECIWPSYCHOTYCZNE", "Działania niepożądane",
 "Konsekwencją blokady receptorów dopaminergicznych jest:",
 ["dystonia", "hiperprolaktynemia", "działanie przeciwwymiotne"],
 ["wzrost masy ciała"]),

(57, "LEKI PRZECIWNOWOTWOROWE", "Mechanizm działania",
 "Przez blokowanie receptorów działa:",
 ["deksametazon"],
 ["cyklofosfamid", "metotreksat", "letrozol"]),

(58, "NARKOTYCZNE LEKI PRZECIWBÓLOWE", "Wskazania",
 "Wskaż prawidłowe stwierdzenie dotyczące opioidów:",
 ["buprenorfina może być stosowana przewlekle",
  "fentanyl dopoliczkowo jest stosowany w leczeniu bólu przebijającego"],
 ["buprenorfina razem z morfiną działa synergistycznie",
  "petydyna powinna być stosowana w postępowaniu okołoporodowym"]),

(59, "LEKI PRZECIWBAKTERYJNE", "Farmakokinetyka",
 "Lek A jest prolekiem, działa bójczo, a jego dawkowanie musi być "
 "modyfikowane w niewydolności nerek. Lekiem A może być:",
 ["oseltamiwir", "metronidazol"],
 ["etambutol", "flukonazol"]),

(60, "NIEOPIOIDOWE LEKI PRZECIWBÓLOWE", "Klasyfikacja",
 "Lek A działa przeciwbólowo, nie działa przeciwzapalnie i nie wykazuje "
 "potencjalnej hepatotoksyczności. Lekiem A jest:",
 ["metamizol", "oksykodon"],
 ["paracetamol", "nimesulid"]),

(61, "LEKI PRZECIWBAKTERYJNE", "Wskazania",
 "W empirycznym leczeniu zakażenia układu moczowego można zastosować:",
 ["fosfomycyna"],
 ["worykonazol", "wankomycyna", "kaspofungina"]),

(62, "LEKI PRZECIWBAKTERYJNE", "Ciąża",
 "U kobiety w ciąży z nadwrażliwością na beta-laktamy w leczeniu zakażenia "
 "dróg oddechowych można zastosować:",
 ["klarytromycyna"],
 ["doksycyklina", "cyprofloksacyna", "amikacyna"]),

(63, "CUKRZYCA", "Działania niepożądane",
 "Wskaż prawidłowe połączenie lek – działanie niepożądane:",
 ["kanagliflozyna – hipotonia", "propylotiouracyl – zapalenie wątroby"],
 ["etynyloestradiol – osteoporoza", "spironolakton – hipernatremia"]),

(65, "LEKI HIPOLIPEMIZUJĄCE", "Działania niepożądane",
 "Pogorszenie lipidogramu może wystąpić w przebiegu leczenia:",
 ["atenolol", "olanzapina", "lopinawir/rytonawir", "etynyloestradiol"],
 ["ezetymib", "rozuwastatyna", "fenofibrat"]),

(67, "HEMOSTAZA", "Antidotum",
 "Rywaroksaban:",
 ["ma swoiste antidotum – andeksanet alfa"],
 ["jest bezpośrednim inhibitorem czynnika IIa",
  "jest prolekiem i wymaga aktywacji przez cytochrom P450",
  "jak wszystkie NOAC jest stosowany w profilaktyce wtórnej ostrego "
  "zespołu wieńcowego"]),

(68, "OTĘPIENIE", "Wskazania",
 "Wskaż prawidłową parę lek – zastosowanie:",
 ["amitryptylina – ból neuropatyczny", "rywastygmina – choroba Alzheimera",
  "selegilina – choroba Parkinsona"],
 ["donepezil – choroba Parkinsona", "lewodopa – choroba Alzheimera",
  "memantyna – ból neuropatyczny"]),

(70, "UKŁAD ODDECHOWY", "Klasyfikacja",
 "Działanie przeciwkaszlowe wykazuje:",
 ["butamirat", "dekstrometorfan"],
 ["ambroksol", "dornaza alfa"]),

(71, "TARCZYCA", "Działania niepożądane",
 "Jatrogenną niedoczynność tarczycy może wywołać:",
 ["węglan litu", "amiodaron"],
 ["etambutol"]),

(72, "ZABURZENIA RYTMU SERCA", "Działania niepożądane",
 "Bradykardię powoduje:",
 ["fentanyl", "iwabradyna", "werapamil"],
 ["salbutamol"]),

(73, "STANY NAGŁE", "Antidotum",
 "Wskaż prawidłowe połączenie trucizna – objawy zatrucia:",
 ["paracetamol – zaburzenia krzepnięcia",
  "glikol etylenowy – kwasica metaboliczna"],
 ["tlenek węgla – duszący kaszel"]),

(76, "FARMAKOKINETYKA", "Farmakokinetyka",
 "W której sytuacji nie będzie efektu terapeutycznego ze względu na "
 "właściwości farmakokinetyczne leku?",
 ["wankomycyna podana doustnie w zapaleniu kości",
  "ramipryl u pacjenta z niewydolnością wątroby"],
 ["dabigatran przy klirensie kreatyniny poniżej 15 ml/min",
  "gentamycyna u szybkich metabolizerów CYP450"]),

(77, "NADCIŚNIENIE TĘTNICZE", "Przeciwwskazania",
 "Wskaż prawdziwe stwierdzenia dotyczące inhibitorów konwertazy "
 "angiotensyny:",
 ["niektóre z nich są prolekami i wymagają metabolizmu do formy aktywnej",
  "u pacjentów odwodnionych mogą wywołać niewydolność nerek",
  "obustronne zwężenie tętnic nerkowych jest przeciwwskazaniem do ich "
  "stosowania"],
 ["są bezpieczne w każdym trymestrze ciąży",
  "powodują hipokaliemię", "są lekami z wyboru w zwężeniu zastawki aortalnej"]),
]

MAX_T, MAX_F = 5, 5
HEADER = (["Section", "Category", "Poziom", "Type", "Question"]
          + [f"True {i}" for i in range(1, MAX_T + 1)]
          + [f"False {i}" for i in range(1, MAX_F + 1)])

POMINIETE = {
    12: "opcja 3 nieczytelna („E…peg alfa”)",
    13: "powtórzenie pyt. 7; opcja 2 sporna (papaweryna czy racekadotryl)",
    28: "brak opcji 3 i 4",
    64: "brak opcji 2 i 3",
    66: "brak opcji 1 i 3",
    69: "brak opcji 2; opcje 3 i 4 merytorycznie sporne",
    74: "brak opcji 1 i 3; obie znane opcje są fałszywe (zero True)",
    75: "znana tylko opcja 1",
    78: "brak opcji 4; pytanie spoza farmakologii (EBM/prawo)",
}


def rows():
    for nr, sec, cat, q, trues, falses in Q:
        assert trues and falses, f"pytanie {nr}"
        assert len(trues) <= MAX_T and len(falses) <= MAX_F, f"pytanie {nr}"
        yield ([sec, cat, "", "testowe", q]
               + trues + [""] * (MAX_T - len(trues))
               + falses + [""] * (MAX_F - len(falses)))


def main():
    data = list(rows())
    nums = [x[0] for x in Q]
    assert nums == sorted(nums) and len(set(nums)) == len(nums)
    assert set(nums) & set(POMINIETE) == set()
    assert len(nums) + len(POMINIETE) == 78, len(nums) + len(POMINIETE)

    base = "farmakologia_egzamin_2022_OnlyPharms"
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
    print("wyeksportowano:", len(data), "z 78 pozycji")
    print("pominieto:", len(POMINIETE), sorted(POMINIETE))
    print("\nSection:")
    for k, v in sorted(Counter(x[1] for x in Q).items()):
        print(f"  {v:3d}  {k}")


if __name__ == "__main__":
    main()
