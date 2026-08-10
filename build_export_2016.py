# -*- coding: utf-8 -*-
"""
Buduje plik eksportowy Only Pharms z egzaminu:
Farmakologia, Stomatologia, lato 2016 (99 pytan - w oryginale brak pyt. 11).

Uklad kolumn zgodny z JAK_KATEGORYZOWAC.txt i z plikami w questions/*.csv:
Section, Category, Poziom, Type, Question, True 1..5, False 1..5
"""
import csv
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, PatternFill
from openpyxl.utils import get_column_letter

# (nr, Section, Category, Question, [True...], [False...])
Q = [
(1, "ZNIECZULENIA MIEJSCOWE", "Farmakokinetyka",
 "Bupiwakaina jest miejscowym środkiem znieczulającym, który ma znacznie silniejsze "
 "i dłuższe działanie niż prokaina. Możliwą przyczyną tej różnicy są poniższe powody "
 "z wyjątkiem:",
 ["zmniejszonej szybkości metabolizmu prokainy w porównaniu z bupiwakainą"],
 ["wyższego współczynnika rozdziału dla bupiwakainy niż prokainy",
  "wyższego wiązania bupiwakainy z albuminami niż prokainy",
  "pKa bupiwakainy jest bardziej zbliżone do pH 7,4"]),

(2, "PADACZKA", "Wskazania",
 "W stanie padaczkowym możesz zastosować:",
 ["diazepam", "tiopental", "fenytoina", "midazolam"],
 ["flumazenil", "etosuksymid", "gabapentyna"]),

(3, "ZNIECZULENIA OGÓLNE", "Mechanizm działania",
 "Znajdź prawdziwe stwierdzenia dotyczące ketaminy:",
 ["czas trwania znieczulenia wynosi 5-10 minut"],
 ["obniża częstość tętna i ciśnienie krwi", "pobudza receptory NMDA",
  "blokuje receptory opioidowe"]),

(4, "UKŁAD PRZYWSPÓŁCZULNY", "Mechanizm działania",
 "Rozszerzenie źrenicy jest typowym objawem działania:",
 ["bieluń (Datura stramonium)", "tropikamid"],
 ["morfina", "muskaryna"]),

(5, "UKŁAD PRZYWSPÓŁCZULNY", "Działania niepożądane",
 "Leki parasympatykomimetyczne:",
 ["są przyczyną krótkowzroczności czynnościowej", "przyspieszają rozwój zaćmy"],
 ["są przyczyną dalekowzroczności czynnościowej", "zmniejszają łzawienie"]),

(6, "CUKRZYCA", "Działania niepożądane",
 "Podczas insulinoterapii może wystąpić:",
 ["zaburzenia widzenia", "lipodystrofia"],
 ["kamica żółciowa", "wzrost stężenia kwasu moczowego we krwi"]),

(7, "UKŁAD POKARMOWY", "Wskazania",
 "W wymiotach pochodzenia błędnikowego skutecznym lekiem jest:",
 ["skopolamina", "dimenhydrynat"],
 ["cyzapryd", "metoklopramid"]),

(8, "NIEOPIOIDOWE LEKI PRZECIWBÓLOWE", "Działania niepożądane",
 "Wskaż właściwe stwierdzenia odnoszące się do metamizolu:",
 ["może być przyczyną nerki analgetycznej",
  "zastosowany w iniekcji dożylnej zagraża wystąpieniem wstrząsu anafilaktycznego",
  "może powodować agranulocytozę", "działa przeciwgorączkowo"],
 ["wykazuje silne działanie przeciwzapalne",
  "jest wybiórczym inhibitorem cyklooksygenazy-2",
  "powoduje hepatotoksyczność zależną od dawki"]),

(9, "NARKOTYCZNE LEKI PRZECIWBÓLOWE", "Farmakokinetyka",
 "Wskaż właściwe stwierdzenia odnoszące się do naloksonu:",
 ["nie wchłania się po podaniu doustnym", "działa krócej niż naltrekson"],
 ["znosi działanie buprenorfiny",
  "w programach metadonowych jest stosowany dożylnie"]),

(10, "UKŁAD POKARMOWY", "Klasyfikacja",
 "Jako lek zapierający stosuje się:",
 ["loperamid", "kodeina", "difenoksylat"],
 ["preparaty z senesu"]),

(12, "LEKI PRZECIWHISTAMINOWE", "Działania niepożądane",
 "Wskaż właściwe stwierdzenia odnoszące się do klemastyny:",
 ["może być przyczyną senności", "ma działanie przeciwświądowe",
  "działa cholinolitycznie",
  "może być podawana jako lek pomocniczy w nagłych odczynach anafilaktycznych"],
 ["nie przenika przez barierę krew-mózg", "jest lekiem II generacji",
  "wykazuje wybiórcze działanie wobec receptora H2"]),

(13, "CUKRZYCA", "Farmakokinetyka",
 "Wskaż właściwe stwierdzenia odnoszące się do metforminy:",
 ["hamuje wchłanianie glukozy i witaminy B12 z przewodu pokarmowego"],
 ["jest stosowana doustnie i dożylnie",
  "jest wskazana zwłaszcza u osób w podeszłym wieku z otyłością",
  "powoduje kwasicę ketonową"]),

(14, "UKŁAD POKARMOWY", "Mechanizm działania",
 "Wskaż prawidłowe połączenie lek – mechanizm działania:",
 ["metoklopramid – antagonista receptorów dopaminergicznych",
  "nalokson – antagonista receptorów opioidowych"],
 ["kodeina – antagonista receptorów opioidowych",
  "dimenhydrynat – agonista receptora H1"]),

(15, "UKŁAD ODDECHOWY", "Wskazania",
 "W leczeniu stanu astmatycznego celowe może być zastosowanie:",
 ["teofilina dożylnie", "bromek ipratropium", "salbutamol", "tlenoterapia"],
 ["propranolol dożylnie", "nieselektywne beta-adrenolityki",
  "leki przeciwkaszlowe zawierające kodeinę"]),

(16, "GLIKOKORTYKOSTEROIDY", "Działania niepożądane",
 "Budezonid stosowany donosowo może być przyczyną:",
 ["krwawienia z nosa", "grzybicy jamy ustnej",
  "wysypki alergicznej na skórze twarzy", "chrypki"],
 ["zespołu Cushinga po podaniu donosowym", "nadciśnienia tętniczego",
  "hiperglikemii po pojedynczej dawce donosowej"]),

(17, "UKŁAD ODDECHOWY", "Klasyfikacja",
 "Działanie przeciwkaszlowe wykazuje:",
 ["pentoksyweryna", "dekstrometorfan"],
 ["ambroksol", "gwajakolosulfonian potasu"]),

(18, "ANTYKONCEPCJA", "Działania niepożądane",
 "Dwuskładnikowe tabletki antykoncepcyjne zwiększają ryzyko:",
 ["chorób zakrzepowo-zatorowych", "zaburzeń nastroju",
  "zaburzeń profilu lipidowego surowicy"],
 ["osteoporozy"]),

(19, "UKŁAD POKARMOWY", "Farmakokinetyka",
 "Wskaż właściwe stwierdzenia odnoszące się do difenoksylatu:",
 ["dobrze wchłania się z przewodu pokarmowego",
  "może być przyczyną zapalenia trzustki",
  "wykazuje powinowactwo do receptorów opioidowych"],
 ["stosowany jest rutynowo w leczeniu biegunek wywołanych bakteriami "
  "wytwarzającymi endotoksyny"]),

(20, "NARKOTYCZNE LEKI PRZECIWBÓLOWE", "Przeciwwskazania",
 "Przeciwwskazaniem do podania morfiny jest:",
 ["ostra porfiria wątrobowa"],
 ["ostry obrzęk płuc", "zawał mięśnia sercowego z uniesieniem odcinka ST",
  "ból nowotworowy"]),

(21, "NARKOTYCZNE LEKI PRZECIWBÓLOWE", "Działania niepożądane",
 "Do efektów ośrodkowych opioidowych leków przeciwbólowych należy:",
 ["mioza", "wymioty", "euforia", "depresja ośrodka oddechowego"],
 ["zaparcie", "skurcz zwieracza Oddiego", "zatrzymanie moczu"]),

(22, "NARKOTYCZNE LEKI PRZECIWBÓLOWE", "Mechanizm działania",
 "Kodeina wykazuje działanie:",
 ["przeciwbólowe", "przeciwkaszlowe"],
 ["przeciwgorączkowe", "przeciwzapalne"]),

(23, "FARMAKOKINETYKA", "Farmakokinetyka",
 "Silny efekt pierwszego przejścia nie występuje przy podaniu leku:",
 ["transdermalnym", "w inhalacjach", "doodbytniczym", "podjęzykowym"],
 ["doustnym", "dożołądkowym przez sondę",
  "doustnym w postaci tabletki o przedłużonym uwalnianiu"]),

(24, "LEKI PRZECIWBAKTERYJNE", "Wskazania",
 "Na metycylinooporne szczepy Staphylococcus aureus (MRSA) działa:",
 ["wankomycyna", "tygecyklina", "daptomycyna", "linezolid"],
 ["kloksacylina", "cefazolina", "amoksycylina z kwasem klawulanowym"]),

(25, "LEKI PRZECIWLĘKOWE", "Wskazania",
 "W leczeniu zespołu abstynencyjnego po odstawieniu alkoholu zastosujesz:",
 ["diazepam"],
 ["flumazenil", "atenolol", "disulfiram"]),

(26, "PADACZKA", "Działania niepożądane",
 "Wskaż właściwe stwierdzenia odnoszące się do lewetyracetamu:",
 ["powoduje zaburzenia zachowania"],
 ["silnie hamuje prąd wapniowy",
  "stosowany jest w chorobie afektywnej dwubiegunowej",
  "jest inhibitorem wychwytu kwasu gamma-aminomasłowego"]),

(27, "FARMAKOKINETYKA", "Farmakokinetyka",
 "Biodostępność leku zależy od:",
 ["stopnia jonizacji leku", "rozpuszczalności w tłuszczach",
  "wielkości cząsteczki"],
 ["wiązania leku z albuminami"]),

(28, "FARMAKOKINETYKA", "Interakcje",
 "Działanie hamujące enzymy cytochromu P-450 wykazuje:",
 ["ketokonazol", "erytromycyna", "sok z grejpfruta"],
 ["spironolakton"]),

(29, "UKŁAD PRZYWSPÓŁCZULNY", "Wskazania",
 "Inhibitory acetylocholinesterazy możesz zastosować w:",
 ["chorobie Alzheimera", "nużliwości mięśni"],
 ["astmie oskrzelowej", "zatruciu glikozydami naparstnicy"]),

(30, "UKŁAD PRZYWSPÓŁCZULNY", "Wskazania",
 "W nietrzymaniu moczu możesz zastosować:",
 ["tolterodyna", "darifenacyna", "oksybutynina"],
 ["fenoksybenzamina"]),

(31, "LEKI PRZECIWBAKTERYJNE", "Wskazania",
 "Działanie przeciwpneumokokowe (p-Streptococcus pneumoniae) ma:",
 ["amoksycylina", "ceftriakson", "klindamycyna"],
 ["cyprofloksacyna"]),

(32, "UKŁAD WSPÓŁCZULNY", "Wskazania",
 "U pacjenta z zaawansowaną hipotonią ortostatyczną możesz zastosować:",
 ["fenylefryna", "metoksamina"],
 ["fentolamina", "duloksetyna"]),

(33, "ZNIECZULENIA OGÓLNE", "Działania niepożądane",
 "Wskaż właściwe stwierdzenia odnoszące się do propofolu:",
 ["powoduje obniżenie ciśnienia tętniczego",
  "może być podawany pacjentom z padaczką",
  "działa depresyjnie na układ oddechowy",
  "jest stosowany w celu indukcji znieczulenia"],
 ["wykazuje silne działanie przeciwbólowe",
  "powoduje wzrost ciśnienia tętniczego",
  "jest antagonistą receptorów NMDA"]),

(34, "UKŁAD WSPÓŁCZULNY", "Mechanizm działania",
 "Działanie hamujące w stosunku do receptorów alfa- i beta-adrenergicznych wykazuje:",
 ["karwedilol", "labetalol"],
 ["metoprolol", "tymolol"]),

(35, "CHOROBA NIEDOKRWIENNA SERCA", "Mechanizm działania",
 "Działanie naczyniorozszerzające związane z uwalnianiem tlenku azotu wykazuje:",
 ["diazotan izosorbidu", "molsidomina", "nebiwolol", "nitroprusydek sodu"],
 ["amlodypina", "werapamil", "doksazosyna"]),

(36, "NADCIŚNIENIE TĘTNICZE", "Mechanizm działania",
 "Działanie naczyniorozszerzające wykazuje:",
 ["dihydralazyna", "diltiazem", "indapamid", "amlodypina"],
 ["propranolol", "efedryna", "fenylefryna"]),

(37, "CHOROBA NIEDOKRWIENNA SERCA", "Wskazania",
 "Celem przerwania ostrego bólu dławicowego możesz zastosować:",
 ["nitrogliceryna"],
 ["dihydroergotamina", "iwabradyna", "moksonidyna"]),

(38, "LEKI PRZECIWDEPRESYJNE", "Klasyfikacja",
 "Do leków stabilizujących nastrój należy:",
 ["kwas walproinowy", "karbamazepina", "lamotrygina", "węglan litu"],
 ["fluoksetyna", "lewetyracetam", "gabapentyna"]),

(39, "NIEWYDOLNOŚĆ SERCA", "Mechanizm działania",
 "Działanie inotropowe dodatnie wykazuje:",
 ["milrinon", "digoksyna", "glukagon", "dopamina"],
 ["metoprolol", "werapamil", "diltiazem"]),

(40, "LEKI MOCZOPĘDNE", "Działania niepożądane",
 "Podczas leczenia klopamidem może wystąpić:",
 ["hipomagnezemia", "hiponatremia", "hiperglikemia"],
 ["hipokalcemia"]),

(41, "LEKI NASENNE", "Mechanizm działania",
 "Działanie pobudzające receptory GABA-A wykazuje:",
 ["zolpidem", "midazolam"],
 ["klozapina", "buspiron"]),

(42, "ZABURZENIA RYTMU SERCA", "Mechanizm działania",
 "Podaj prawidłowe połączenie lek – mechanizm działania:",
 ["dorzolamid – hamowanie anhydrazy węglanowej",
  "adenozyna – zahamowanie prądu wapniowego"],
 ["chlortalidon – zahamowanie szybkiego prądu sodowego",
  "amiodaron – pobudzanie kanałów potasowych"]),

(43, "ZABURZENIA RYTMU SERCA", "Wskazania",
 "Celem przywrócenia rytmu zatokowego u pacjenta z migotaniem przedsionków "
 "możesz zastosować:",
 ["metoprolol", "digoksyna"],
 ["triamteren", "fenytoina"]),

(44, "ZNIECZULENIA OGÓLNE", "Farmakokinetyka",
 "Wskaż właściwe stwierdzenia odnoszące się do atrakurium:",
 ["jest metabolizowane przez esterazy osoczowe", "może powodować drgawki",
  "jest niedepolaryzującym środkiem zwiotczającym"],
 ["działa zwiotczająco około 1 minutę"]),

(45, "LEKI PRZECIWBAKTERYJNE", "Wskazania",
 "W pozaszpitalnym (domowym) zapaleniu płuc zastosujesz:",
 ["aksetyl cefuroksymu", "klarytromycyna", "lewofloksacyna"],
 ["kotrimoksazol"]),

(46, "LEKI MOCZOPĘDNE", "Działania niepożądane",
 "Hiperurykemię możesz obserwować podczas leczenia:",
 ["torasemid", "hydrochlorotiazyd"],
 ["spironolakton", "mannitol"]),

(47, "LEKI PRZECIWGRZYBICZE", "Mechanizm działania",
 "Epoksydazę skwalenową hamuje:",
 ["terbinafina"],
 ["ketokonazol", "nystatyna", "amorolfina"]),

(48, "LEKI PRZECIWBAKTERYJNE", "Działania niepożądane",
 "Działanie ototoksyczne wykazuje:",
 ["wankomycyna", "furosemid", "amikacyna"],
 ["etambutol"]),

(49, "BÓLE GŁOWY", "Wskazania",
 "Celem przerwania migrenowego bólu głowy możesz zastosować:",
 ["kwas acetylosalicylowy", "ibuprofen", "sumatryptan"],
 ["buprenorfina"]),

(50, "NADCIŚNIENIE TĘTNICZE", "Działania niepożądane",
 "Wskaż prawidłowe połączenie lek – działanie niepożądane:",
 ["enalapryl – zaburzenia smaku", "salbutamol – drżenia mięśniowe",
  "atorwastatyna – bóle mięśniowe", "akarboza – wzdęcia brzucha"],
 ["metformina – kwasica ketonowa", "paracetamol – agranulocytoza",
  "ibuprofen – methemoglobinemia"]),

(51, "LEKI PRZECIWBAKTERYJNE", "Klasyfikacja",
 "Do leków I rzutu w terapii gruźlicy należy:",
 ["izoniazyd", "ryfampicyna", "pirazynamid"],
 ["moksyfloksacyna"]),

(52, "HEMOSTAZA", "Działania niepożądane",
 "Heparyna może powodować:",
 ["łysienie", "osteoporozę", "niedobór mineralokortykosteroidów",
  "trombocytopenię"],
 ["hiperaldosteronizm", "hipokaliemię", "methemoglobinemię"]),

(53, "LEKI PRZECIWGRZYBICZE", "Wskazania",
 "W inwazyjnej kandydozie skuteczne jest zastosowanie:",
 ["worykonazol", "dezoksycholan amfoterycyny B"],
 ["ekonazol", "nystatyna"]),

(54, "TARCZYCA", "Mechanizm działania",
 "Wskaż prawidłowe połączenie lek – mechanizm działania:",
 ["budezonid – zmiana ekspresji genów kodujących białka prozapalne",
  "tiamazol – hamowanie syntezy T3 i T4 poprzez blokowanie aktywności jodazy"],
 ["ranitydyna – blokowanie przekaźnictwa serotoninergicznego",
  "dimenhydrynat – pobudzenie receptorów histaminowych"]),

(55, "LEKI PRZECIWPASOŻYTNICZE", "Wskazania",
 "Do usuwania wszy ludzkiej z owłosionej skóry głowy stosuje się preparaty zawierające:",
 ["permetryna"],
 ["kwas borny", "oktenidyna", "kwas octowy stężony"]),

(56, "NADCIŚNIENIE TĘTNICZE", "Działania niepożądane",
 "Wartości ciśnienia tętniczego krwi może zwiększyć:",
 ["piroksykam", "deksametazon"],
 ["eplerenon", "morfina"]),

(57, "CUKRZYCA", "Klasyfikacja",
 "Wskaż właściwe stwierdzenia odnoszące się do insuliny izofanowej (NPH):",
 ["jest stosowana często na noc"],
 ["to insulina długodziałająca", "to analog długodziałający",
  "jest stosowana przed wszystkimi głównymi posiłkami"]),

(58, "LEKI PRZECIWBAKTERYJNE", "Farmakokinetyka",
 "Wskaż właściwe stwierdzenia odnoszące się do netilmycyny:",
 ["posiada efekt poantybiotykowy", "wykazuje synergizm z beta-laktamami"],
 ["dobrze wchłania się z przewodu pokarmowego",
  "jest przeciwwskazana u chorych z niewydolnością wątroby"]),

(59, "LEKI PRZECIWBAKTERYJNE", "Działania niepożądane",
 "Wybierz prawidłowo dobrane pary lek – możliwe działanie niepożądane:",
 ["imipenem – drgawki", "ceftriakson – zakrzepowe zapalenie żył",
  "cefadroksyl – wysypka skórna", "penicylina V – wstrząs anafilaktyczny"],
 ["wankomycyna – methemoglobinemia", "gentamycyna – przerost dziąseł",
  "amoksycylina – hipertermia złośliwa"]),

(60, "LEKI PRZECIWBAKTERYJNE", "Wskazania",
 "Erytromycyna jest aktywna wobec:",
 ["Mycoplasma pneumoniae", "Legionella pneumophila", "Chlamydia trachomatis",
  "Listeria monocytogenes"],
 ["Pseudomonas aeruginosa", "Escherichia coli", "Bacteroides fragilis"]),

(62, "CUKRZYCA", "Mechanizm działania",
 "Wskaż właściwe stwierdzenia odnoszące się do pramlintydu:",
 ["redukuje glikemię poposiłkową", "hamuje sekrecję glukagonu",
  "zwalnia opróżnianie żołądka"],
 ["zwiększa apetyt"]),

(63, "FARMAKOKINETYKA", "Mechanizm działania",
 "Inhibitor allosteryczny:",
 ["zmniejsza odpowiedź po związaniu się agonisty z receptorem"],
 ["uniemożliwia połączenie się agonisty z receptorem",
  "zmienia strukturę III-rzędową agonisty",
  "wiąże się z agonistą i neutralizuje jego działanie"]),

(64, "FARMAKOKINETYKA", "Farmakokinetyka",
 "Leki prekursorowe są aktywowane do postaci aktywnych:",
 ["w wątrobie", "w osoczu", "w nabłonku jelita"],
 ["w śledzionie"]),

(65, "FARMAKOKINETYKA", "Farmakokinetyka",
 "Duża objętość dystrybucji leku może sugerować:",
 ["istnienie kompartmentu głębokiego, czyli narządu kumulującego lek",
  "małą szybkość wydalania", "dużą lipofilność"],
 ["dużą zawartość wody w organizmie"]),

(66, "LEKI PRZECIWDEPRESYJNE", "Działania niepożądane",
 "Węglan litu powoduje:",
 ["drżenie", "zaburzenia czynności tarczycy", "polidypsję", "obrzęki"],
 ["agranulocytozę", "zwłóknienie płuc", "methemoglobinemię"]),

(67, "HEMOSTAZA", "Mechanizm działania",
 "Wskaż właściwe stwierdzenia odnoszące się do prasugrelu:",
 ["jest prolekiem", "wiąże się trwale z receptorem P2Y12 na płytkach krwi",
  "może być podawany łącznie z kwasem acetylosalicylowym"],
 ["jest podawany w profilaktyce migotania przedsionków u pacjentów "
  "z przeciwwskazaniem do stosowania antagonistów witaminy K"]),

(68, "HEMOSTAZA", "Mechanizm działania",
 "Wskaż właściwe stwierdzenia odnoszące się do eteksylanu dabigatranu:",
 ["swoiście i odwracalnie hamuje trombinę", "jest lekiem prekursorowym",
  "jest skuteczny w zapobieganiu powikłaniom zakrzepowym migotania przedsionków"],
 ["podczas stosowania wymaga monitorowania układu krzepnięcia"]),

(69, "LEKI PRZECIWBAKTERYJNE", "Sposób podania",
 "Wskaż właściwe stwierdzenia odnoszące się do chlorheksydyny:",
 ["stosowana jest w wodnych preparatach w postaci diglukonianu chlorheksydyny",
  "ma niewielkie działanie uczulające lub drażniące skórę",
  "jest przeciwwskazana podczas operacji w uchu środkowym"],
 ["jest odporna na działanie środków powierzchniowo czynnych"]),

(70, "LEKI PRZECIWBAKTERYJNE", "Interakcje",
 "Reakcję disulfiramową może wywołać:",
 ["cefamandol", "metronidazol"],
 ["oksazepam", "metformina"]),

# --- pytania jednokrotnego wyboru (71-100) ---

(71, "CHOROBA PARKINSONA", "Działania niepożądane",
 "Wady zastawek serca mogą wystąpić podczas leczenia:",
 ["pergolid"],
 ["selegilina", "entakapon", "amantadyna", "procyklidyna"]),

(72, "PADACZKA", "Działania niepożądane",
 "Ubytki w polu widzenia mogą wystąpić podczas leczenia:",
 ["wigabatryna"],
 ["zonisamid", "gabapentyna", "pregabalina", "lamotrygina"]),

(73, "UKŁAD ODDECHOWY", "Mechanizm działania",
 "Omalizumab:",
 ["jest przeciwciałem monoklonalnym anty-IgE"],
 ["hamuje fosfodiesterazę typu 2", "jest agonistą receptorów beta-2",
  "hamuje fosfolipazę A2", "działa hamująco na receptory dla leukotrienów"]),

(74, "UKŁAD ODDECHOWY", "Mechanizm działania",
 "Antagonistą receptora adenozynowego A2 jest:",
 ["teofilina"],
 ["meksyletyna", "sotalol", "mianseryna", "finasteryd"]),

(75, "LEKI PRZECIWDEPRESYJNE", "Mechanizm działania",
 "Antagonistą receptora 5-HT2 jest:",
 ["trazodon"],
 ["mirtazapina", "moklobemid", "dezypramina", "wenlafaksyna"]),

(76, "LEKI PRZECIWDEPRESYJNE", "Wskazania",
 "W leczeniu uzależnienia od nikotyny możesz zastosować:",
 ["bupropion"],
 ["duloksetyna", "ipsapiron", "tropisetron", "zileuton"]),

(77, "NADCIŚNIENIE TĘTNICZE", "Działania niepożądane",
 "Silny kaszel może wystąpić podczas leczenia:",
 ["chinapryl"],
 ["walsartan", "jodek potasu", "minoksydyl", "eplerenon"]),

(78, "UKŁAD PRZYWSPÓŁCZULNY", "Wskazania",
 "W celu rozszerzenia źrenicy celem diagnostyki zmian na dnie oka możesz zastosować:",
 ["tropikamid"],
 ["pirazynamid", "topiramat", "pirydostygmina", "donepezil"]),

(79, "LEKI PRZECIWPSYCHOTYCZNE", "Działania niepożądane",
 "U pacjenta leczonego przez wiele lat z powodu schizofrenii występują działania "
 "niepożądane: akinezja, sztywność i drżenie. Są one wynikiem:",
 ["działania przeciwdopaminergicznego"],
 ["działania przeciwcholinergicznego",
  "obniżenia ekspresji kwasu gamma-aminomasłowego (GABA)",
  "hamowania wychwytu zwrotnego serotoniny", "działania przeciwserotoninowego"]),

(80, "NIEOPIOIDOWE LEKI PRZECIWBÓLOWE", "Działania niepożądane",
 "Pacjentka zostaje przyjęta do szpitala po spożyciu co najmniej 30 tabletek "
 "nieznanego leku. Wstępne badanie fizykalne nie wykazuje żadnych nieprawidłowości. "
 "Trzydzieści sześć godzin później aktywność ASPAT w surowicy wynosi 1500 U/L, "
 "a ALAT 2000 U/L. Jaki to lek?",
 ["paracetamol"],
 ["kodeina", "ebastyna", "diazepam", "salbutamol"]),

(81, "UKŁAD POKARMOWY", "Mechanizm działania",
 "Omeprazol:",
 ["hamuje aktywność H+/K+-ATPazy"],
 ["jest antagonistą receptorów dla gastryny", "jest antagonistą receptorów H2",
  "jest antagonistą receptorów M3", "jest analogiem prostaglandyny E"]),

(82, "LEKI PRZECIWBAKTERYJNE", "Interakcje",
 "Pacjentka przychodzi do lekarza z powodu nudności, wymiotów i bólów brzucha "
 "godzinę po spożyciu kieliszka wina do kolacji. Trzy dni temu zaczęła leczenie "
 "zakażenia rzęsistkiem pochwowym. Jaki lek przyjmuje pacjentka?",
 ["metronidazol"],
 ["ceftriakson", "chlorochina", "doksycyklina", "mebendazol"]),

(83, "LEKI PRZECIWBAKTERYJNE", "Mechanizm działania",
 "U pacjenta z zakażeniem HIV i zapaleniem płuc wywołanym przez Pneumocystis "
 "jirovecii zastosowano skuteczne leczenie trimetoprimem-sulfametoksazolem "
 "dzięki hamowaniu aktywności:",
 ["reduktazy dihydrofolianowej"],
 ["wbudowania steroli do błon komórkowych", "kwasu p-aminobenzoesowego",
  "topoizomerazy II", "syntezy białka PBP"]),

(84, "HEMOSTAZA", "Antidotum",
 "Odwrócenie efektu przeciwzakrzepowego heparyny uzyskasz po podaniu:",
 ["siarczan protaminy"],
 ["heparynaza", "kwas traneksamowy", "witamina K", "antytrombina III"]),

(85, "LEKI PRZECIWDEPRESYJNE", "Działania niepożądane",
 "Przedawkowanie trójpierścieniowych leków przeciwdepresyjnych może doprowadzić "
 "do zgonu pacjenta w wyniku:",
 ["zaburzeń rytmu serca"],
 ["depresji ośrodka oddechowego", "śpiączki", "hipertermii",
  "ostrej niewydolności nerek"]),

(86, "LEKI PRZECIWDEPRESYJNE", "Mechanizm działania",
 "Mechanizm działania litu w psychozie maniakalno-depresyjnej jest związany z:",
 ["hamowaniem przemian fosfatydyloinozytoli w komórce"],
 ["hamowaniem wychwytu zwrotnego noradrenaliny",
  "hamowaniem wychwytu zwrotnego serotoniny",
  "działaniem antagonistycznym na receptory dopaminergiczne D2",
  "działaniem cholinolitycznym"]),

(87, "FARMAKOKINETYKA", "Farmakokinetyka",
 "Indeks terapeutyczny jest miarą:",
 ["bezpieczeństwa"],
 ["siły działania leku", "selektywności", "skuteczności",
  "stopnia wiązania z białkami krwi"]),

(88, "ZNIECZULENIA OGÓLNE", "Działania niepożądane",
 "Z którym z następujących działań niepożądanych jest związane działanie ketaminy?",
 ["omamy i koszmary senne"],
 ["podrażnienie dróg oddechowych", "obniżenie objętości wyrzutowej serca",
  "hipertermia złośliwa", "rabdomioliza"]),

(89, "UKŁAD PRZYWSPÓŁCZULNY", "Antidotum",
 "Pacjent z łzawieniem, wymiotami i biegunką, puls 45/min, ciśnienie tętnicze "
 "90/60 mm Hg. Objawy mogą wskazywać na zatrucie:",
 ["związkami fosforoorganicznymi"],
 ["metanolem", "toksyną botulinową", "glikolem etylenowym", "digoksyną"]),

(90, "LEKI PRZECIWGRZYBICZE", "Wskazania",
 "U pacjentki leczonej z powodu ostrej białaczki limfoblastycznej stwierdzono "
 "powikłanie – obecność Candida albicans w surowicy. Który z poniższych leków "
 "przeciwgrzybiczych należy zastosować?",
 ["amfoterycyna B"],
 ["nystatyna", "klotrimazol", "ketokonazol", "kapsaicyna"]),

(91, "LEKI PRZECIWGRZYBICZE", "Wskazania",
 "U pacjenta z AIDS rozwija się gorączka, ból szyi i światłowstręt. Badanie płynu "
 "mózgowo-rdzeniowego ujawnia Cryptococcus neoformans. Który z leków zastosować?",
 ["flukonazol"],
 ["flucytozyna", "gryzeofulwina", "cykloseryna", "kotrimoksazol"]),

(92, "LEKI PRZECIWPASOŻYTNICZE", "Wskazania",
 "Jaki lek przeciwmalaryczny należy zastosować profilaktycznie w przypadku podróży "
 "do rejonu z opornością na chlorochinę?",
 ["meflochina"],
 ["prymachina", "doksycyklina", "pirymetamina", "ryfampicyna"]),

(93, "LEKI PRZECIWPASOŻYTNICZE", "Wskazania",
 "W zakażeniu Giardia lamblia możesz zastosować:",
 ["metronidazol"],
 ["nifurtimoks", "mebendazol", "suramina", "pyrantel"]),

(94, "LEKI PRZECIWPASOŻYTNICZE", "Wskazania",
 "W leczeniu zakażenia Toxoplasma gondii jest stosowana:",
 ["pirymetamina"],
 ["iwermektyna", "prazykwantel", "niklozamid", "tiabendazol"]),

(95, "UKŁAD POKARMOWY", "Wskazania",
 "Pacjent z podwyższonym stężeniem w surowicy wazoaktywnego peptydu jelitowego "
 "w związku z nowotworem trzustki (VIP-oma) powinien otrzymać:",
 ["oktreotyd"],
 ["gastryna", "glukagon", "sulfasalazyna", "famotydyna"]),

(96, "UKŁAD POKARMOWY", "Wskazania",
 "U pacjenta z rozpoznaniem zespołu Zollingera-Ellisona możesz zastosować:",
 ["lansoprazol"],
 ["famotydyna", "mizoprostol", "propantelina",
  "cytrynian potasowo-bizmutawy"]),

(97, "NADCIŚNIENIE TĘTNICZE", "Wskazania",
 "U pacjenta z nadciśnieniem tętniczym i współistniejącym białkomoczem celowe "
 "jest zastosowanie:",
 ["peryndopryl"],
 ["nikardypina", "nitrogliceryna", "hydrochlorotiazyd", "moksonidyna"]),

(98, "PADACZKA", "Działania niepożądane",
 "Hepatotoksyczne działanie wykazuje:",
 ["kwas walproinowy"],
 ["cetyryzyna", "mupirocyna", "penicylina krystaliczna", "bromheksyna"]),

(99, "STANY NAGŁE", "Wskazania",
 "U 60-letniego chorego z obrzękiem płuc i ciśnieniem krwi 190/110 mmHg zastosujesz:",
 ["nitrogliceryna we wlewie dożylnym"],
 ["nifedypina podjęzykowo", "kaptopryl doustnie", "klonidyna doustnie",
  "amiodaron dożylnie"]),

(100, "UKŁAD WSPÓŁCZULNY", "Działania niepożądane",
 "Ksylometazolina:",
 ["może doprowadzić do wystąpienia tzw. „efektu z odbicia”"],
 ["jest antagonistą receptorów alfa",
  "skuteczność kliniczną wykazuje po 5 dniach terapii",
  "jest podawana doustnie", "dobrze penetruje przez barierę krew-mózg"]),
]

MAX_T, MAX_F = 5, 5
HEADER = (["Section", "Category", "Poziom", "Type", "Question"]
          + [f"True {i}" for i in range(1, MAX_T + 1)]
          + [f"False {i}" for i in range(1, MAX_F + 1)])


def rows():
    for nr, sec, cat, q, trues, falses in Q:
        assert trues, f"pytanie {nr}: brak True"
        assert falses, f"pytanie {nr}: brak False"
        assert len(trues) <= MAX_T, f"pytanie {nr}: za duzo True"
        assert len(falses) <= MAX_F, f"pytanie {nr}: za duzo False"
        yield ([sec, cat, "", "testowe", q]
               + trues + [""] * (MAX_T - len(trues))
               + falses + [""] * (MAX_F - len(falses)))


def main():
    data = list(rows())
    nums = [x[0] for x in Q]
    assert nums == sorted(nums) and len(set(nums)) == len(nums)
    assert 11 not in nums, "pyt. 11 nie istnieje w oryginale"
    assert len(data) == 98, len(data)

    base = "farmakologia_egzamin_2016_OnlyPharms"
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

    head_fill = PatternFill("solid", fgColor="2F5597")
    for c in ws[1]:
        c.font = Font(bold=True, color="FFFFFF")
        c.fill = head_fill
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
    print("wierszy:", len(data), " (brak pyt. 11 w PDF; pyt. 61 usuniete jako duplikat)")
    print("Trues:", sum(len(t) for *_, t, _ in Q),
          " Falses:", sum(len(f) for *_, f in Q))
    print("\nSection:")
    for k, v in sorted(Counter(x[1] for x in Q).items()):
        print(f"  {v:3d}  {k}")
    print("\nCategory:")
    for k, v in sorted(Counter(x[2] for x in Q).items()):
        print(f"  {v:3d}  {k}")


if __name__ == "__main__":
    main()
