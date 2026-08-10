# -*- coding: utf-8 -*-
"""Automatyczna klasyfikacja pytan ARCYBAZY na Section/Category Only Pharms."""
import re, unicodedata


def nrm(s):
    s = unicodedata.normalize('NFKD', s.lower()).encode('ascii', 'ignore').decode()
    return re.sub(r'\s+', ' ', s)


# Section -> lista wzorcow (rdzenie slow, bez polskich znakow)
SECTIONS = [
 ("LEKI PRZECIWBAKTERYJNE", r"""antybiotyk|penicylin|amoksycylin|ampicylin|kloksacylin|piperacylin|
   cefalospor|cefazolin|cefuroksym|ceftriakson|cefotaksym|ceftazydym|cefepim|cefadroksyl|cefamandol|
   karbapenem|imipenem|meropenem|ertapenem|aztreonam|wankomycyn|teikoplanin|dalbawancyn|
   aminoglikozyd|gentamycyn|amikacyn|netilmycyn|streptomycyn|tobramycyn|spektynomycyn|
   makrolid|erytromycyn|klarytromycyn|azytromycyn|spiramycyn|
   tetracyklin|doksycyklin|limecyklin|tygecyklin|minocyklin|
   fluorochinolon|chinolon|cyprofloksacyn|lewofloksacyn|moksyfloksacyn|norfloksacyn|ofloksacyn|
   linezolid|daptomycyn|kolistyn|polimyksyn|klindamycyn|linkozamid|metronidazol|tynidazol|
   kotrimoksazol|trimetoprim|sulfametoksazol|sulfonamid|sulfasalazyn|fosfomycyn|nitrofurantoin|
   ryfampicyn|ryfaksymin|rifaksymin|izoniazyd|pyrazynamid|pirazynamid|etambutol|etionamid|gruzlic|pratk|
   chloramfenikol|fidaksomycyn|mupirocyn|kwas fusydowy|bacytracyn|gramicydyn|chinuprystyn|dalfoprystyn|
   mrsa|vre|esbl|mssa|beztlenowc|gronkowc|paciorkowc|pneumokok|enterokok|
   staphylococ|streptococ|klebsiella|pseudomonas|escherichia|proteus|bacteroides|clostridium|
   legionella|chlamydia|mycoplasma|bordetella|listeria|actinomyces|propionibacterium|helicobacter|
   antyseptyk|chlorheksydyn|podchloryn|poliwidon|oktenidyn|nadmanganian|azotan srebra|
   rzezaczk|kila |tezec|promienic|czyracz|tradzik|zakazen|posiew|antybiotykoterapi|antybiogram|
   farmakoprofilaktyk|bakteriuri|eradykacj|zapalenie pluc|angin|osteomyelit|zapalenie wsierdzi|izw"""),
 ("LEKI PRZECIWGRZYBICZE", r"""grzybic|przeciwgrzybicz|amfoterycyn|flukonazol|itrakonazol|worykonazol|
   pozakonazol|ketokonazol|klotrimazol|mikonazol|ekonazol|amorolfin|terbinafin|nystatyn|natamycyn|
   flucytozyn|kaspofungin|mikafungin|anidulafungin|gryzeofulwin|candida|aspergillus|cryptococc|
   histoplasma|coccidioides|dermatofit|epidermophyton|epoksydaz skwalenow|kandydoz"""),
 ("LEKI PRZECIWWIRUSOWE", r"""przeciwwirusow|acyklowir|walacyklowir|gancyklowir|walgancyklowir|
   oseltamiwir|zanamiwir|rymantadyn|amantadyn(?!.*parkinson)|rybawiryn|remdesiwir|remdesywir|
   molnupirawir|abakawir|lamiwudyn|zydowudyn|tenofowir|rytonawir|lopinawir|interferon|
   opryszczk|gryp[ay]|hiv|wzw|covid|wirus"""),
 ("LEKI PRZECIWPASOŻYTNICZE", r"""pasozyt|przeciwpasozytnicz|robaczyc|tasiemc|glist|owsic|
   albendazol|mebendazol|prazykwantel|niklozamid|niklosamid|pyrantel|iwermektyn|permetryn|
   tiabendazol|prymachin|meflochin|chlorochin|chinin|proguanil|pirymetamin|atowakwon|
   malari|giardia|entamoeba|toxoplasma|rzesistek|wszawic|swierzb|ascaris|pentamidyn|suramin|
   nifurtimoks|prazikwantel"""),
 ("LEKI PRZECIWNOWOTWOROWE", r"""przeciwnowotworow|cytostatyk|chemioterapi|cytotoksyczn|
   cyklofosfamid|metotreksat|doksorubicyn|cisplatyn|winkrystyn|winblastyn|paklitaksel|fluorouracyl|
   letrozol|anastrozol|tamoksyfen|imatynib|rytuksymab|trastuzumab|pembrolizumab|niwolumab|
   filgrastym|talidomid|bleomycyn|etopozyd|bialaczk|nowotwor|onkolog"""),
 ("HEMOSTAZA", r"""heparyn|enoksaparyn|dalteparyn|nadroparyn|fondaparynuks|warfaryn|acenokumarol|
   antagonist(a|ow) witaminy k|dabigatran|rywaroksaban|apiksaban|edoksaban|noac|
   klopidogrel|prasugrel|tikagrelor|tiklopidyn|abciksimab|eptifibatyd|tirofiban|
   aktywator(a|em)? plazminogenu|t.pa|cilostazol|pentoksyfilin|naftydrofuryl|natydrofuryl|diosmin|chromani przestankow|alteplaz|streptokinaz|urokinaz|tenekteplaz|trombolit|fibrynolit|
   kwas traneksamowy|aminokapronow|protamin|andeksanet|idarucyzumab|
   przeciwkrzepliw|antyagregacyj|agregacj plytek|krzepnie|zakrzep|krwawien|inr|aptt|
   czynnik xa|trombin|p2y12|plytk"""),
 ("NADCIŚNIENIE TĘTNICZE", r"""nadcisnien|hipotensyj|cisnienia tetnicz|przelom nadcisnien|
   enalapryl|ramipryl|peryndopryl|perindopr|lizynopryl|kaptopryl|chinapryl|trandolapryl|benazepryl|
   inhibitor(y|ow)? konwertazy|acei|inhibitorow ace|
   losartan|walsartan|telmisartan|kandesartan|irbesartan|sartan|
   amlodypin|nitrendypin|felodypin|lacydypin|nikardypin|nifedypin|lerkanidypin|
   doksazosyn|prazosyn|terazosyn|urapidil|urapidyl|klonidyn|moksonidyn|metyldopa|guanetydyn|
   kanal(ow|y)? wapniow|antagonisc(i|ow) wapni|epoprostenol|riocyguat|midodryn|aliskiren|rezerpin|dihydralazyn|hydralazyn|nitroprusydek|labetalol|
   indapamid|chlortalidon|hydrochlorotiazyd|klopamid|tiazyd|
   bialkomocz|nefropati cukrzycow"""),
 ("NIEWYDOLNOŚĆ SERCA", r"""niewydolnosc serca|niewydolnosci serca|frakcj wyrzutow|hfref|
   digoksyn|glikozyd|naparstnic|milrinon|lewosimendan|dobutamin|inotropow|
   sakubitryl|iwabradyn(?!.*bradykard)|eplerenon|spironolakton"""),
 ("CHOROBA NIEDOKRWIENNA SERCA", r"""choroba wiencow|choroby wiencow|dlawic|dusznic|
   niedokrwienn(a|ej) serca|chns|ozw|zawal|stemi|nstemi|prinzmetal|
   nitroglicery|azotan(y|ow|u|ach)?( organiczn)?|wazodylatacyj|diazotan izosorbidu|monoazotan|molsidomin|ranolazyn|
   trimetazydyn|blaszk miazdzycow"""),
 ("ZABURZENIA RYTMU SERCA", r"""antyarytmiczn|przeciwarytmiczn|arytmi|migotanie przedsionk|
   rytmu zatokow|odstep qt|czestoskurcz|bradykardi|amiodaron|sotalol|propafenon|flekainid|
   meksyletyn|adenozyn|chinidyn|dronedaron|werapamil|diltiazem|kardiowersj|chronotropow|dromotropow"""),
 ("LEKI HIPOLIPEMIZUJĄCE", r"""reduktaz(y|a) hmg|hmg.?coa|statyn|atorwastatyn|simwastatyn|symwastatyn|rozuwastatyn|
   prawastatyn|fluwastatyn|fibrat|fenofibrat|gemfibrozyl|ezetymib|kolestyramin|kolesewelam|
   ewolokumab|alirokumab|inklisiran|pcsk|niacyn|lipidogram|profil lipidow|cholesterol|
   hipolipemizuj|hiperlipidemi|hipercholesterol|rabdomioliz|miopati.*statyn"""),
 ("LEKI MOCZOPĘDNE", r"""moczopedn|diuretyk|furosemid|torasemid|kwas etakrynowy|bumetanid|
   mannitol|acetazolamid|dorzolamid|amilorid|triamteren|tolwaptan|akwaretyk|
   petlow|anhydraz weglanow|dehydrataz weglanow|natriuretyczn|hipokaliemi|hiperkaliemi|
   hiponatremi|hipernatremi|hiperurykemi|hiperurikemi|zasadowic metaboliczn|obrzek pluc"""),
 ("UKŁAD WSPÓŁCZULNY", r"""wspolczuln|adrenergiczn|sympatykomimetyk|katecholamin|
   amfetamin|metylfenidat|adrenalin|noradrenalin|efedryn|pseudoefedryn|fenylefryn|metoksamin|izoprenalin|
   ksylometazolin|nafazolin|tetryzolin|oksymetazolin|klemastyn(?!.*histamin)|
   fentolamin|fenoksybenzamin|
   beta.?adrenolit|beta.?bloker|metoprolol|bisoprolol|bizoprolol|atenolol|propranolol|karwedilol|
   tymolol|timolol|betaksolol|nebiwolol|salbutamol|formoterol|salmeterol|fenoterol|indakaterol|
   receptor(ow|y|a)? (alfa|beta)|alfa.?1|alfa.?2|beta.?2|mirabegron|hipotoni ortostatyczn"""),
 ("UKŁAD PRZYWSPÓŁCZULNY", r"""przywspolczuln|cholinergiczn|cholinolit|muskarynow|
   atropin|skopolamin|hioscyn|pilokarpin|karbachol|neostygmin|pirydostygmin|
   acetylocholinesteraz|acetylocholinoesteraz|edrofonium|
   oksybutynin|tolterodyn|solifenacyn|darifenacyn|ipratropium|tiotropium|glikopironium|
   bielun|lulek|nuzliwosc|myasthenia|miastenia|d.tubokuraryn|zwiotczaj|atrakurium|
   sukcynylocholin|rokuronium|fosforoorganiczn|pestycyd|zrenic|tropikamid|jaskr|
   nietrzymaniu moczu|zatrzymaniu moczu"""),
 ("UKŁAD ODDECHOWY", r"""oskrzel|astm|pochp|kaszel|kaszlow|wykrztus|mukolit|sekretolit|
   teofilin|aminofilin|roflumilast|montelukast|zafirlukast|zileuton|omalizumab|
   ambroksol|bromheksyn|acetylocystein|dornaz|gwajafenezyn|gwajakolosulfonian|
   dekstrometorfan|butamirat|pentoksyweryn|okseladyn|kodein(?!.*ból)|
   beklometazon|budezonid|cyklezonid|flutykazon|mometazon|
   niezyt nosa|kataru|zatkaniem nosa|wziewn|nebulizacj|duszno"""),
 ("UKŁAD POKARMOWY", r"""pokarmow|zoladk|jelit|trzustk|watrob(?!.*niewydolnosci watroby)|
   wrzodow|refluks|zaparc|biegunk|wymiot|nudnosc|przeczyszczaj|zapieraj|
   omeprazol|pantoprazol|lansoprazol|esomeprazol|inhibitor(y|ow)? pompy protonow|
   ranitydyn|famotydyn|cymetydyn|mizoprostol|sukralfat|bizmut|
   metoklopramid|itopryd|domperidon|cyzapryd|cisaprid|
   ondansetron|palonosetron|granisetron|tropisetron|aprepitant|dimenhydrynat|tietylperazyn|
   loperamid|difenoksylat|racekadotryl|oktreotyd|senes|makrogol|laktuloz|
   parafina|olej rycynowy|siarczan sodu|fosforan sodu|siarczan magnezu|
   mesalazyn|pankreatyn|ursodeoksychol|metylonaltrekson|
   zollinger|choroba lesniowskiego|colitis|vip.?oma"""),
 ("CUKRZYCA", r"""cukrzyc|glikemi|hipoglikemi|hiperglikemi|insulin|
   metformin|gliklazyd|glibenklamid|glimepiryd|glikwidon|glipizyd|sulfonylomoczn|
   repaglinid|nateglinid|akarboz|miglitol|pioglitazon|rozyglitazon|
   sitagliptyn|saksagliptyn|linagliptyn|wildagliptyn|dpp.?4|
   liraglutyd|semaglutyd|eksenatyd|dulaglutyd|glp.?1|inkretyn|
   dapagliflozyn|empagliflozyn|kanagliflozyn|flozyn|sglt|
   pramlintyd|glukagon|hba1c|kwasica ketonow|kwasica mleczanow"""),
 ("TARCZYCA", r"""tarczyc|tyroksyn|lewotyroksyn|liotyronin|tiamazol|metimazol|karbimazol|
   propylotiouracyl|jodek potasu|jodow(e|ych) srodk|niedoczynnosc tarczyc|nadczynnosc tarczyc|
   hashimoto|konwersj t4|t3 i t4|przelom tarczycow"""),
 ("GLIKOKORTYKOSTEROIDY", r"""glikokortykosteroid|kortykosteroid|gks |prednizon|prednizolon|
   metyloprednizolon|metylprednizolon|deksametazon|betametazon|hydrokortyzon|triamcynolon|
   kortyzol|nadnercz|addison|cushing"""),
 ("HORMONY PŁCIOWE", r"""steryd(y|ow)? anaboliczn|steroid(y|ow)? anaboliczn|maskulinizacj|testosteron|androgen|estrogen(?!.*antykonc)|gestagen(?!.*antykonc)|
   finasteryd|dutasteryd|klomifen|hormon(y|ow) plciow|raloksyfen"""),
 ("ANTYKONCEPCJA", r"""antykoncep|etynyloestradiol|lewonorgestrel|drospirenon|dezogestrel|
   gestagen|dwuskladnikow(e|ych) tabletk|owulacj|kapacytacj"""),
 ("OSTEOPOROZA", r"""osteoporoz|bisfosfonian|alendronian|ryzedronian|zoledronian|ibandronian|
   denosumab|teryparatyd|ibandronow|zolendronian|zoledronian|rank|raloksyfen|kalcytonin|witamina d|wapni.*kosc|martwica kosci"""),
 ("NIEOPIOIDOWE LEKI PRZECIWBÓLOWE", r"""nlpz|niesteroidow|cyklooksygenaz|cox|
   paracetamol|metamizol|ibuprofen|ketoprofen|naproksen|diklofenak|indometacyn|piroksykam|
   meloksykam|nimesulid|celekoksyb|koksyb|kwas acetylosalicylow|salicyl|
   przeciwgoraczkow|przeciwzapaln.*bol|nefopam"""),
 ("NARKOTYCZNE LEKI PRZECIWBÓLOWE", r"""opioid|morfin|kodein|oksykodon|fentanyl|petydyn|
   metadon|buprenorfin|pentazocyn|tramadol|nalokson|naltrekson|tapentadol|
   drabin(y|ie) analgetyczn|bol nowotworow|bol przebijajac|mioz.*opioid"""),
 ("ZNIECZULENIA MIEJSCOWE", r"""miejscowo znieczul|znieczulaj|lidokain|bupiwakain|prokain|
   artykain|mepiwakain|prylokain|benzokain|ropiwakain|kapsaicyn"""),
 ("ZNIECZULENIA OGÓLNE", r"""znieczuleni(a|e) ogoln|anestezj|indukcj znieczuleni|
   propofol|tiopental|etomidat|ketamin|halotan|sewofluran|izofluran|desfluran|
   podtlenek azotu|neuroleptanalgezj|droperidol|hipertermia zlosliw"""),
 ("PADACZKA", r"""padaczk|drgawk|przeciwdrgawkow|napad(y|ow)? (uogolnion|toniczno)|stan padaczkow|
   fenytoin|karbamazepin|okskarbazepin|kwas walproinow|walproinian|lamotrygin|lewetyracetam|
   lewetiracetam|topiramat|gabapentyn|pregabalin|tiagabin|wigabatryn|etosuksymid|fenobarbital|
   zonisamid|prog drgawkow"""),
 ("CHOROBA PARKINSONA", r"""parkinson|l.?dopa|lewodopa|karbidopa|benserazyd|entakapon|tolkapon|
   selegilin|rasagilin|pramipeksol|ropinirol|pergolid|kabergolin|bromokryptyn|
   amantadyn|procyklidyn|triheksyfenidyl|dyskinez"""),
 ("OTĘPIENIE", r"""piracetam|winpocetyn|nicergolin|funkcje kognitywn|alzheimer|otepien|donepezil|riwastygmin|rywastygmin|galantamin|memantyn"""),
 ("LEKI PRZECIWDEPRESYJNE", r"""przeciwdepresyj|depresj|ssri|snri|tlpd|trojpierscieniow|
   fluoksetyn|citalopram|escitalopram|sertralin|paroksetyn|fluwoksamin|
   wenlafaksyn|duloksetyn|mirtazapin|mianseryn|trazodon|bupropion|
   amitryptylin|imipramin|dezypramin|klomipramin|doksepin|
   moklobemid|inhibitor(y|ow) mao|tyramin|
   lit(u|em)?\b|weglan litu|stabilizuj(ac|ac)ych nastroj|afektywn|manii|maniakaln|
   zespol(u|em)? serotoninow|nikotyn"""),
 ("LEKI PRZECIWPSYCHOTYCZNE", r"""przeciwpsychotyczn|neuroleptyk|psychoz|schizofreni|
   haloperydol|chloropromazyn|perazyn|sulpiryd|promazyn|flufenazyn|
   klozapin|olanzapin|kwetiapin|rysperydon|risperidon|arypiprazol|zyprazydon|amisulpryd|
   pozapiramidow|akatyzj|dystoni|zlosliwy zespol neuroleptyczn|dyskinezy pozn|
   hiperprolaktynemi|dopaminergiczn.*blokad|blokad.*dopaminergiczn"""),
 ("LEKI PRZECIWLĘKOWE", r"""przeciwleko|lek(u|iem)? uogolnion|anksjolit|benzodiazepin|
   diazepam|lorazepam|alprazolam|oksazepam|klonazepam|midazolam|klorazepat|
   buspiron|hydroksyzyn|flumazenil|zespol abstynencyjn|klometiazol|
   terapi(a|i) leku|w leczeniu leku"""),
 ("LEKI NASENNE", r"""nasenn|bezsennosc|zolpidem|zopiklon|zaleplon|ramelteon|melatonin|
   estazolam|nitrazepam|temazepam"""),
 ("LEKI PRZECIWHISTAMINOWE", r"""przeciwhistaminow|histaminow(y|ego)? typu 1|receptor(a|ow)? h1|
   loratadyn|desloratadyn|cetyryzyn|lewocetyryzyn|feksofenadyn|ebastyn|bilastyn|rupatadyn|
   akrywastyn|antazolin|difenhydramin|prometazyn|ketotifen|klemastyn|
   alergiczn(y|ego) niezyt|pokrzywk|betahistyn|zawrot(y|ow) glowy|choroba lokomocyjn|blednikow"""),
 ("BÓLE GŁOWY", r"""migren|bol(u|ow)? glowy|tryptan|sumatryptan|zolmitryptan|ryzatryptan|
   eletryptan|ergotamin|dihydroergotamin|klasterow"""),
 ("STANY NAGŁE", r"""wstrzas|resuscytacj|anafilaktyczn|zatruci|odtrutk|antidotum|
   fomepizol|deferoksamin|hydroksykobalamin|dimerkaprol|wegiel aktywowany|
   roztwor nacl|0,9% roztwor|krystaloid|hydroksyetyloskrobi|plynoterapi|krystaloid|koloid|dekstran|albumin|
   metanol|glikol etylenowy|cyjanek|tlenek wegla|przedawkowani"""),
 ("LEKI TOKSYCZNE W CIĄŻY", r"""ciazy|ciezarn|teratogen|talidomid|etretynat|izotretynoin|
   karmieni piersi|laktacj"""),
 ("FARMAKOKINETYKA", r"""farmakokinetyk|biodostepnosc|okres poltrwania|t1/2|polokres|klirens|
   objetosc dystrybucji|kompartment|efekt pierwszego przejscia|pierwszego przejscia|
   metabolizm|metabolit|prolek|lek(i|ow)? prekursorow|cytochrom|cyp|indukuje enzym|
   induktor|inhibitor(em)? enzym|glikoprotein(a|y) p|p.?gp|
   rzedowosc|biorownowazn|dostepnosc biologiczn|terapi(a|i) monitorowan|stan stacjonarn|kinetyk|samoindukcj|indeks terapeutyczn|wiazani z bialkami|albumin(a|ami)|
   dawk(a|i) nasycaj|wydalani|wchlanian|jonizacj|receptor(a|ow)? blonow|receptor jadrow|
   allosteryczn|agonist(a|y)? czesciow|antagonist(a|y)? kompetycyj|
   interakcj|metabolizer"""),
 ("LEKI ROŚLINNE I SUPLEMENTY", r"""roslinn|ziol|suplement|dziurawc|zenszen|milorzab|
   witamin|zelaz|cyjanokobalamin|kwas foliowy|niedokrwistosc|retinol|erytropoetyn|saponin|garbnik|pierwiosnk|lukrecj|rzewien|senes(u)?\b|foeniculum|kozlek|
   produkt(em)? leczniczy|biotechnologiczn|biopodobn|krwiopochodn"""),
 ("JASKRA", r"""jaskr|cieczy wodnist|srodgalkow"""),
 ("LEKI IMMUNOSUPRESYJNE", r"""immunosupresyj|immunosupresj|cyklosporyn|takrolimus|pimekrolimus|
   sirolimus|ewerolimus|mykofenolan|azatiopryn|leflunomid|metylfenidat(?!x)|kalcyneuryn|
   adalimumab|infliksymab|etanercept|tocilizumab|ustekinumab|natalizumab|bazyliksymab|
   rytuksymab|rituksymab|ofatumumab|paliwizumab|przeciwcial(a|o|em) monoklonaln|
   przeszczep|odrzucani przeszczepu|luszczyc|stwardnieni rozsian|
   reumatoidaln|kolchicyn|dziegciow"""),
 ("FARMAKOLOGIA OGÓLNA I EBM", r"""qaly|aotmit|agencja oceny technologii|prisma|
   komisj(a|i) bioetyczn|badani(a|e|u) kliniczn|eksperyment(u|em)? leczniczy|swiadomej zgody|
   charakterystyk(a|i) produktu leczniczego|chpl|wyrob medyczny|refundacj|
   nadzor(u)? nad bezpieczenstwem|psur|farmakowigilancj|urzed(u)? rejestracji|
   metaanaliz|badani(a|e) kohortow|badani(a|e) obserwacyjn|hierarchii badan|
   pierwotnym zrodlem wiedzy|wytyczne praktyki|placebo|
   doping|terapii genowej|lekiem terapii genowej|typarwowek|fomiwirsen|mipomersen|
   substancje uzalezniajac|mezolimbiczn|jadro polleza"""),
]

CATS = [
 ("Ciąża", r"""ciazy|ciezarn|teratogen|karmieni piersi|laktacj"""),
 ("Antidotum", r"""antidotum|odtrutk|zatruci|przedawkowani|neutralizacj|odwrocen(ie)? (efektu|dzialani)"""),
 ("Interakcje", r"""interakcj|jednoczesne(go)? stosowani|lacznie z|skojarzeni|
   indukuje enzym|induktor|inhibitor(em)? cytochrom|cyp|nie mozesz zastosowac|
   reakcj disulfiram|z alkoholem"""),
 ("Monitorowanie", r"""monitorow|kontrolowac stezeni|inr\b|aptt|oznaczani stezeni"""),
 ("Dawkowanie", r"""dawkowani|dawk(e|i|a) (dobow|nasycaj|podtrzymuj)|modyfikacj(a|i) dawk"""),
 ("Sposób podania", r"""sposob podani|drog(a|i|e) podani|podawan(y|a|e) (doustnie|dozylnie|
   podskornie|domiesniowo|wziewnie|miejscowo)|wziewnie|doustnie|dozylnie|podskornie|
   pozajelitowo|donosowo|przezskorn|dopoliczkow"""),
 ("Przeciwwskazania", r"""przeciwwskaz|nie nalezy stosowac|nie mozna stosowac|
   nalezy (zrezygnowac|rezygnowac)|nalezy unikac|nie wolno|jest zabroni"""),
 ("Działania niepożądane", r"""dzialani(e|a|em) niepozadan|niepozadan|
   moze (powodowac|wywolac|byc przyczyna|wystapic)|powoduje|ryzyko (wystapienia|
   uszkodzenia|miopati)|toksyczn|uszkodzeni|objaw(y|em) uboczn"""),
 ("Wskazania", r"""wskazani|zastosujesz|mozesz zastosowac|stosuje sie|stosowan(y|a|e) (jest|w)|
   w leczeniu|w terapii|w profilaktyce|lek(iem)? z wyboru|celowe jest|uzasadnion|
   zalecisz|nalezy stosowac|w przerywaniu|mozna wybrac|nalezy leczyc|mozna stosowac|
   mozna podac|mozna zastosowac|nalezy podac|nalezy zastosowac|w celu |w lagodzeniu|
   doraznie|skuteczn|lekiem pierwszego rzutu|pierwszego rzutu|drugiego rzutu|
   wskazan(y|a|e) (jest|w)|leczeniu (choroby|zakazen|bolu|nadcisnien)|
   u pacjent(a|ow|ki) z|w napadzie|w stanie|zastosowa"""),
 ("Mechanizm działania", r"""mechanizm|hamuje|blokuje|pobudza|agonist|antagonist|
   receptor|enzym|kanal(y|ow)? (sodow|potasow|wapniow)|inhibitor|wiaze sie z"""),
 ("Farmakokinetyka", r"""farmakokinetyk|biodostepnosc|poltrwania|klirens|dystrybucji|
   metabolizm|metabolit|prolek|wchlanian|wydalani|pierwszego przejscia|kinetyk|
   wiazani z bialkami"""),
 ("Struktura chemiczna", r"""pochodn(a|ej|ymi) (kwasu|zwiazku)|struktur(a|e) chemiczn|pierscien"""),
 ("Klasyfikacja", r"""nalezy do|do lek(ow)? .* nalezy|jest lekiem|zalicza sie|
   klasyfikacj|grup(y|a) lek|przeciwciał(o|em) monoklonaln|analog(i|iem)? """),
]

def _rx(p):
    return re.compile(re.sub(r'\n\s*', '', p), re.I)


SECTIONS = [(n, _rx(p)) for n, p in SECTIONS]
CATS = [(n, _rx(p)) for n, p in CATS]


# "Klasyfikacja" tylko przy wyraznym sygnale przynaleznosci do grupy
CAT_W = {"Ciąża": 3, "Antidotum": 3, "Monitorowanie": 3, "Dawkowanie": 3,
         "Struktura chemiczna": 3, "Przeciwwskazania": 3, "Interakcje": 2,
         "Sposób podania": 2, "Farmakokinetyka": 2, "Działania niepożądane": 2,
         "Wskazania": 2, "Klasyfikacja": 1, "Mechanizm działania": 1}


def classify(qtext, opts):
    stem = nrm(qtext)
    blob = nrm(qtext + " " + " ".join(opts))
    best, score = None, 0
    for name, rx in SECTIONS:
        n = len(set(m.group(0) for m in rx.finditer(blob)))
        if n > score:
            best, score = name, n
    cb, cs = None, 0
    for name, rx in CATS:
        w = CAT_W.get(name, 1)
        s = w * (3 * len(rx.findall(stem)) + len(rx.findall(blob)))
        if s > cs:
            cb, cs = name, s
    return best, cb or "Klasyfikacja", score
