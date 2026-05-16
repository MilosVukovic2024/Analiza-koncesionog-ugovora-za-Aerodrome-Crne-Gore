// Analiza Ugovora o koncesiji za aerodrome Podgorica i Tivat
// Nezavisni pregled u javnom interesu - perspektiva nezavisnog/parlamentarnog/civilnog tijela

const fs = require('fs');
const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  Header, Footer, AlignmentType, LevelFormat,
  BorderStyle, WidthType, ShadingType,
  HeadingLevel, PageBreak, PageNumber
} = require('docx');

const FONT = "Calibri";
const COLOR_RED = "C0392B";
const COLOR_YELLOW = "B7950B";
const COLOR_GREEN = "1E8449";
const COLOR_INK = "1B1B1F";
const COLOR_MUTED = "555555";
const COLOR_RULE = "9AA0A6";

const border = { style: BorderStyle.SINGLE, size: 4, color: COLOR_RULE };
const tBorders = { top: border, bottom: border, left: border, right: border };

const H1 = (t) => new Paragraph({
  heading: HeadingLevel.HEADING_1,
  spacing: { before: 360, after: 200 },
  children: [new TextRun({ text: t, font: FONT, bold: true, size: 36, color: COLOR_INK })],
});
const H2 = (t) => new Paragraph({
  heading: HeadingLevel.HEADING_2,
  spacing: { before: 300, after: 160 },
  children: [new TextRun({ text: t, font: FONT, bold: true, size: 28, color: COLOR_INK })],
});
const H3 = (t, color = COLOR_INK) => new Paragraph({
  heading: HeadingLevel.HEADING_3,
  spacing: { before: 240, after: 120 },
  children: [new TextRun({ text: t, font: FONT, bold: true, size: 24, color })],
});
const P = (t, opts = {}) => new Paragraph({
  spacing: { after: 120, line: 320 },
  alignment: opts.center ? AlignmentType.CENTER : AlignmentType.JUSTIFIED,
  children: [new TextRun({ text: t, font: FONT, size: 22, color: COLOR_INK, italics: opts.italic, bold: opts.bold })],
});
const Q = (t) => new Paragraph({
  spacing: { before: 80, after: 120, line: 300 },
  indent: { left: 360, right: 360 },
  alignment: AlignmentType.JUSTIFIED,
  border: { left: { style: BorderStyle.SINGLE, size: 12, color: COLOR_RULE, space: 8 } },
  children: [new TextRun({ text: t, font: FONT, size: 20, italics: true, color: COLOR_MUTED })],
});
const Bullet = (t) => new Paragraph({
  numbering: { reference: "bul", level: 0 },
  spacing: { after: 80, line: 300 },
  children: [new TextRun({ text: t, font: FONT, size: 22, color: COLOR_INK })],
});
const Num = (t) => new Paragraph({
  numbering: { reference: "num", level: 0 },
  spacing: { after: 80, line: 300 },
  children: [new TextRun({ text: t, font: FONT, size: 22, color: COLOR_INK })],
});
const Spacer = () => new Paragraph({ children: [new TextRun({ text: " ", font: FONT, size: 16 })] });
const PageBr = () => new Paragraph({ children: [new PageBreak()] });
const Rule = () => new Paragraph({
  border: { bottom: { style: BorderStyle.SINGLE, size: 8, color: COLOR_RULE, space: 6 } },
  spacing: { after: 120 },
  children: [new TextRun({ text: "", font: FONT, size: 4 })],
});

const FindingBox = ({ tag, tagColor, code, title, clause, current, redline, rationale, priority, impact }) => {
  const fill = tagColor === COLOR_RED ? "FBEAE5" : tagColor === COLOR_YELLOW ? "FCF3CF" : "E8F5E9";
  const rows = [
    new TableRow({
      children: [new TableCell({
        borders: tBorders,
        width: { size: 9026, type: WidthType.DXA },
        shading: { fill, type: ShadingType.CLEAR },
        margins: { top: 100, bottom: 100, left: 160, right: 160 },
        children: [
          new Paragraph({
            spacing: { after: 60 },
            children: [
              new TextRun({ text: `${tag} · ${code}  `, font: FONT, bold: true, size: 18, color: tagColor }),
              new TextRun({ text: title, font: FONT, bold: true, size: 24, color: COLOR_INK }),
            ]
          }),
          new Paragraph({
            spacing: { after: 0 },
            children: [
              new TextRun({ text: "Klauzula: ", font: FONT, bold: true, size: 18, color: COLOR_MUTED }),
              new TextRun({ text: clause, font: FONT, size: 18, color: COLOR_MUTED }),
              new TextRun({ text: "    Prioritet: ", font: FONT, bold: true, size: 18, color: COLOR_MUTED }),
              new TextRun({ text: priority, font: FONT, size: 18, color: COLOR_MUTED }),
            ]
          }),
        ]
      })]
    }),
    new TableRow({
      children: [new TableCell({
        borders: tBorders,
        width: { size: 9026, type: WidthType.DXA },
        margins: { top: 160, bottom: 100, left: 160, right: 160 },
        children: [
          new Paragraph({
            spacing: { after: 60 },
            children: [new TextRun({ text: "Aktuelni tekst (sažeto):", font: FONT, bold: true, size: 20, color: COLOR_INK })]
          }),
          ...current.map(t => new Paragraph({
            spacing: { after: 80, line: 300 },
            indent: { left: 240 },
            border: { left: { style: BorderStyle.SINGLE, size: 10, color: tagColor, space: 8 } },
            children: [new TextRun({ text: t, font: FONT, size: 20, italics: true, color: COLOR_MUTED })]
          })),
          new Paragraph({
            spacing: { before: 160, after: 60 },
            children: [new TextRun({ text: "Predložena redakcija (redline):", font: FONT, bold: true, size: 20, color: COLOR_INK })]
          }),
          ...redline.map(t => new Paragraph({
            spacing: { after: 80, line: 300 },
            indent: { left: 240 },
            border: { left: { style: BorderStyle.SINGLE, size: 10, color: COLOR_GREEN, space: 8 } },
            children: [new TextRun({ text: t, font: FONT, size: 20, color: COLOR_INK })]
          })),
          new Paragraph({
            spacing: { before: 160, after: 60 },
            children: [new TextRun({ text: "Obrazloženje:", font: FONT, bold: true, size: 20, color: COLOR_INK })]
          }),
          ...(Array.isArray(rationale) ? rationale : [rationale]).map(t => new Paragraph({
            spacing: { after: 80, line: 300 },
            alignment: AlignmentType.JUSTIFIED,
            children: [new TextRun({ text: t, font: FONT, size: 20, color: COLOR_INK })]
          })),
          ...(impact ? [
            new Paragraph({
              spacing: { before: 120, after: 60 },
              children: [new TextRun({ text: "Poslovni i fiskalni uticaj:", font: FONT, bold: true, size: 20, color: COLOR_INK })]
            }),
            new Paragraph({
              spacing: { after: 60, line: 300 },
              alignment: AlignmentType.JUSTIFIED,
              children: [new TextRun({ text: impact, font: FONT, size: 20, color: COLOR_INK })]
            })
          ] : []),
        ]
      })]
    }),
  ];
  return new Table({
    width: { size: 9026, type: WidthType.DXA },
    columnWidths: [9026],
    rows,
  });
};

const titlePage = [
  new Paragraph({
    spacing: { before: 1200, after: 240 },
    alignment: AlignmentType.CENTER,
    children: [new TextRun({ text: "NEZAVISNA ANALIZA", font: FONT, bold: true, size: 28, color: COLOR_MUTED })]
  }),
  new Paragraph({
    spacing: { after: 120 },
    alignment: AlignmentType.CENTER,
    children: [new TextRun({ text: "Ugovor o koncesiji za aerodrome", font: FONT, bold: true, size: 48, color: COLOR_INK })]
  }),
  new Paragraph({
    spacing: { after: 720 },
    alignment: AlignmentType.CENTER,
    children: [new TextRun({ text: "Podgorica i Tivat", font: FONT, bold: true, size: 48, color: COLOR_INK })]
  }),
  new Paragraph({
    spacing: { after: 120 },
    alignment: AlignmentType.CENTER,
    children: [new TextRun({ text: "Pregled klauzula sa preporukama za redline", font: FONT, size: 24, color: COLOR_MUTED, italics: true })]
  }),
  new Paragraph({
    spacing: { after: 1200 },
    alignment: AlignmentType.CENTER,
    children: [new TextRun({ text: "Perspektiva: nezavisno tijelo / poslanik / javni interes", font: FONT, size: 22, color: COLOR_MUTED })]
  }),
  new Paragraph({
    spacing: { after: 60 },
    alignment: AlignmentType.CENTER,
    children: [new TextRun({ text: "Pripremljeno: maj 2026.", font: FONT, size: 20, color: COLOR_MUTED })]
  }),
  new Paragraph({
    spacing: { after: 60 },
    alignment: AlignmentType.CENTER,
    children: [new TextRun({ text: "Predmet: Predlog za davanje koncesije za aerodrome Podgorica i Tivat (Vlada Crne Gore)", font: FONT, size: 20, color: COLOR_MUTED })]
  }),
  new Paragraph({
    spacing: { after: 60 },
    alignment: AlignmentType.CENTER,
    children: [new TextRun({ text: "i Nacrt Ugovora o koncesiji (66 članova + 23 priloga)", font: FONT, size: 20, color: COLOR_MUTED })]
  }),
  PageBr(),
];

const disclaimer = [
  new Paragraph({
    spacing: { before: 400, after: 120 },
    alignment: AlignmentType.CENTER,
    children: [new TextRun({ text: "NAPOMENA O PRIRODI DOKUMENTA", font: FONT, bold: true, size: 22, color: COLOR_MUTED })]
  }),
  P("Ova analiza nije pravni savjet u smislu Zakona o advokaturi niti dokument zvanične državne procjene. Sačinjena je kao instrument za nezavisno čitanje ugovora u javnom interesu - radi parlamentarne rasprave, novinarske obrade, civilnog nadzora i zaštite imovine Crne Gore. Prije bilo kakvog formalnog djelovanja (amandman, prigovor, parnica) potreban je nezavisan pravni pregled kvalifikovanog advokata.", { italic: true }),
  P("Sve klauzulske reference odnose se na Nacrt Ugovora o koncesiji za aerodrome Podgorica i Tivat dostavljen kao prilog Predloga Vlade Crne Gore. Brojevi članova i priloga odgovaraju verziji u priloženom PDF dokumentu. Citati su sažeti zbog dužine.", { italic: true }),
  PageBr(),
];

const toc = [
  H1("Sadržaj"),
  P("1. Izvršni rezime"),
  P("2. Kontekst posla i metodologija"),
  P("3. Proceduralna pitanja tendera"),
  P("4. Registar odstupanja - CRVENI signali (10)"),
  P("5. Registar odstupanja - ŽUTI signali (10)"),
  P("6. Registar odstupanja - ZELENI signali (6)"),
  P("7. Pregovaračka strategija i eskalacioni put"),
  P("8. Preporuke za Skupštinu Crne Gore"),
  P("9. Aneks: Klauzule koje zaslužuju dodatnu provjeru"),
  PageBr(),
];

const exec = [
  H1("1. Izvršni rezime"),
  P("Nacrt Ugovora o koncesiji za aerodrome Podgorica i Tivat dodjeljuje Incheon Airport Consortiumu (IIAC + KIND, oba u državnom vlasništvu Republike Koreje) ekskluzivno pravo upravljanja, modernizacije, proširenja, korišćenja, održavanja i finansiranja dva ključna međunarodna aerodroma Crne Gore na period od 30 godina."),
  P("Finansijska struktura posla: 100 miliona EUR avansne koncesione naknade, 35,21% koncesione naknade na bruto prihod tokom 30 godina, inicijalni investicioni program od 132.003 miliona EUR, ciljani IRR sopstvenog kapitala 12,27%. Procijenjena vrijednost zemljišta i objekata koji prelaze u upravljanje koncesionara: oko 257 miliona EUR (235M zemljište + 22M zgrade)."),
  H3("Šta zaslužuje pažnju prije potpisivanja"),
  P("Iz nezavisne perspektive, deset klauzula iz Nacrta predstavljaju materijalne crvene signale za interes Crne Gore. Centralne tačke su sljedeće (svaka se obrazlaže u poglavlju 4):"),
  Num("Neopozivo i bezuslovno odricanje od suverenog imuniteta u svim jurisdikcijama (čl. 51.1.3) - najagresivnija formulacija odricanja imuniteta u međunarodnoj PPP praksi."),
  Num("Faktički monopol na razvoj međunarodnih aerodroma u Crnoj Gori za narednih 30 godina (čl. 35.1.2(b)) - izgradnja bilo kog novog komercijalnog aerodroma osim u Beranama, Ulcinju i Nikšiću predstavlja default države."),
  Num("Engleska verzija Ugovora preovladava nad crnogorskom (čl. 58) - potencijalna kolizija sa članom 13 Ustava Crne Gore."),
  Num("Pravo Projektnog SPV da raskine Ugovor iz \"razloga javne politike\" (čl. 41.1(d)) uz punu naknadu vlasničkog kapitala - klauzula opasna zbog svoje neodređenosti i obima naknade."),
  Num("EU acquis rizik prebačen na državu (čl. 36.5) - troškovi usklađivanja sa propisima EU preko 2 miliona EUR snosi država kroz kompenzaciju kroz Koncesionu naknadu."),
  Num("Gross-up obaveza za poreze (Prilog 10, čl. 4) - solidarna i neograničena odgovornost države za naknadu poreza na Plaćanja prilikom raskida."),
  Num("Koncesiona naknada se plaća i tokom Slučaja više sile (čl. 34.2.8) - neuobičajeno za projektno finansiranje."),
  Num("Asimetrija slučajeva neizvršenja obaveza (čl. 38) - 14 kategorija defaulta Projektnog SPV vs. 2 uske kategorije za Ugovorni organ."),
  Num("Nedovoljne garancije (čl. 30) - 8M EUR Garancija za izvršenje + 5M EUR Garancija za dobro izvršenje za stratešku imovinu vrijednosti preko 1,4 milijarde EUR."),
  Num("Neriješeni imovinsko-pravni sporovi na lokaciji Tivat - sam Predlog Vlade priznaje rizik aktiviranja raskidnih klauzula i odštetnih zahtjeva."),
  Spacer(),
  H3("Preporučeni put naprijed"),
  P("Nezavisna analiza ne preporučuje odbacivanje koncesionog modela, ali preporučuje uslovno odobrenje Skupštine sa ugrađenim izmjenama, koje su detaljno opisane u poglavlju 7. Ključno je: (a) renegocirati 10 crvenih klauzula prije potpisivanja; (b) ugraditi parlamentarne kontrolne tačke svakih 5 godina; (c) javno objaviti Ugovor i sve anekse prije Datuma zaključenja; (d) razriješiti imovinsko-pravne sporove u Tivtu prije potpisivanja, ne nakon."),
  PageBr(),
];

const context = [
  H1("2. Kontekst posla i metodologija"),
  H2("2.1 Predmet koncesije"),
  P("Koncesija obuhvata dva ključna međunarodna aerodroma Crne Gore - Podgorica (251,57 ha katastarskih parcela 541/6 i 541/12 KO Golubovci) i Tivat (157,86 ha kopnenog dijela + 8,83 ha akvatorijuma, podijeljeno u tri faze realizacije). Predmet su radovi modernizacije, projektovanja, proširenja, upravljanja, korišćenja, održavanja i finansiranja aerodroma."),
  H2("2.2 Strane Ugovora"),
  P("Koncedent je Crna Gora, koju zastupa Vlada Crne Gore (\"Ugovorni organ\"). Koncesionar je konzorcijum Incheon Airport, sačinjen od:"),
  Bullet("Incheon International Airport Corporation (IIAC), Republika Koreja - operater glavnog korejskog aerodroma, država u vlasništvu."),
  Bullet("Korea Overseas Infrastructure & Urban Development Corporation (KIND), Republika Koreja - korejska razvojna agencija, država u vlasništvu."),
  P("Operativni nosilac koncesije biće Projektno SPV koje će biti osnovano u Crnoj Gori isključivo za potrebe projekta, sa kojim će Ugovor stupiti na snagu kroz Ugovor o pristupanju (Prilog 23)."),
  H2("2.3 Finansijska struktura ponude"),
  P("Sljedeći finansijski parametri proizilaze iz finansijske ponude Incheon-a koju je Tenderska komisija ocijenila kao prvorangiranu:"),
  Bullet("Inicijalni investicioni program: 132.003.537,90 EUR"),
  Bullet("Avansna koncesiona naknada: 100.000.000 EUR"),
  Bullet("Procenat koncesione naknade: 35,21% bruto prihoda, polugodišnje"),
  Bullet("Ciljani IRR sopstvenog kapitala iz baznog slučaja: 12,27%"),
  Bullet("Naknada za izradu projekta (IFC): 1.200.000 EUR, plativa na račun IBRD-a u JP Morgan AG, Frankfurt"),
  Bullet("Sredstvo obezbjeđenja za zatvaranje: 2.000.000 EUR (do Finansijskog zatvaranja)"),
  Bullet("Sredstvo obezbjeđenja za izvršenje: 8.000.000 EUR"),
  Bullet("Garancija za dobro izvršenje posla: 5.000.000 EUR (godišnja)"),
  H2("2.4 Trajanje"),
  P("Koncesioni period iznosi 30 godina od Datuma početka (čl. 37). Predviđena su produženja u slučaju Više sile (čl. 34.3.1) i, indirektno, kroz mehanizme kompenzacije i odlaganja u slučajevima Značajne nepovoljne mjere države i Kvalifikovane izmjene zakona."),
  H2("2.5 Metodologija analize"),
  P("Ovaj pregled je proveden kroz: (a) klauzulu-po-klauzulu čitanje cijelog Nacrta Ugovora (66 članova + 23 priloga, ukupno preko 200 stranica); (b) poređenje sa međunarodnim standardima projektnog finansiranja za aerodromske PPP (referentno: koncesije u Hrvatskoj - Zagreb, Bugarskoj - Sofija, Albaniji - Tirana, Sjevernoj Makedoniji - Skoplje); (c) provjeru saobraznosti sa Zakonom o koncesijama Crne Gore (\"Sl. list CG\", broj 8/09) i Uredbom o bližem načinu sprovođenja postupka javnog nadmetanja (\"Sl. list CG\", broj 67/09); (d) ocjenu uticaja sa stanovišta javnog interesa i fiskalne odgovornosti države."),
  H2("2.6 Sistem signalizacije"),
  P("Klauzule su klasifikovane u tri kategorije:"),
  Bullet("CRVENI signal - materijalno odstupanje od standarda, prijetnja javnom interesu, traži eskalaciju i izmjenu prije potpisivanja."),
  Bullet("ŽUTI signal - odstupanje koje je u međunarodnoj praksi prisutno ali se može pregovarati u korist države."),
  Bullet("ZELENI signal - prihvatljiv i/ili poželjan tekst sa stanovišta javnog interesa."),
  PageBr(),
];

const procedural = [
  H1("3. Proceduralna pitanja tendera"),
  P("Pored materijalnih klauzula Ugovora, analiza javnog interesa zahtijeva i pregled procedure pod kojom je odabran Prvorangirani ponuđač. Procedura je trajala od 10. novembra 2017. (Zaključak Vlade br. 07-453) do dana izrade Predloga, kroz dvostepeni postupak:"),
  Num("Prva faza: javni oglas za pretkvalifikaciju (Sl. list CG 57/19). Sedam prijava, četiri kvalifikovana ponuđača, jedan (GMR Airports) povukao se 6. marta 2020."),
  Num("Druga faza: poziv za dostavljanje ponuda 20. decembra 2024. Dvije ponude pristigle 9. maja 2025."),
  Num("Tenderska komisija od 13 članova izvršila vrjednovanje u tri faze - tehnička ponuda, plan razvoja avio dostupnosti, investicioni plan, pa finansijska ponuda."),
  Num("Konačni rezultati: Incheon 96,18 poena (FK2 procenat naknade 35,21%, IRR 12,27%); Corporation America 65,15 poena (FK2 17,28%, IRR 12,45%)."),
  H2("3.1 Spor sa drugorangiranim ponuđačem"),
  P("Corporation America Airports S.A je 1. avgusta 2025. izjavila prigovor protiv Rang liste. Komisija za koncesije je 1. septembra 2025. donijela Rješenje br. 01/97 kojim se Rang lista vratila na ponovno odlučivanje. Slijedio je upravni spor pred Upravnim sudom i drugi prigovor (27. oktobra 2025), te konačno odluka kojom je drugi prigovor odbijen."),
  P("Sa stanovišta javnog interesa, postojanje spora oko bodovanja u dvostepenom postupku predstavlja procesni rizik: ako se u kasnijoj fazi (uključujući međunarodnu investicionu arbitražu) utvrdi povreda postupka, Crna Gora se može izložiti istovremeno gubitku posla i obavezi naknade štete jednoj ili obje strane. Preporuka: pribaviti pravno mišljenje o pravosnažnosti svih odluka u tenderskom postupku prije potpisivanja."),
  H2("3.2 Mišljenje Ministarstva finansija"),
  P("Predlog Vlade citira pozitivno mišljenje Ministarstva finansija na Predlog Ugovora o koncesiji, uz isticanje pozitivnog uticaja zbog značajnog obima novih investicija. Međutim, Predlog ne sadrži cjelovit tekst mišljenja niti rezerve koje je Ministarstvo finansija eventualno imalo, što ograničava nezavisni pregled."),
  H2("3.3 Mišljenje Zaštitnika imovinsko-pravnih interesa"),
  P("Predlog citira Mišljenje Zaštitnika imovinsko-pravnih interesa br. M.br. 37/25 od 3. aprila 2026. godine, kao i prethodno Mišljenje povodom Zahtjeva za davanje mišljenja od 15. aprila 2025. godine. Predlog navodi da su izmjene Ugovora rađene u skladu sa Mišljenjem, ali bez objavljenog teksta Mišljenja Zaštitnika nije moguće provjeriti u kojoj su mjeri zaštitne odredbe ugrađene u finalni Nacrt."),
  H2("3.4 Imovinsko-pravni sporovi na lokaciji Tivat"),
  P("Predlog Vlade eksplicitno upozorava (citat):"),
  Q("\"...na lokaciji aerodroma Tivat, naročito na priobalnom području u Tivtu, i dalje postoji značajan broj neriješenih imovinsko-pravnih sporova, koji potencijalno mogu da ugroze dinamiku izvršenja ugovorenih obaveza, odnosno posljedično i da izazovu eventualno aktiviranje raskidnih klauzula i odštetnih zahtjeva od strane Koncesionara ukoliko Koncedent ne izvrši u ugovorenim rokovima svoje ugovorene obaveze.\""),
  P("Ovo je najvažniji procesni signal u cijelom Predlogu. Ako Vlada ulazi u potpisivanje Ugovora sa znanjem da ne kontroliše imovinsko-pravni status značajnog dijela koncesionog područja u Tivtu, a Ugovor predviđa kompenzaciju Koncesionaru u slučaju da država ne riješi te sporove u rokovima, država se izlaže izvjesnoj obavezi plaćanja naknade po Prilogu 10. Preporuka: imovinsko-pravni sporovi u Tivtu moraju biti razriješeni ili eksplicitno isključeni iz koncesionog područja prije potpisivanja, ne nakon."),
  H2("3.5 Procjena vrijednosti imovine"),
  P("Predlog navodi procijenjenu vrijednost zemljišta na 235.111.490 EUR i građevinskih objekata na 22.220.249 EUR. Nije jasno na koju je datum izvršena procjena niti od strane koga - što je u skladu sa članom 9 stav 4 Zakona o koncesijama trebalo da uradi organ uprave nadležan za poslove imovine. Avansna koncesiona naknada od 100 miliona EUR + investicioni program od 132 miliona EUR + 35,21% godišnjeg prihoda tokom 30 godina, čine ekonomsku ravnotežu posla koja je u principu pozitivna za državu, pod uslovom da ne dođe do plaćanja po raskidu."),
  PageBr(),
];

// CRVENI SIGNALI
const redFindings = [
  FindingBox({
    tag: "CRVENI", tagColor: COLOR_RED, code: "C1",
    title: "Neopozivo odricanje od suverenog imuniteta u svim jurisdikcijama",
    clause: "čl. 51.1.3 i 51.1.4",
    priority: "Tier 1 - obavezno otkloniti",
    current: [
      "Ugovorni organ se ovim neopozivo i bezuslovno odriče, u najvećoj mjeri dozvoljenoj Primjenjivim pravom, svakog imuniteta od jurisdikcije (uključujući imunitet od tužbe, presude ili izvršenja) i svakog imuniteta od izvršenja ili sprovođenja bilo koje arbitražne ili sudske odluke koja proističe ili je u vezi sa ovim Ugovorom o koncesiji, u svakoj jurisdikciji u kojoj postupak može biti pokrenut.",
      "Odricanje se prema članu 51.1.4 primjenjuje na sve postupke priznanja, izvršenja i sprovođenja ICC arbitražne odluke."
    ],
    redline: [
      "Ugovorni organ se ovim opozivo i uslovno odriče imuniteta od jurisdikcije i imuniteta od izvršenja u vezi sa ovim Ugovorom o koncesiji, isključivo pred Međunarodnom arbitražom u skladu sa članom 50.3 i u jurisdikcijama u kojima je Ugovorni organ izričito pristao na izvršenje, a isključivo nad imovinom Crne Gore koja se koristi u komercijalne svrhe.",
      "Ovo odricanje izričito ne obuhvata: (a) imovinu namijenjenu obavljanju javnih funkcija (diplomatska, vojna, kulturna, monetarna rezerva centralne banke); (b) imovinu u skladu sa članom 23 Bečke konvencije o diplomatskim odnosima; (c) imovinu kategorisanu kao javno dobro u smislu Zakona o državnoj imovini."
    ],
    rationale: [
      "Aktuelna formulacija predstavlja najagresivniji oblik odricanja od suverenog imuniteta u međunarodnoj PPP praksi. Standard koji koriste države članice OECD-a (uključujući Njemačku, Francusku, Italiju) je tzv. 'commercial property carve-out' - odricanje samo nad imovinom u komercijalnoj upotrebi.",
      "Posljedice aktuelne formulacije: u slučaju arbitražnog gubitka, Koncesionar može zaplijeniti račune ambasada, vojnu opremu, kulturne predmete u inostranstvu, devizne rezerve Centralne banke - što se već dešavalo drugim državama (presedan: tužba NML Capital protiv Argentine).",
      "Iako je odricanje suverenog imuniteta uobičajeno u međunarodnom projektnom finansiranju, ono se ograničava na komercijalnu imovinu i na konkretne sudove. Neopozivost klauzule onemogućava buduće parlamente da modifikuju ovo odricanje."
    ],
    impact: "Potencijalno gubitak diplomatskih predstavništava, deviznih rezervi ili druge državne imovine u inostranstvu u slučaju nepovoljne arbitražne odluke. Procijenjeni rizik: 100M-500M EUR u najgorem scenariju."
  }),
  FindingBox({
    tag: "CRVENI", tagColor: COLOR_RED, code: "C2",
    title: "Faktički monopol aerodroma za 30 godina (zabrana novih aerodroma)",
    clause: "čl. 35.1.2(b) u vezi sa Prilogom 10 čl. 1.2",
    priority: "Tier 1 - obavezno otkloniti",
    current: [
      "Značajna nepovoljna mjera države označava i biće ograničena na sljedeće okolnosti: (b) izgradnja novog međunarodnog aerodroma sa komercijalnim uslugama u Crnoj Gori, osim Aerodroma Berane, Aerodroma Ulcinj i Aerodroma Nikšić.",
      "U slučaju Značajne nepovoljne mjere države, Ugovorni organ plaća Projektnom SPV: 100% Neizmirenog prioritetnog duga + Inicijalni Sopstveni kapital + Povrat Sopstvenog kapitala do 3 godine + Otpremnine + Troškovi podizvođača + Troškovi prenosa (Prilog 10, čl. 1.2)."
    ],
    redline: [
      "Brisati tačku (b) člana 35.1.2 u cjelosti, ili je zamijeniti formulacijom:",
      "(b) eksproprijacija postojećih komercijalnih aerodroma Podgorica i/ili Tivat ili dodjela paralelne komercijalne koncesije za iste aerodrome trećem licu tokom Koncesionog perioda. Izgradnja novih aerodromskih kapaciteta na drugim lokacijama u Crnoj Gori, uključujući regionalne, sportske, sezonske ili specijalizovane aerodrome, ne predstavlja Značajnu nepovoljnu mjeru države."
    ],
    rationale: [
      "Aktuelna klauzula zabranjuje Crnoj Gori da u narednih 30 godina razvija nove međunarodne aerodrome na bilo kojoj lokaciji izuzev tri navedene (Berane, Ulcinj, Nikšić, gdje aerodromi gotovo da ne funkcionišu). Ovo praktično znači gubitak suvereniteta nad budućim planiranjem vazdušnog saobraćaja.",
      "Hipotetički primjer rizika: ako se za 15 godina, zbog razvoja turizma ili regionalne saradnje, ukaže potreba za novim aerodromom (npr. u Cetinju ili Pljevljima), država bi morala platiti Koncesionaru cjelokupnu naknadu kao da je raskinula Ugovor - što po formuli iz Priloga 10 može iznositi 500M+ EUR.",
      "Standardni međunarodni standard u PPP koncesijama je 'non-competing project' klauzula koja se ograničava na isti aerodrom ili na aerodrome unutar definisanog radijusa (npr. 100 km). Klauzula koja se odnosi na cjelokupnu teritoriju države je neuobičajena."
    ],
    impact: "Gubitak suverene odluke o vazdušnom saobraćaju za 30 godina. Potencijalna obaveza naknade ako neki budući parlament odluči da razvija nove aerodromske kapacitete: 500M+ EUR."
  }),
  FindingBox({
    tag: "CRVENI", tagColor: COLOR_RED, code: "C3",
    title: "Engleska verzija preovladava nad crnogorskom",
    clause: "čl. 58",
    priority: "Tier 1 - obavezno otkloniti",
    current: [
      "Ovaj Ugovor o koncesiji je zaključen na crnogorskom i engleskom jeziku. U slučaju nesaglasnosti između engleske i crnogorske verzije, engleska verzija Ugovora o koncesiji će preovladati, ali Strane će u dobroj vjeri uzeti u obzir i crnogorsku verziju."
    ],
    redline: [
      "Ovaj Ugovor o koncesiji je zaključen na crnogorskom i engleskom jeziku, pri čemu obje verzije imaju jednaku pravnu snagu (autentičnost). U slučaju nesaglasnosti između crnogorske i engleske verzije, primjenjivaće se princip prijateljskog rješenja u skladu sa članom 50.1, a ako spor opstane, tumačenje će se vršiti u korist crnogorske verzije, kao verzije u službenom jeziku Crne Gore u skladu sa članom 13 Ustava Crne Gore. Engleska verzija će se koristiti kao pomoćni alat za tumačenje terminologije međunarodnog projektnog finansiranja."
    ],
    rationale: [
      "Član 13 Ustava Crne Gore: \"U Crnoj Gori je u službenoj upotrebi crnogorski jezik.\" Ugovor kojim država ustupa stratešku imovinu na 30 godina ne može imati pravnu snagu na jeziku koji nije službeni - što je principijelno pitanje suvereniteta.",
      "Praktične posljedice: u svakom sporu pred ICC arbitražom u Beču, Koncesionar će insistirati na engleskom tekstu. Crnogorski sudovi (u dijelu u kojem prinudne odredbe domaćeg prava preovladavaju - čl. 51.1.2) takođe će morati da poštuju engleski tekst.",
      "Reference: Hrvatska Koncesija aerodroma Zagreb (2013) - obje verzije autentične; Bugarska Koncesija aerodroma Plovdiv - bugarska verzija preovladava; Albanija Koncesija aerodroma Tirana - albanski preovladava."
    ],
    impact: "Sistemski rizik tumačenja u korist Koncesionara u svim sporovima tokom 30 godina."
  }),
  FindingBox({
    tag: "CRVENI", tagColor: COLOR_RED, code: "C4",
    title: "Raskid od strane Projektnog SPV iz \"razloga javne politike\"",
    clause: "čl. 41.1(d) u vezi sa Prilogom 10 čl. 1.2",
    priority: "Tier 1 - obavezno otkloniti ili precizno definisati",
    current: [
      "Projektno SPV će imati pravo da raskine ovaj Ugovor o koncesiji na osnovu: ... (d) razloga javne politike.",
      "U slučaju takvog raskida, primjenjuje se naknada iz Priloga 10, član 1.2, identična kao u slučaju Značajne nepovoljne mjere države."
    ],
    redline: [
      "Brisati tačku (d) člana 41.1 u cjelosti.",
      "Alternativno: (d) odluku Ustavnog suda Crne Gore kojom se utvrđuje da je ovaj Ugovor o koncesiji u suprotnosti sa Ustavom Crne Gore, uz prethodno iscrpljivanje svih pravnih sredstava i konsultacije sa Stranama u dobroj vjeri u roku od 180 dana. Tačka (d) ne obuhvata redovne promjene vlade, izmjene zakonodavstva (koje su obuhvaćene članom 36) ili druge promjene u državnoj politici."
    ],
    rationale: [
      "Termin \"razlozi javne politike\" nije definisan u Ugovoru. Koncesionar može jednostrano proglasiti da neka odluka Vlade, Skupštine, lokalne samouprave ili regulatornog tijela predstavlja \"razlog javne politike\" i raskinuti Ugovor, uz punu kompenzaciju.",
      "Posljedica: država nema kontrolu nad uslovima sopstvenog izlaska iz Ugovora. Naknada po Prilogu 10 čl. 1.2 uključuje: 100% senior duga + Inicijalni sopstveni kapital + Povrat sopstvenog kapitala do 3 godine + otpremnine + troškovi podizvođača + troškovi prenosa. Procjena: 300-700M EUR.",
      "Ovo je jedna od najopasnijih klauzula u Nacrtu jer je: (a) jednostrana - samo SPV može da je aktivira; (b) neograničena - nema definicije, ograničenja, ni filtriranja; (c) skupa - puna kompenzacija."
    ],
    impact: "Potencijalna obaveza naknade 300-700M EUR u slučaju neprecizirane političke promjene koju Koncesionar interpretira kao 'javnu politiku'."
  }),
  FindingBox({
    tag: "CRVENI", tagColor: COLOR_RED, code: "C5",
    title: "EU acquis rizik prebačen na državu",
    clause: "čl. 36.5",
    priority: "Tier 1 - obavezno otkloniti",
    current: [
      "Ako dođe do neke Kvalifikovane izmjene zakona zbog promjene u Važećem zakonodavstvu... koje je neophodno uskladiti sa propisima, direktivama, odlukama ili drugim obavezujućim dokumentima bilo kog nadležnog organa Evropske unije i ako je Projektno SPV dužno da se izloži dodatnom Kapitalnom trošku od najmanje EUR 2.000.000, Strane će razmatrati i dogovoriti pitanja na koja se upućuje u članu 36.1."
    ],
    redline: [
      "Brisati član 36.5 u cjelosti, ili ga zamijeniti formulacijom: 'U slučaju da Kvalifikovana izmjena zakona predstavlja transpoziciju EU acquis-a u domaće zakonodavstvo Crne Gore u skladu sa procesom EU integracija, Projektno SPV neće imati pravo na bilo koju naknadu osim ako se može dokazati da je izmjena: (a) bila razumno nepredvidiva u trenutku zaključenja Ugovora; (b) prouzrokovala dodatne troškove veće od 5.000.000 EUR; i (c) odnosila se isključivo na Aerodromske usluge a ne na aerodromske operatore generalno. U svim slučajevima, Projektno SPV snosi prvih 5.000.000 EUR troškova kao komercijalni rizik.'"
    ],
    rationale: [
      "Crna Gora je u procesu pridruživanja Evropskoj uniji i dužna je da transponuje cjelokupan EU acquis u domaće zakonodavstvo. Aerodromski sektor je posebno regulisan EU pravilima (Direktiva 2009/12/EZ o naknadama, Uredba 1008/2008 o pristupu tržištu, Uredba 2018/1139 o civilnom vazduhoplovstvu).",
      "Aktuelna klauzula znači da svaka transpozicija EU acquis-a koja košta SPV preko 2M EUR daje SPV-u pravo na naknadu od Vlade. U realnosti, ovo će se aktivirati za svaku ozbiljniju regulatornu izmjenu - jer 2M EUR je nizak prag za infrastrukturni projekat.",
      "Princip međunarodnog projektnog finansiranja je da Koncesionar preuzima rizik 'regulatorne promjene koja se mogla predvidjeti'. EU pridruživanje je javna politika Crne Gore već dvije decenije - svako razumno informisano lice moralo je očekivati transpoziciju EU pravila."
    ],
    impact: "Procijenjena akumulirana obaveza tokom 30 godina: 50-200M EUR samo iz EU acquis transpozicije."
  }),
  FindingBox({
    tag: "CRVENI", tagColor: COLOR_RED, code: "C6",
    title: "Solidarna i neograničena gross-up obaveza za poreze",
    clause: "Prilog 10, član 4",
    priority: "Tier 1 - obavezno otkloniti",
    current: [
      "Ukoliko bilo koja naknada po osnovu raskida, koju je Ugovorni organ dužan da plati Projektnom SPV u skladu sa članovima 1, 2 i 3 ovog Priloga 10, podliježe oporezivanju koje Projektno SPV duguje državnom organu u Crnoj Gori, tada će Ugovorni organ, uz ograničenja iz stava 3, snositi solidarnu i neograničenu odgovornost i biće dužan da isplati Projektnom SPV takav dodatni iznos koji će Projektno SPV dovesti u isti neto položaj nakon oporezivanja u kojem bi bilo da predmetna naknada po osnovu raskida nije bila predmet oporezivanja."
    ],
    redline: [
      "Ukoliko bilo koja naknada po osnovu raskida koju je Ugovorni organ dužan da plati Projektnom SPV podliježe oporezivanju u Crnoj Gori, Projektno SPV će snositi pun iznos tog poreza kao svoju redovnu poresku obavezu. Ugovorni organ nije obavezan da vrši gross-up plaćanja niti da preuzima poreske obaveze Projektnog SPV. Eventualne poreske olakšice, oslobađanja ili kreditni mehanizmi mogu biti predmet posebnog sporazuma između Strana, koji ne može povećati obavezu Ugovornog organa van iznosa propisanih ovim Prilogom."
    ],
    rationale: [
      "Gross-up klauzula u kontekstu plaćanja prilikom raskida znači sljedeće: ako Vlada plaća Koncesionaru npr. 400M EUR po raskidu, a porez na to iznosi 9% (40M EUR), Vlada mora platiti dodatnih 40M EUR + porez na te 40M (3,6M EUR) + porez na to itd. - dok porezna obaveza ne padne na nulu. Ukupno opterećenje može biti 10-15% više od nominalnog iznosa.",
      "\"Solidarna i neograničena odgovornost\" Ugovornog organa za poreske obaveze drugog lica je neuobičajena formulacija - u suštini znači da Vlada plaća poreze koje Koncesionar inače duguje Crnoj Gori.",
      "Princip neutralnosti poreza u međunarodnom projektnom finansiranju važi za poreze koji se naplaćuju od strane zemlje različite od zemlje domaćina (foreign withholding tax). Domaći porezi su redovan trošak poslovanja Koncesionara."
    ],
    impact: "Dodatno povećanje plaćanja prilikom raskida za 10-15%. Na 500M EUR primarne naknade, dodatnih 50-75M EUR."
  }),
  FindingBox({
    tag: "CRVENI", tagColor: COLOR_RED, code: "C7",
    title: "Koncesiona naknada se plaća i tokom Slučaja više sile",
    clause: "čl. 34.2.8",
    priority: "Tier 1 - obavezno otkloniti",
    current: [
      "Nastanak Slučaja više sile neće uticati na obavezu Projektnog SPV da plati Godišnju koncesionu naknadu."
    ],
    redline: [
      "Tokom trajanja Slučaja više sile, obaveza Projektnog SPV da plati Godišnju koncesionu naknadu se: (a) obustavlja u potpunosti ako Slučaj više sile sprječava više od 50% Aerodromskih usluga tokom najmanje 30 dana; (b) srazmjerno umanjuje za procenat smanjenja Bruto prihoda tokom Slučaja više sile, ako je to smanjenje manje od 50%. U svakom slučaju, kumulativno odlaganje plaćanja Koncesione naknade ne može preći 12 mjeseci."
    ],
    rationale: [
      "Aktuelna formulacija je interno nekonzistentna sa članom 34.2.3 koji oslobađa SPV od obaveza pružanja Aerodromskih usluga tokom Više sile. Ako SPV ne mora pružati usluge (i ne ostvaruje prihod), zašto bi plaćao naknadu na nulto-prihod?",
      "Sa stanovišta javnog interesa, ovo izgleda kao prednost za državu - ali u stvarnosti vodi do: (a) SPV traži produženje koncesionog perioda po članu 34.3 da nadoknadi izgubljene godine; (b) SPV traži raskid iz Slučaja više sile po članu 34.4 nakon 6 mjeseci uz punu kompenzaciju po Prilogu 10 čl. 3.",
      "Standardna međunarodna praksa: za vrijeme Slučaja više sile koji obustavlja operacije, koncesiona naknada se obustavlja ili srazmjerno umanjuje."
    ],
    impact: "U slučaju ozbiljnog Slučaja više sile (npr. pandemija COVID-19, oružani sukob), SPV će na 6+ mjeseci tražiti raskid i kompenzaciju umjesto obustave."
  }),
  FindingBox({
    tag: "CRVENI", tagColor: COLOR_RED, code: "C8",
    title: "Asimetrija slučajeva neizvršenja obaveza",
    clause: "čl. 38.1 vs. 38.2",
    priority: "Tier 1 - uravnotežiti",
    current: [
      "Slučajevi neizvršenja obaveza Projektnog SPV (čl. 38.1): 14 kategorija - odustajanje, insolventnost, povreda Koncesionih dokumenata, povreda Dokumenata o finansiranju, povreda izjava/garancija, neosiguranje, kašnjenje 180 dana, propust garancija, plaćanje korupcije, prevarne radnje, itd.",
      "Slučajevi neizvršenja obaveza Ugovornog organa (čl. 38.2): samo 2 kategorije - neplaćanje 90 dana ili materijalno kršenje koje 60 dana onemogućava Projektno SPV.",
      "Cure period za imovinske sporove (čl. 39.2.3): 8 mjeseci za Ugovorni organ, 60 dana za sve ostale slučajeve."
    ],
    redline: [
      "Proširiti član 38.2 (Slučajevi neizvršenja obaveza Ugovornog organa) tako da uključi: (c) neizdavanje Saglasnosti od strane Ugovornog organa kada je to njegova obaveza; (d) povredu izričitih obaveza saradnje po članovima 4 i 8; (e) povredu izjava i garancija u članu 48.3.",
      "Smanjiti cure period za imovinsko-pravne sporove sa 8 mjeseci na 90 dana, uz dodatnu mogućnost produženja od 90 dana uz pisani dogovor obje strane.",
      "Eliminisati ili strogo definisati neke od najobimnijih default kategorija u čl. 38.1, posebno (c)(ii) povredu Dokumenata o finansiranju, koja prebacuje rizik finansiranja na koncesioni odnos."
    ],
    rationale: [
      "Asimetrija nije neuobičajena u PPP - investitor preuzima više rizika operativno - ali aktuelni odnos 14:2 je značajno van standarda. Tipično odstojanje u međunarodnoj praksi je 8-10 vs. 4-5.",
      "Kategorija 38.1(c)(ii) je posebno problematična: bilo kakva bitna povreda Dokumenata o finansiranju (koji su izvan kontrole Ugovornog organa) može aktivirati raskid Ugovora. Ovo daje Finansijerima de facto kontrolu nad raskidom.",
      "Cure period od 8 mjeseci za imovinske sporove (39.2.3) je neuobičajeno dug - što jasno reflektuje znanje obje strane da imovinski sporovi u Tivtu nisu razriješeni."
    ],
    impact: "Asimetrični odnos znači da je vjerovatnoća raskida Ugovora 7-10x veća na strani SPV-a nego na strani države."
  }),
  FindingBox({
    tag: "CRVENI", tagColor: COLOR_RED, code: "C9",
    title: "Nedovoljne garancije za izvršenje (15M ukupno za 1,4B EUR asset)",
    clause: "čl. 30.1, 30.2, 30.3",
    priority: "Tier 1 - uvećati",
    current: [
      "Sredstvo obezbjeđenja za Zatvaranje: 2.000.000 EUR (do Finansijskog zatvaranja).",
      "Sredstvo obezbjeđenja za izvršenje: 8.000.000 EUR (do 30 dana nakon Potvrde o završetku radova, oko 6 godina od početka).",
      "Garancija za dobro izvršenje posla: 5.000.000 EUR (godišnja, tokom cijelog koncesionog perioda)."
    ],
    redline: [
      "Sredstvo obezbjeđenja za Zatvaranje: 5.000.000 EUR.",
      "Sredstvo obezbjeđenja za izvršenje: 25.000.000 EUR (oko 19% inicijalnog investicionog programa od 132M), kao bezuslovna bankarska garancija na prvi poziv.",
      "Garancija za dobro izvršenje posla: 15.000.000 EUR godišnje + 5% kumulativnog godišnjeg prihoda Aerodroma za prethodnu godinu, što god je veće.",
      "Dodati: Garancija za održavanje (Maintenance Bond) od 10.000.000 EUR za period između godine 25 i kraja koncesije, kao osiguranje da će se imovina vratiti u dobrom stanju."
    ],
    rationale: [
      "Ukupne garancije od 15M EUR (8+5+2) za stratešku imovinu vrijednosti procijenjenu na 257M EUR, sa inicijalnim investicijama od 132M EUR i potencijalnom obavezom Vlade pri raskidu od 500M+ EUR, su disproporcionalne.",
      "Međunarodni standard za aerodromske koncesije: Garancija za izvršenje 15-25% inicijalne CAPEX; Performance bond godišnji 3-5% prihoda.",
      "Posebna briga: Garancija za dobro izvršenje od 5M je manja od godišnjeg iznosa kumulativnih ugovornih kazni za neispunjenje usluga (Prilog 13)."
    ],
    impact: "Direktna izloženost Vlade riziku neizvršenja od oko 110-200M EUR (razlika između stvarne potrebe i garantovanog iznosa)."
  }),
  FindingBox({
    tag: "CRVENI", tagColor: COLOR_RED, code: "C10",
    title: "Neriješeni imovinsko-pravni sporovi u Tivtu su rizik koji se prenosi na državu",
    clause: "Sam Predlog Vlade + Ugovor čl. 39.2.3, 38.2",
    priority: "Tier 1 - obavezno otkloniti prije potpisivanja",
    current: [
      "Predlog Vlade citat: \"...na lokaciji aerodroma Tivat, naročito na priobalnom području u Tivtu, i dalje postoji značajan broj neriješenih imovinsko-pravnih sporova, koji potencijalno mogu da ugroze dinamiku izvršenja ugovorenih obaveza, odnosno posljedično i da izazovu eventualno aktiviranje raskidnih klauzula i odštetnih zahtjeva od strane Koncesionara...\"",
      "Ugovor čl. 39.2.3 propisuje cure period od 8 mjeseci za Ugovorni organ u slučaju imovinsko-pravnih sporova - što praktično priznaje očekivanje takvih sporova."
    ],
    redline: [
      "Prije potpisivanja Ugovora o koncesiji, Vlada je dužna da:",
      "(a) izvrši pravnu reviziju svih katastarskih parcela na lokaciji Aerodrom Tivat i identifikuje sve neriješene imovinsko-pravne sporove (broj predmeta, strane, predmet spora, vrijednost);",
      "(b) razriješi sporove ili obezbijedi pravne osnove (eksproprijacija, ovjereni notarski zapis, javnobeležnička isprava) koji omogućavaju mirno upravljanje koncesionim područjem;",
      "(c) ekspresno iz koncesionog područja izuzme parcele koje su predmet spora;",
      "(d) o (a), (b), (c) izvijesti Skupštinu Crne Gore i Komisiju za koncesije prije Datuma zaključenja Ugovora;",
      "U Ugovor unijeti odredbu: 'Datum zaključenja ne može nastupiti dok Vlada ne dostavi Projektnom SPV izvještaj o završetku imovinsko-pravne due diligence procedure za koncesiono područje, sa pozitivnom ocjenom Zaštitnika imovinsko-pravnih interesa.'"
    ],
    rationale: [
      "Vlada Crne Gore svjesno ulazi u Ugovor sa znanjem o postojanju imovinsko-pravnih sporova koji mogu da aktiviraju raskidne klauzule. To je izvjestan rizik, a ne hipotetički.",
      "Sa stanovišta člana 9 stav 4 Zakona o koncesijama, procjenu vrijednosti nepokretnosti koja je predmet koncesije vrši nadležni organ uprave. Ako nije izvršena pravna provjera vlasništva, procjena vrijednosti je sumnjiva.",
      "Ekonomska posljedica: ako se imovinski sporovi aktiviraju 5 godina nakon Datuma početka, kada je SPV već investirao 100M+ EUR, mehanizam Priloga 10 čl. 1.2 generiše kompenzaciju od 300-500M EUR."
    ],
    impact: "Procijenjeni rizik aktiviranja: vjerovatnoća 40-60% u prvih 10 godina; očekivana obaveza 300-500M EUR."
  }),
];

// ŽUTI SIGNALI
const yellowFindings = [
  FindingBox({
    tag: "ŽUTI", tagColor: COLOR_YELLOW, code: "Z1",
    title: "ICC arbitraža u Beču - alternative",
    clause: "čl. 50.3.5",
    priority: "Tier 2 - pregovarati",
    current: ["Sjedište arbitraže će biti u Beču. Arbitraža će se voditi pred tri arbitra po Pravilima ICC. Postupak se vodi i na engleskom i na crnogorskom jeziku."],
    redline: [
      "Razmotriti alternativne lokacije: Hag (Stalni arbitražni sud), Brisel (CEPANI), Ljubljana (Stalni arbitražni sud Slovenije) ili Pariz (ICC sjedište).",
      "Eksplicitno predvidjeti da arbitri moraju imati iskustvo u međunarodnom infrastrukturnom finansiranju i znanje balkanskog prava."
    ],
    rationale: [
      "Beč nije problematičan po sebi - VIAC je legitimna institucija. Međutim, ICC pravila u Beču znače: arbitri se imenuju u Parizu, jezik usmenog postupka je engleski (čl. 50.3.6 (a)), procijenjeni trošak 3-arbitražne ICC u Beču za spor 100M EUR: 800.000 - 1.500.000 EUR po strani.",
      "Alternative: Hag (UNCITRAL pravila, ne ICC) - jeftinije i neutralnije; CEPANI Brisel - ima više slavenskog jezičkog kapaciteta; Ljubljana - blizu Crne Gore, slovenački arbitri razumiju regionalno pravo."
    ],
    impact: "Direktni trošak arbitraže: 1-3M EUR po sporu. Indirektni: udaljenost foruma od domaćeg pravosuđa."
  }),
  FindingBox({
    tag: "ŽUTI", tagColor: COLOR_YELLOW, code: "Z2",
    title: "Step-in prava finansijera - kriterijumi za naslednika SPV-a",
    clause: "Prilog 20",
    priority: "Tier 2 - pregovarati",
    current: [
      "Period obustave od 60 dana za Finansijere nakon Obavještenja o neizvršenju.",
      "Finansijeri mogu imenovati dodatnog obveznika za step-in i postaviti zamjenu za SPV, uz saglasnost Ugovornog organa.",
      "U slučaju zamjene SPV, produženi period od 86 dana."
    ],
    redline: [
      "Definisati kriterijume za saglasnost Ugovornog organa na zamjenu SPV-a: (a) potencijalni naslednik mora imati iste tehničke i finansijske kvalifikacije koje su važile na pretkvalifikaciji 2019; (b) naslednik ne smije biti pravno lice iz države koja je predmet sankcija EU ili UN; (c) Ugovorni organ ima pravo veta na zamjenu iz razloga nacionalne bezbjednosti.",
      "Maksimalni step-in rok od 120 dana ne može se produžavati bez ponovljene saglasnosti Ugovornog organa.",
      "Odricanje od imuniteta u Direktnom ugovoru ograničiti istim kako je predloženo u C1."
    ],
    rationale: [
      "Step-in prava finansijera su standardna u međunarodnom projektnom finansiranju, ali kriterijumi za naslednika moraju biti precizni jer SPV može biti zamijenjen čak i strateški nepoželjnim subjektom.",
      "Posebno važno za aerodrom Podgorica kao bliz vojnim objektima i NATO obavezama - naslednik ne smije biti subjekt iz države koja nije saveznik."
    ],
    impact: "Strateški rizik nepoželjnog naslednika; direktni rizik produženja step-in roka."
  }),
  FindingBox({
    tag: "ŽUTI", tagColor: COLOR_YELLOW, code: "Z3",
    title: "Definicija \"Bruto prihoda\" za obračun koncesione naknade",
    clause: "čl. 29 (i definicije iz čl. 1.1)",
    priority: "Tier 2 - pregovarati i precizirati",
    current: [
      "Procenat koncesione naknade 35,21% obračunava se na osnovu Bruto prihoda iz nerevidiranih polugodišnjih izvještaja Projektnog SPV.",
      "Definicija Bruto prihoda iz čl. 1.1 nije dostupna u priloženom dijelu Nacrta - mora se posebno provjeriti."
    ],
    redline: [
      "Bruto prihod treba da uključuje sve prihode SPV bez izuzetaka, uključujući:",
      "(a) Avijacijski prihod (aerodromske takse, naknade za korišćenje terminala, ground handling);",
      "(b) Neavijacijski prihod (zakup retail prostora, parking, mjenjačnice, hoteli, oglašavanje, duty free, taksi, transportne usluge);",
      "(c) Sve prihode od zakupa Aerodromskih lokacija trećim licima;",
      "(d) Sve prihode od zajedničkih ulaganja, partnerstava i poslovnih jedinica;",
      "(e) Prihod od osiguranja u dijelu koji nadoknađuje izgubljeni operativni prihod.",
      "Isključenja se moraju eksplicitno navesti i biti minimalna (npr. samo iznos PDV-a koji se prenosi državi)."
    ],
    rationale: [
      "Razlika između Incheon-ovih 35,21% i CAAP-ovih 17,28% je 17,93 procentnih poena - što na 30 godina i prosječni godišnji prihod od 50M EUR čini razliku od oko 270M EUR. Vrlo je važno da je osnovica obračuna (Bruto prihod) široko definisana.",
      "Tipične zamke u definicijama Bruto prihoda: (a) isključenje neavijacijskih prihoda; (b) prebijanje 'rabata' za aviokompanije; (c) odbici za 'troškove naplate'; (d) prebijanje sa kapitalnim ulaganjima.",
      "Bez kontrolisanog auditing prava Ugovornog organa, SPV može utvrđivati Bruto prihod jednostrano."
    ],
    impact: "Razlika u definiciji Bruto prihoda može značiti 10-50M EUR godišnje izgubljene naknade."
  }),
  FindingBox({
    tag: "ŽUTI", tagColor: COLOR_YELLOW, code: "Z4",
    title: "Plaćanja prilikom raskida - formula i odlaganje",
    clause: "Prilog 10, čl. 1.1, 1.2, 2, 3",
    priority: "Tier 2 - pregovarati",
    current: [
      "Raskid zbog Slučaja neizvršenja obaveza SPV (čl. 1.1): 100% Neizmirenog prioritetnog duga - Troškovi novog Tendera - Prihodi iz osiguranja. Plaća se 12 mjeseci nakon Datuma raskida.",
      "Raskid zbog Značajne nepovoljne mjere države ILI iz razloga javne politike (čl. 1.2): 100% senior duga + Inicijalni sopstveni kapital + Povrat sopstvenog kapitala do 3 godine + Otpremnine + Troškovi podizvođača + Troškovi prenosa.",
      "Troškovi novog Tendera ograničeni na 1,2M EUR."
    ],
    redline: [
      "Limit Troškova novog Tendera: ukloniti gornju granicu od 1,2M EUR (ili je dignuti na 5M EUR), jer stvarni trošak ponovljenog tendera u prosjeku iznosi 3-7M EUR.",
      "Dodati klauzulu o mitigaciji: Vlada može vratiti SPV u prvobitno stanje ako u roku od 6 mjeseci ukloni uzrok Značajne nepovoljne mjere države.",
      "Period odlaganja plaćanja produžiti do 24 mjeseca sa zateznom kamatom ne većom od EURIBOR + 200 baznih poena.",
      "Povrat sopstvenog kapitala ograničiti na 1 godinu umjesto 3 godine, ili eliminisati u potpunosti."
    ],
    rationale: [
      "Aktuelna formula daje Koncesionaru 100% sigurnost povrata kapitala u svim slučajevima osim svoga vlastitog defaulta - što ne odgovara komercijalnom riziku koji bi trebalo da snosi privatni investitor.",
      "Povrat sopstvenog kapitala do 3 godine = SPV dobija ne samo svoj uloženi kapital nazad, nego i 3 godine očekivane dobiti (NPV oko 30-40% početnog kapitala). Ovo je suštinski 'profit guarantee' za 3 godine.",
      "Tipična međunarodna praksa: povrat sopstvenog kapitala se ne garantuje. Sopstveni kapital je risk capital."
    ],
    impact: "Razlika između sadašnje i predložene formule može iznositi 50-100M EUR po raskidu."
  }),
  FindingBox({
    tag: "ŽUTI", tagColor: COLOR_YELLOW, code: "Z5",
    title: "Ograničeno pravo intervenisanja Ugovornog organa",
    clause: "čl. 40",
    priority: "Tier 2 - pregovarati",
    current: [
      "Pravo intervenisanja samo za: (a) bitan poremećaj saobraćaja, (b) nemogućnost javnosti da bezbjedno koristi, (c) vanredne situacije, (d) prijetnja nacionalnoj bezbjednosti, ratno stanje.",
      "Ugovorni organ obavještava SPV 24 sata unaprijed.",
      "Ugovorni organ se ne smatra odgovornim za gubitak prihoda ili posljedični gubitak."
    ],
    redline: [
      "Proširiti listu okolnosti za intervenisanje:",
      "(e) nezavisni stručni nalaz o sistemskoj povredi minimalnih standarda usluga;",
      "(f) presuda nadležnog suda Crne Gore ili presuda Suda EU koja zahtijeva intervenciju;",
      "(g) javno zdravstvene mjere (pandemija, biološka prijetnja);",
      "(h) zaštita životne sredine u skladu sa članom 20."
    ],
    rationale: [
      "Pravo intervenisanja u javnom interesu mora obuhvatiti više situacija nego što je u Nacrtu - posebno javno zdravstvene mjere (pandemija) i zaštita životne sredine."
    ],
    impact: "Bez proširenja, država nema poluga da intervene u određenim situacijama javnog interesa bez aktiviranja raskidnih klauzula."
  }),
  FindingBox({
    tag: "ŽUTI", tagColor: COLOR_YELLOW, code: "Z6",
    title: "Kapa za ugovorne kazne (penalty cap)",
    clause: "čl. 24.2 i Prilog 13",
    priority: "Tier 2 - pregovarati",
    current: [
      "Obaveza Projektnog SPV za ugovorne kazne za svaku Godinu koncesije ograničena je na Maksimalni gornji prag (Prilog 13).",
      "Ako maksimalna ugovorna kazna se dostigne dvije uzastopne godine, to predstavlja Slučaj neizvršenja (čl. 38.1(h))."
    ],
    redline: [
      "Provjeriti i u Ugovoru izričito navesti iznos Maksimalnog gornjeg praga iz Priloga 13. Ako je prag manji od 5% godišnjeg Bruto prihoda - pregovarati podizanje na minimum 10% godišnjeg Bruto prihoda.",
      "Dodati klauzulu: Ako se Maksimalni gornji prag dostigne u jednoj koncesionoj godini, Ugovorni organ ima pravo na pojačani nadzor, dodatna periodična izvještavanja i pravo zahtijevati od SPV poseban korektivni plan (uz cure period od 60 dana)."
    ],
    rationale: [
      "Kapiranje ugovornih kazni u PPP je standardno, ali nivo kapa mora biti dovoljno visok da kazne stvarno utiču na ponašanje SPV-a.",
      "Trigger za default tek nakon 2 uzastopne godine je preliberalan - omogućava SPV-u da 'preživi' s niskim učinkom 2 godine, što za putnike i privredu znači ozbiljnu štetu."
    ],
    impact: "Loš učinak SPV-a tokom 2 godine može značiti milione izgubljenih prihoda od turizma, bez automatske posljedice po SPV."
  }),
  FindingBox({
    tag: "ŽUTI", tagColor: COLOR_YELLOW, code: "Z7",
    title: "Definicija Slučaja više sile - obim i isključenja",
    clause: "čl. 34.1",
    priority: "Tier 2 - pregovarati",
    current: [
      "Slučaj više sile je široko definisan, uključujući: prirodne nepogode, požar, eksploziju, radioaktivnu kontaminaciju, epidemije, štrajkove nacionalnog nivoa, sudare vazduhoplova, politička dešavanja izvan Crne Gore koja traju 14 dana.",
      "Isključeno: nemar SPV-a, ekonomske poteškoće, predvidive okolnosti."
    ],
    redline: [
      "Suziti definiciju Slučaja više sile dodavanjem isključenja:",
      "(e) pandemije osim ako su proglasile WHO ili nadležni državni organ kao globalnu pandemiju i traju duže od 90 dana;",
      "(f) terorističke akte koji se odnose isključivo na Aerodrome (umjesto bilo gdje van Crne Gore).",
      "(g) sudare vazduhoplova - osim ako se odnose direktno na infrastrukturu Aerodroma i zatvaraju je duže od 60 dana.",
      "Dodati klauzulu da SPV mora dokazati uzročnu vezu između Slučaja više sile i nemogućnosti izvršenja, sa standardom dokaza 'preponderance of evidence'."
    ],
    rationale: [
      "Aktuelna definicija je široka i može se aktivirati u brojnim slučajevima - što daje SPV-u poluga za izbjegavanje obaveza i/ili traženje produženja koncesionog perioda.",
      "Posebno problematično: 'politička dešavanja izvan Crne Gore' znači da rat u Ukrajini, blokada gasovoda, embargo na Rusiju mogu predstavljati Slučaj više sile čak i ako Crna Gora ne učestvuje."
    ],
    impact: "Široka definicija povećava vjerovatnoću Slučaja više sile što vodi produženju koncesionog perioda i/ili kompenzaciji."
  }),
  FindingBox({
    tag: "ŽUTI", tagColor: COLOR_YELLOW, code: "Z8",
    title: "Naknada za izradu projekta IFC-u (1,2M EUR)",
    clause: "čl. 31",
    priority: "Tier 2 - razjasniti",
    current: [
      "U roku od 5 dana od Datuma zaključenja, Prvorangirani ponuđač ili Projektno SPV plaća IFC-u 1.200.000 EUR Naknade za izradu Projekta.",
      "Plaća se na račun JP Morgan AG, Frankfurt, korisnik IBRD."
    ],
    redline: [
      "Razjasniti pravnu osnovu plaćanja IFC-u - na osnovu kojeg ugovora između Vlade Crne Gore i IFC-a?",
      "Provjeriti da li je iznos od 1,2M EUR razuman u poređenju sa drugim IFC mandatima.",
      "Ako Naknada predstavlja kompenzaciju za IFC kao savjetnika u tenderskom postupku, treba provjeriti da li su uslovi tog mandata javno dostupni."
    ],
    rationale: [
      "IFC ima dvostruku ulogu u ovakvim projektima: savjetnik privatizacije za vladu i potencijalni finansijer za pobjednika tendera. Konflikt interesa je realan.",
      "Naknada se plaća direktno IFC-u (preko IBRD u JP Morgan AG) - što nije transparentno kroz crnogorske finansijske institucije."
    ],
    impact: "Iznos je relativno mali (1,2M EUR), ali pitanje transparentnosti i pravne osnove je principijelno."
  }),
  FindingBox({
    tag: "ŽUTI", tagColor: COLOR_YELLOW, code: "Z9",
    title: "Zaštita zaposlenih - 5 godina je dobro, šta nakon?",
    clause: "čl. 2.2.3 i čl. 23",
    priority: "Tier 2 - pregovarati",
    current: [
      "Plan za transfer zaposlenih (čl. 2.2.3) uključuje zabranu raskida ugovora o radu sa zaposlenima na neodređeno vrijeme u roku od 5 godina nakon transfera.",
      "Nakon 5 godina nema dodatnih obaveza."
    ],
    redline: [
      "Produžiti zaštitu od raskida ugovora o radu na neodređeno vrijeme sa 5 na 10 godina.",
      "Dodati klauzulu: Nakon perioda zaštite, eventualna kolektivna otpuštanja moraju biti predmet konsultacija sa Ugovornim organom najmanje 6 mjeseci unaprijed, i moraju ispunjavati uslove IFC Standarda učinka 2.",
      "Dodati obavezu: SPV mora godišnje izvještavati o nivou zaposlenosti, prosječnoj plati, kolektivnim ugovorima i mjerama bezbjednosti i zdravlja na radu."
    ],
    rationale: [
      "5-godišnja zaštita je IFC standard, ali za 30-godišnju koncesiju to je samo 16,7% perioda. Nakon 5 godina, SPV može slobodno restrukturirati radnu snagu.",
      "Sa stanovišta javnog interesa, zaposleni u strateškim infrastrukturnim sektorima zaslužuju veću zaštitu nego u običnim privatizacijama."
    ],
    impact: "Direktni socijalni rizik za oko 500 zaposlenih u Aerodromima Crne Gore d.o.o."
  }),
  FindingBox({
    tag: "ŽUTI", tagColor: COLOR_YELLOW, code: "Z10",
    title: "Transparentnost - obaveza javnog godišnjeg izvještavanja",
    clause: "čl. 57 i 19 (Nadzor)",
    priority: "Tier 2 - pregovarati",
    current: [
      "Ugovor o koncesiji se može objaviti na web sajtu Ugovornog organa (čl. 57.3).",
      "Nije eksplicitno predviđena obaveza javnog izvještavanja o ključnim pokazateljima koncesije."
    ],
    redline: [
      "Dodati novi član 57.4: 'Projektno SPV i Ugovorni organ će zajedno objavljivati godišnji javni izvještaj o realizaciji koncesije, koji obuhvata: (a) broj putnika i kretanja vazduhoplova; (b) bruto prihod po kategorijama; (c) plaćena koncesiona naknada; (d) izvršene investicije; (e) zaposlenost; (f) bezbjednosne incidente; (g) pritužbe putnika i njihovo rješavanje. Izvještaj se objavljuje na web sajtovima Ugovornog organa i SPV-a najkasnije 4 mjeseca nakon kraja Godine koncesije.'",
      "Skupština Crne Gore ima pravo na godišnju raspravu o realizaciji koncesije, uz prisustvo predstavnika SPV-a."
    ],
    rationale: [
      "Aktuelni Nacrt omogućava objavu Ugovora ali ne predviđa kontinuirano javno izvještavanje - što je standard u modernim PPP koncesijama.",
      "Bez kontinuirane transparentnosti, kontrola Skupštine i javnosti tokom 30 godina je nominalna."
    ],
    impact: "Nedostatak transparentnosti otežava civilni i parlamentarni nadzor tokom 30 godina."
  }),
];

// ZELENI SIGNALI
const greenFindings = [
  FindingBox({
    tag: "ZELENI", tagColor: COLOR_GREEN, code: "Y1",
    title: "Rezervisane aktivnosti za vojsku, NATO, policiju, carinu",
    clause: "čl. 17 i 17.6",
    priority: "Prihvatljivo",
    current: [
      "Projektno SPV je dužno da omogući Vojsci Crne Gore, Ugovornom organu i članicama NATO besplatno, neometano korišćenje pista, rulnih staza i aerodromskih instalacija u skladu sa mjerodavnim pravom i obavezama Crne Gore u NATO.",
      "Carina, imigracija, karantin, policija - sve o trošku SPV-a."
    ],
    redline: [
      "Bez izmjena. Ovo je jedna od najbolje formulisanih klauzula u Nacrtu.",
      "Poželjan dodatak: Eksplicitno navesti obavezu SPV-a da o trošku obezbijedi tehničku opremu za bezbjednost (skener, kontrola pasoša) u skladu sa standardima EU i NATO."
    ],
    rationale: [
      "Aerodromi su nacionalna bezbjednosna infrastruktura i klauzula čuva NATO i vojne prerogative.",
      "Crna Gora je članica NATO od juna 2017 - klauzula je u skladu sa članom 5 NATO ugovora."
    ]
  }),
  FindingBox({
    tag: "ZELENI", tagColor: COLOR_GREEN, code: "Y2",
    title: "Zabrana otpuštanja zaposlenih 5 godina (IFC standard 2)",
    clause: "čl. 2.2.3(b)",
    priority: "Prihvatljivo, može se ojačati (vidi Z9)",
    current: ["Zabrana raskida ugovora o radu sa zaposlenima koji rade na neodređeno vrijeme u roku od 5 godina nakon transfera."],
    redline: ["Bez izmjena u ovom dijelu (vidi Z9 za predlog produženja roka)."],
    rationale: [
      "Klauzula je u skladu sa IFC Performance Standard 2 (Labor and Working Conditions) i sa Zakonom o radu Crne Gore.",
      "Korisno: Plan kolektivnog otpuštanja (čl. 2.2.3(c)) zahtijeva analizu alternativa i konsultacije."
    ]
  }),
  FindingBox({
    tag: "ZELENI", tagColor: COLOR_GREEN, code: "Y3",
    title: "Nezavisni nadzorni organ",
    clause: "čl. 5 i Prilog 14",
    priority: "Prihvatljivo",
    current: [
      "Strane uspostavljaju Nezavisni nadzorni organ koji prati i izvještava o ispunjavanju obaveza po Ugovoru. Detalji u Prilogu 14."
    ],
    redline: [
      "Bez izmjena u principu. Provjera Priloga 14 potrebna - posebno: (a) sastav i nezavisnost organa; (b) finansiranje; (c) ovlašćenja (savjetodavna ili obavezujuća)."
    ],
    rationale: [
      "Nezavisno nadzorno tijelo je dobra praksa u dugoročnim PPP koncesijama.",
      "Mora se obezbijediti stvarna nezavisnost - finansiranje 50:50 od strane Ugovornog organa i SPV-a, sa rotacijom predsjedavajućeg."
    ]
  }),
  FindingBox({
    tag: "ZELENI", tagColor: COLOR_GREEN, code: "Y4",
    title: "Ugovor se objavljuje na web sajtu",
    clause: "čl. 57.3",
    priority: "Prihvatljivo",
    current: ["Ovaj Ugovor o koncesiji neće predstavljati Povjerljive informacije i Ugovorni organ ga može objaviti uključujući putem objavljivanja na njegovom web sajtu."],
    redline: ["Pojačati formulaciju: 'Ovaj Ugovor o koncesiji i svi prilozi su javni dokument i Ugovorni organ je dužan da ga objavi na svom web sajtu u roku od 30 dana od Datuma zaključenja, na crnogorskom i engleskom jeziku.'"],
    rationale: [
      "Aktuelni tekst kaže 'može' - što je permisivno. Predloženi tekst čini obavezujuće.",
      "Konkretni rok i jezik objave su standard u javnim dokumentima."
    ]
  }),
  FindingBox({
    tag: "ZELENI", tagColor: COLOR_GREEN, code: "Y5",
    title: "Ekološke i društvene obaveze prema IFC standardima",
    clause: "čl. 20 i Prilog 15",
    priority: "Prihvatljivo",
    current: [
      "Projektno SPV je dužno da poštuje IFC Performance Standards i Equator Principles tokom Koncesionog perioda.",
      "Konsultant za ekološka i društvena pitanja prati usklađenost."
    ],
    redline: ["Bez izmjena u principu. Provjera Priloga 15 potrebna - posebno standardi za bird strike, buku, kvalitet vazduha, otpadne vode."],
    rationale: [
      "IFC standardi (PS1-PS8) su globalno priznati standardi - dobra praksa.",
      "Crna Gora kao članica UN obavezna je na ovaj nivo."
    ]
  }),
  FindingBox({
    tag: "ZELENI", tagColor: COLOR_GREEN, code: "Y6",
    title: "Ograničenje SPV-a na isključivo aerodromsku djelatnost",
    clause: "čl. 48.1(a)",
    priority: "Prihvatljivo",
    current: ["Projektno SPV je propisno osnovano, zakonito posluje, kao privatno društvo... isključivo u cilju razvoja, upravljanja i rada Aerodroma."],
    redline: ["Bez izmjena. Single-purpose vehicle je standard za PPP."],
    rationale: [
      "SPV koje se bavi samo Aerodromima sprječava 'cross-subsidiziranje' iz drugih biznisa i čuva transparentnost prihoda i troškova."
    ]
  }),
];

const strategy = [
  H1("7. Pregovaračka strategija i eskalacioni put"),
  H2("7.1 Trostepena strategija pregovaranja"),
  P("Nezavisna analiza preporučuje sljedeći trostepeni pristup pregovaranja sa Vladom Crne Gore i, posredno, sa Konzorcijumom Incheon:"),
  H3("Tier 1 - Must-Haves (10 crvenih signala)"),
  P("Sve C1-C10 stavke iz poglavlja 4 su uslovi bez kojih se Ugovor ne bi smio potpisati. Posebno: C1 (suvereni imunitet), C2 (monopol), C3 (jezik), C4 (raskid iz razloga javne politike), C10 (Tivat imovinski sporovi). Ovih pet je realno crvena linija."),
  P("Eskalacioni put: ako Konzorcijum Incheon ne pristane na materijalne izmjene C1-C10:"),
  Num("Skupština Crne Gore donosi Odluku kojom uslovljava davanje koncesije izmjenama Nacrta Ugovora."),
  Num("Ministarstvo saobraćaja, kao nadležni organ, otvara dodatni krug pregovora sa Konzorcijumom."),
  Num("Ako se ne postigne dogovor u roku od 90 dana, postupak se vraća drugorangiranom ponuđaču (Corporation America) u skladu sa Uredbom o bližem načinu sprovođenja postupka."),
  Num("U krajnjem slučaju, postupak se ponavlja."),
  H3("Tier 2 - Should-Haves (10 žutih signala)"),
  P("Z1-Z10 stavke su poželjne izmjene koje treba pregovarati. Skupština može uslovno odobriti dodjelu koncesije sa zahtjevom da Vlada izvijesti o ishodu pregovora prije Datuma zaključenja."),
  P("Strategija: koristiti Tier 3 ustupke za sklapanje Tier 2 pobjeda. Konzorcijum Incheon će vjerovatno biti spreman za kompromis na Z3 (definicija Bruto prihoda) i Z4 (formula plaćanja prilikom raskida) jer su u pitanju brojevi, ne principi."),
  H3("Tier 3 - Nice-to-Haves"),
  P("Ostala formalna i tehnička pitanja koja se mogu prepustiti tokom finalne dorade."),
  H2("7.2 Kalendar pregovaranja"),
  Num("Sedmice 1-2: Skupština prima Predlog Vlade, formira radnu grupu za nezavisnu analizu, raspravlja o crvenim signalima."),
  Num("Sedmice 3-4: Komisija Skupštine drži javnu raspravu sa pozivom predstavnicima Vlade, Konzorcijuma Incheon, Aerodroma Crne Gore d.o.o., Zaštitnika imovinsko-pravnih interesa, civilnog sektora."),
  Num("Sedmice 5-8: Skupština donosi uslovnu Odluku o davanju koncesije sa anexom uslovljavanja."),
  Num("Sedmice 9-16: Vlada vodi dopunske pregovore sa Konzorcijumom Incheon o izmjenama."),
  Num("Sedmica 17: Vlada izvještava Skupštinu o rezultatima pregovora. Ako su sve materijalne izmjene postignute - Skupština daje finalnu saglasnost. Ako ne - aktivira se eskalacioni put."),
  H2("7.3 Eskalacioni protokol unutar pregovaračkog tima Vlade"),
  Num("Pregovaračka grupa (ministar saobraćaja + savjetnici + spoljni pravni savjetnik) - operativni nivo."),
  Num("Eskalacija 1: Premijer + ministar finansija + ministar saobraćaja - strateški nivo."),
  Num("Eskalacija 2: Vlada u punom sastavu - vladin nivo."),
  Num("Eskalacija 3: Skupština i predsjednik Crne Gore - državni nivo."),
  P("Crveni signali C1-C10 zahtijevaju eskalaciju na nivo 3 prije bilo kog kompromisa."),
  PageBr(),
];

const recommendations = [
  H1("8. Preporuke za Skupštinu Crne Gore"),
  H2("8.1 Šta tražiti od Vlade prije glasanja"),
  Num("Cjelovit tekst Mišljenja Ministarstva finansija o Predlogu Ugovora (ne samo rezime)."),
  Num("Cjelovit tekst Mišljenja Zaštitnika imovinsko-pravnih interesa br. M.br. 37/25."),
  Num("Detaljan izvještaj o stanju imovinsko-pravnih sporova na lokaciji Tivat sa procjenom vrijednosti."),
  Num("Detaljan izvještaj o procjeni vrijednosti imovine koja je predmet koncesije - sa metodologijom, datumom i potpisom procjenitelja."),
  Num("Cjelovit Aneks Definicije iz Ugovora (Bruto prihod, Avijacijski prihod, Neavijacijski prihod, IRR)."),
  Num("Prilog 8 (Osiguranje) - tačni iznosi."),
  Num("Prilog 13 (Kriterijumi učinka) - posebno Maksimalni gornji prag ugovornih kazni."),
  Num("Cjelovit Finansijski model SPV-a (po mogućnosti sa NDA za poslanike)."),
  Num("Cjelovit mandat IFC-a u tenderskom postupku."),
  H2("8.2 Šta dodati u Odluku Skupštine o davanju koncesije"),
  P("Odluka iz člana 9 stav 3 Zakona o koncesijama može uslovljavati davanje koncesije. Preporučeni dodaci:"),
  Num("Davanje koncesije se uslovljava ispunjavanjem najmanje sljedećih materijalnih izmjena Nacrta Ugovora: [lista 10 crvenih signala iz poglavlja 4]."),
  Num("Vlada Crne Gore izvještava Skupštinu o izvršenoj proceduri usklađivanja Ugovora i pribavljanju saglasnosti Konzorcijuma Incheon u roku od 60 dana od dana stupanja na snagu ove Odluke."),
  Num("Ako Konzorcijum Incheon u roku od 60 dana ne pristane na materijalne izmjene, postupak davanja koncesije se vraća na ponovno odlučivanje Komisiji za koncesije i Tenderskoj komisiji."),
  Num("Ugovor o koncesiji ne može biti potpisan prije nego što Vlada Crne Gore Skupštini dostavi izvještaj o razrješenju imovinsko-pravnih sporova na lokaciji Tivat ili o izuzimanju spornih parcela iz koncesionog područja."),
  Num("Skupština Crne Gore osniva Stalnu parlamentarnu komisiju za nadzor koncesije aerodroma."),
  Num("Godišnji izvještaj o realizaciji koncesije se javno objavljuje od strane Ugovornog organa i razmatra na javnoj sjednici Skupštine."),
  H2("8.3 Mehanizmi nadzora tokom 30 godina"),
  P("30-godišnja koncesija znači da će Ugovor preživjeti 7-8 saziva Skupštine. Mehanizmi koje treba ugraditi sada:"),
  Num("Stalna parlamentarna komisija sa pravom uvida u sve dokumente Aerodromskog foruma (Prilog 5) i Nezavisnog nadzornog organa (Prilog 14)."),
  Num("Periodični ekonomski stress-test svake 5. godine - nezavisna procjena ekonomske ravnoteže koncesije."),
  Num("Klauzula o reotvaranju (re-opener) - obavezna ponovljena procjena materijalnih uslova svake 10. godine."),
  Num("Javno dostupna baza podataka o ispunjavanju kriterijuma učinka iz Priloga 13 - putem web sajta Ministarstva saobraćaja."),
  H2("8.4 Zaštita od političke instrumentalizacije"),
  P("Sa stanovišta civilnog nadzora, koncesija aerodroma ne smije postati partijsko pitanje:"),
  Bullet("Skupština Crne Gore izglasava Odluku o davanju koncesije sa kvalifikovanom većinom (2/3) - kao za odluke od strateškog značaja - ako to dozvoljava poslovnik."),
  Bullet("U Stalnu parlamentarnu komisiju za nadzor uključiti predstavnike vlasti i opozicije srazmjerno."),
  Bullet("Eventualne buduće izmjene Ugovora (čl. 55) zahtijevaju istu kvalifikovanu većinu u Skupštini."),
  Bullet("Civilni sektor (MANS, Institut alternativa, CGO) ima pravo na konsultativnu ulogu u Stalnoj parlamentarnoj komisiji."),
  PageBr(),
];

const annex = [
  H1("9. Aneks: Klauzule koje zaslužuju dodatnu provjeru"),
  P("Sljedeće klauzule i prilozi nisu detaljno analizirani u glavnom dijelu (zbog dostupnosti) i zaslužuju nezavisnu provjeru kvalifikovanog pravnog savjetnika:"),
  H3("Prilozi koji se moraju dobaviti i analizirati"),
  Bullet("Prilog 1 (Prethodni uslovi) - obje strane imaju 12 mjeseci da ih ispune."),
  Bullet("Prilog 2 (Inicijalni prenos imovine i obaveza) - šta se tačno prenosi sa Aerodroma Crne Gore d.o.o. na SPV."),
  Bullet("Prilog 3 (Program radova) - definitivna lista investicionih obaveza po fazama."),
  Bullet("Prilog 8 (Osiguranje) - minimalni iznosi i tipovi osiguranja."),
  Bullet("Prilog 9 (Povratni prenos imovine) - precizne procedure, što se vraća."),
  Bullet("Prilog 13 (Kriterijumi učinka) - Maksimalni gornji prag ugovornih kazni."),
  Bullet("Prilog 16 (Sredstva za rezervisane aktivnosti) - prostor za državne organe."),
  Bullet("Prilog 19 (Zaposleni koji su predmet transfera) - precizna lista."),
  Bullet("Prilog 21 (Regulisane aerodromske takse) - cjenovnik."),
  Bullet("Finansijski model Konzorcijuma Incheon (osnovni i alternativni scenarij)."),
  H3("Klauzule iz teksta Ugovora koje zaslužuju ekspertsku pravnu provjeru"),
  Bullet("čl. 3 (Promjena u vlasničkoj strukturi) - mehanizmi protiv preprodaje SPV-a."),
  Bullet("čl. 6 (Vlasništvo) - šta SPV stiče u vlasništvo."),
  Bullet("čl. 10 (Slučajevi zagađenja) - ko snosi rizik istorijskog zagađenja."),
  Bullet("čl. 18 (Kriterijumi učinka i nivoi usluga) - operativni standardi."),
  Bullet("čl. 21 (Osiguranje) - iznosi i obim."),
  Bullet("čl. 22 (Podugovaranje) - ograničenja i nadzor."),
  Bullet("čl. 25 (Radovi) i čl. 26 (Radovi na zadržanom zemljištu) - obim, kvalitet, vremenski plan."),
  Bullet("čl. 28 (Izmjena) - mehanizam Change Order-a koji može povećati troškove."),
  Bullet("čl. 32 (Porezi, troškovi i plaćanja) - poreska struktura."),
  Bullet("čl. 33 (Finansiranje) - struktura kredita, hedžing, naknade."),
  Bullet("čl. 43 (Predaja nakon očekivanog isteka) - stanje imovine na kraju koncesije."),
  Bullet("čl. 45 (Obezbjeđenje za radove prilikom predaje) - garancija za stanje imovine."),
  H3("Vanjske provjere"),
  Bullet("Pravna analiza usklađenosti sa Ustavom Crne Gore (član 13 - jezik, član 91 - državna imovina)."),
  Bullet("Analiza usklađenosti sa Zakonom o koncesijama, Zakonom o državnoj imovini, Zakonom o vazdušnom saobraćaju."),
  Bullet("Pravna analiza odluka u tenderskom postupku - posebno postupanje sa prigovorima drugorangiranog ponuđača."),
  Bullet("Komparativna analiza sa koncesijama aerodroma u Hrvatskoj, Bugarskoj, Albaniji, Sjevernoj Makedoniji."),
  Bullet("Fiskalna projekcija u različitim scenarijima."),
  PageBr(),
];

const finalSection = [
  H1("Završna napomena"),
  P("Ovaj dokument je pripremljen kao nezavisna analiza u javnom interesu, na osnovu cjelovitog teksta Predloga Vlade i Nacrta Ugovora o koncesiji za aerodrome Podgorica i Tivat. Cilj nije bio da diskredituje koncesioni model niti konkretne ponuđače, već da pruži alat za informisanu raspravu - parlamentarnu, novinarsku i civilnu - o ugovoru koji će uticati na Crnu Goru tokom narednih 30 godina."),
  P("Koncesija strateških aerodroma je legitiman instrument moderne državne politike kada se izvrši sa pažnjom, transparentnošću i sa očuvanjem javnog interesa. Aktuelni Nacrt sadrži klauzule koje, ako se otklone ili izmijene u skladu sa preporukama iz ovog dokumenta, mogu omogućiti da Crna Gora ostvari ekonomske koristi od privatizacije bez sistemske izloženosti rizicima koji u dugom roku premašuju te koristi."),
  P("Konačna odluka leži na Skupštini Crne Gore, koja je suveren predstavnik građana. Ovaj dokument želi tom suverenom tijelu pružiti čisto čitanje teksta - i ničega više."),
  Spacer(), Rule(), Spacer(),
  P("Pravna napomena: Analiza nije pravni savjet niti zvanični ekspertski nalaz u smislu Zakona o sudskim vještacima Crne Gore. Sve klauzulske reference odnose se na Nacrt Ugovora koji je sastavni dio Predloga Vlade. Za formalne pravne korake nadležno je kvalifikovano lice - advokat ili sudski vještak ekonomske/pravne struke - registrovano u skladu sa zakonima Crne Gore.", { italic: true }),
  P("Pripremljeno: maj 2026. Verzija: 1.0 - prvi nezavisni pregled.", { italic: true }),
];

const buildFindingSection = (title, findings, intro) => {
  const out = [H1(title)];
  if (intro) out.push(P(intro));
  out.push(Spacer());
  for (const f of findings) { out.push(f); out.push(Spacer()); }
  out.push(PageBr());
  return out;
};

const redSection = buildFindingSection(
  "4. Registar odstupanja - CRVENI signali (10)",
  redFindings,
  "Sljedećih 10 stavki predstavlja materijalna odstupanja od standarda međunarodnog projektnog finansiranja i Zakona o koncesijama Crne Gore. Sve zaslužuju eskalaciju i izmjenu prije potpisivanja Ugovora."
);
const yellowSection = buildFindingSection(
  "5. Registar odstupanja - ŽUTI signali (10)",
  yellowFindings,
  "Sljedećih 10 stavki nisu deal-breakeri, ali su predmet pregovaranja. U dobro vođenoj pregovaračkoj fazi, najmanje polovina ovih stavki može biti riješena u korist Crne Gore."
);
const greenSection = buildFindingSection(
  "6. Registar odstupanja - ZELENI signali (6)",
  greenFindings,
  "Sljedećih 6 stavki predstavljaju klauzule koje su u skladu sa međunarodnim standardima i javnim interesom Crne Gore. Treba ih sačuvati u finalnom Ugovoru."
);

const doc = new Document({
  creator: "Nezavisna analiza koncesije aerodroma",
  title: "Analiza Ugovora o koncesiji za aerodrome Podgorica i Tivat",
  description: "Klauzula-po-klauzula analiza sa preporukama za redline, iz perspektive nezavisnog/parlamentarnog/civilnog pregleda",
  styles: {
    default: { document: { run: { font: FONT, size: 22 } } },
    paragraphStyles: [
      { id: "Heading1", name: "Heading 1", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { font: FONT, size: 36, bold: true, color: COLOR_INK },
        paragraph: { spacing: { before: 360, after: 200 }, outlineLevel: 0 } },
      { id: "Heading2", name: "Heading 2", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { font: FONT, size: 28, bold: true, color: COLOR_INK },
        paragraph: { spacing: { before: 300, after: 160 }, outlineLevel: 1 } },
      { id: "Heading3", name: "Heading 3", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { font: FONT, size: 24, bold: true, color: COLOR_INK },
        paragraph: { spacing: { before: 240, after: 120 }, outlineLevel: 2 } },
    ]
  },
  numbering: {
    config: [
      { reference: "bul", levels: [{ level: 0, format: LevelFormat.BULLET, text: "•", alignment: AlignmentType.LEFT,
        style: { paragraph: { indent: { left: 540, hanging: 360 } } } }] },
      { reference: "num", levels: [{ level: 0, format: LevelFormat.DECIMAL, text: "%1.", alignment: AlignmentType.LEFT,
        style: { paragraph: { indent: { left: 540, hanging: 360 } } } }] },
    ]
  },
  sections: [{
    properties: { page: { size: { width: 11906, height: 16838 }, margin: { top: 1440, right: 1440, bottom: 1440, left: 1440 } } },
    headers: { default: new Header({ children: [new Paragraph({
      alignment: AlignmentType.RIGHT,
      border: { bottom: { style: BorderStyle.SINGLE, size: 4, color: COLOR_RULE, space: 4 } },
      children: [new TextRun({ text: "Analiza Ugovora o koncesiji za aerodrome Podgorica i Tivat", font: FONT, size: 18, color: COLOR_MUTED })]
    })] }) },
    footers: { default: new Footer({ children: [new Paragraph({
      alignment: AlignmentType.CENTER,
      children: [
        new TextRun({ text: "Strana ", font: FONT, size: 18, color: COLOR_MUTED }),
        new TextRun({ children: [PageNumber.CURRENT], font: FONT, size: 18, color: COLOR_MUTED }),
        new TextRun({ text: " od ", font: FONT, size: 18, color: COLOR_MUTED }),
        new TextRun({ children: [PageNumber.TOTAL_PAGES], font: FONT, size: 18, color: COLOR_MUTED }),
      ]
    })] }) },
    children: [
      ...titlePage,
      ...disclaimer,
      ...toc,
      ...exec,
      ...context,
      ...procedural,
      ...redSection,
      ...yellowSection,
      ...greenSection,
      ...strategy,
      ...recommendations,
      ...annex,
      ...finalSection,
    ]
  }]
});

Packer.toBuffer(doc).then(buffer => {
  const outPath = "/sessions/admiring-ecstatic-maxwell/mnt/outputs/Analiza_Ugovora_o_koncesiji_aerodroma.docx";
  fs.writeFileSync(outPath, buffer);
  console.log("Document written: " + outPath);
  console.log("Size: " + buffer.length + " bytes");
});
