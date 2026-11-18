// Erstellt das Sanierungs-Handbuch für Fix & Flip als Word-Dokument.
// Aufruf: NODE_PATH=$(npm root -g) node sanierungs_handbuch.js AUSGABE.docx
const fs = require('fs');
const path = require('path');
const { Document, Packer, Paragraph, TextRun, PageBreak, Footer, PageNumber, AlignmentType, TableOfContents } = require('docx');
const { FONT, ACC, t, b, P, H1, H2, H3, BL, NL, gap, table, box, numbering, styles } = require(path.join(__dirname, 'doc_helfer.js'));

const OUT = process.argv[2];
const C = [];
const br = () => C.push(new Paragraph({ children: [new PageBreak()] }));

// Ein Sanierungsschritt: Steckbrief-Tabelle + Ablauf + Tipps + Warnungen
function step(nr, title, meta, ablauf, tipps, warn) {
  C.push(H2(`Schritt ${nr}: ${title}`));
  C.push(table([1800, 7800], [
    ['Steckbrief', title],
    ['Ziel', meta.ziel],
    ['Wer', meta.wer],
    ['Dauer', meta.dauer],
    ['Kosten (Richtwert)', meta.kosten],
  ], { boldFirst: true }));
  C.push(H3('So gehst du vor'));
  ablauf.forEach(a => C.push(BL(a)));
  if (tipps.length) { C.push(gap()); C.push(box('Tipps & Tricks', tipps.map(x => ['✓ ', ...[].concat(x)].map(r => typeof r === 'string' ? t(r) : r)), 'tip')); }
  if (warn.length) { C.push(gap()); C.push(box('Achtung', warn.map(x => ['⚠ ', ...[].concat(x)].map(r => typeof r === 'string' ? t(r) : r)), 'warn')); }
}

// ---------- Titel ----------
C.push(new Paragraph({ spacing: { before: 1800, after: 200 }, children: [new TextRun({ text: 'FIX & FLIP', font: FONT, size: 26, bold: true, color: ACC, characterSpacing: 60 })] }));
C.push(new Paragraph({ spacing: { after: 200 }, children: [new TextRun({ text: 'Sanierungs-Handbuch', font: FONT, size: 56, bold: true, color: '1C2430' })] }));
C.push(P([t('Wie eine Flip-Sanierung Schritt für Schritt abläuft – mit Reihenfolge, Zuständigkeiten, Tipps, Tricks und den Fehlern, die dich ein ganzes Projekt kosten können. Mit Werkzeug- und Materialübersicht.', { size: 24, color: '555555' })]));
C.push(new Paragraph({ spacing: { before: 600, after: 60 }, children: [t('Stand: November 2026 · Begleitdatei: ', { color: '555555' }), b('Werkzeug_und_Materialliste.xlsx', { color: '555555' })] }));
C.push(P([t('Alle Kosten und Zeiten sind Richtwerte (bundesweite Ratgeber 2026) und schwanken regional stark. Normen und Vorschriften sind vereinfacht dargestellt – im Einzelfall Fachbetrieb, Sachverständigen bzw. Behörde fragen. Keine Rechtsberatung.', { size: 18, italics: true, color: '777777' })]));
br();
C.push(H1('Inhalt'));
C.push(new TableOfContents('Inhalt', { hyperlink: true, headingStyleRange: '1-2' }));
C.push(P([t('Das Inhaltsverzeichnis füllt Word beim Öffnen (ggf. Rechtsklick → „Feld aktualisieren“).', { size: 18, italics: true, color: '777777' })]));
br();

// ---------- 1 Regeln ----------
C.push(H1('1. Die goldenen Regeln einer Flip-Sanierung'));
C.push(P('Eine Flip-Sanierung ist keine Traumhaus-Sanierung. Du sanierst nicht für dich, sondern für den Käufer, der in Damme ein solides, helles, energetisch ordentliches Haus sucht – und zwar so schnell und so planbar wie möglich.'));
C.push(NL([b('Reihenfolge schlägt Tempo. '), t('Grob vor fein, oben vor unten, Technik vor Oberfläche, nass vor trocken. Wer die Reihenfolge bricht, reißt fertige Arbeit wieder auf.')], 'num'));
C.push(NL([b('Erst die Ursache, dann die Optik. '), t('Feuchte, undichtes Dach, marode Leitungen zuerst. Eine neue Tapete auf einer nassen Wand ist rausgeworfenes Geld – und beim Verkauf ein Haftungsrisiko.')], 'num'));
C.push(NL([b('Sanieren auf Nachbarschafts-Niveau. '), t('Käufer zahlen für ein Haus in Damme-West keinen Aufpreis für Designer-Armaturen. Investiere in das, was sie sehen und was im Gutachten steht.')], 'num'));
C.push(NL([b('Ein Flip-Standard für alle Projekte. '), t('Immer dieselbe Wandfarbe, derselbe Boden, dieselbe Fliese, dasselbe Schalterprogramm. Das spart Entscheidungen, Zeit und Geld (Mengenrabatt) – und es sieht trotzdem neu aus.')], 'num'));
C.push(NL([b('Nichts verschließen, was nicht geprüft und fotografiert ist. '), t('Bevor Wände und Böden zugehen: Druckprüfung, Messprotokolle, Fotos aller Leitungen. Das ist deine Versicherung beim Verkauf.')], 'num'));
C.push(NL([b('Kein Geld vor Leistung. '), t('Zahle nur für erbrachte, abgenommene Arbeit. Vorkasse ist der häufigste Weg, Geld und Zeit zu verlieren.')], 'num'));
C.push(NL([b('Zeit ist Geld. '), t('Jeder Monat kostet Zinsen, Versicherung, Strom und Nerven. Ein guter Bauzeitenplan ist mehr wert als ein billiges Angebot.')], 'num'));
C.push(gap());
C.push(H2('Die Reihenfolge auf einen Blick'));
C.push(table([700, 3300, 3100, 2500], [
  ['Nr.', 'Schritt', 'Was passiert', 'Typische Dauer (120-m²-Haus)'],
  ['A', 'Vor dem Kauf', 'Bestandsaufnahme, Kosten, Go/No-Go', '1–2 Wochen'],
  ['B', 'Planung', 'Angebote, Bauzeitenplan, Genehmigungen, Versicherung', '4–8 Wochen (vor Übergabe)'],
  ['1', 'Sichern, dokumentieren, entrümpeln', 'Schlösser, Zähler, Fotos, Container', '1 Woche'],
  ['2', 'Schadstoffe', 'Asbest, alte Mineralwolle, Teer-Kleber prüfen/entfernen', '0–2 Wochen'],
  ['3', 'Rückbau / Entkernung', 'Böden, Fliesen, Bäder, nicht tragende Wände raus', '1–2 Wochen'],
  ['4', 'Gebäudehülle', 'Dach, Fenster, Abdichtung, Trocknung', '1–4 Wochen'],
  ['5', 'Rohinstallation', 'Abwasser → Wasser/Heizung → Elektro', '2–3 Wochen'],
  ['6', 'Innenausbau grob', 'Trockenbau, Dämmung, Putz, Estrich', '2–4 Wochen (+ Trocknung)'],
  ['7', 'Bad & Fliesen', 'Abdichtung, Fliesen, Vorwand', '1–2 Wochen'],
  ['8', 'Feininstallation', 'Schalter, Leuchten, WC, Armaturen, Heizkörper, Prüfprotokolle', '1 Woche'],
  ['9', 'Maler & Türen', 'Spachteln, Streichen, Innentüren', '1–2 Wochen'],
  ['10', 'Böden', 'Vinyl/Laminat, Sockelleisten', '1 Woche'],
  ['11', 'Außen', 'Fassade, Garten, Eingang, Licht', '1 Woche (parallel möglich)'],
  ['12', 'Abschluss', 'Endreinigung, Abnahme, Mängel, Dokumentation, Energieausweis', '1 Woche'],
  ['13', 'Home Staging & Fotos', 'Möblieren, fotografieren, inserieren', '2–3 Tage'],
], { zebra: true }));
C.push(P([t('Für ein Haus mit mittlerer Sanierung (Bad, Elektrik, Böden, Maler, teils Fenster) sind 10–14 Wochen Bauzeit realistisch – wenn die Handwerker vor dem Start feststehen.', { italics: true })], { spacing: { before: 80 } }));

// ---------- 2 Phase A ----------
br();
C.push(H1('2. Phase A – Vor dem Kauf: Bestandsaufnahme'));
C.push(P('Hier entscheidet sich, ob du Geld verdienst. Was du vor dem Kauf nicht findest, zahlst du später aus deinem Gewinn.'));
C.push(H2('Was du (mit dem Bausachverständigen) prüfen musst'));
C.push(table([2300, 4200, 3100], [
  ['Bauteil', 'Worauf achten', 'Warnsignal = teuer'],
  ['Dach', 'Alter der Eindeckung, Dachstuhl (Holzwurm, Feuchte), Dämmung, Dachrinnen', 'Eternit/Wellplatten (Asbest), nasse Sparren, durchhängende Linien'],
  ['Keller & Feuchte', 'Feuchtigkeitsmessung an Außenwänden, Salzausblühungen, Geruch', 'Nasse Wände bis 1 m Höhe, frische Farbe nur im Keller (überstrichen?)'],
  ['Statik', 'Risse in Wänden/Decken, Setzungen, Durchbiegungen', 'Treppenförmige Risse, Türen klemmen, Risse > 2 mm'],
  ['Elektrik', 'Baujahr Verteilung, FI-Schutzschalter, Leitungen (2- oder 3-adrig), Sicherungen', 'Schraubsicherungen, keine Schutzleiter, Stoffkabel'],
  ['Wasser & Abwasser', 'Leitungsmaterial, Wasserdruck, Abflüsse, Grundleitung (Kamera!)', 'Bleirohre (bis ca. 1973), verzinkte Rohre mit Rost, Wurzeln in der Grundleitung'],
  ['Heizung', 'Alter, Typ, Wartungsprotokolle, Heizkörper, Schornstein', 'Kessel über 30 Jahre, Nachtspeicheröfen (Asbest möglich)'],
  ['Fenster & Türen', 'Verglasung, Rahmen, Dichtungen, Rollläden', 'Einfachverglasung, Holz morsch'],
  ['Schadstoffe', 'Baujahr bis 1993: Asbest möglich (Floor-Flex-Platten, Kleber, Putze, Fensterbänke, Eternit)', 'Alte schwarze Kleber unter Parkett (PAK), alte gelbe Mineralwolle'],
  ['Unterlagen', 'Baugenehmigungen, Grundrisse, Energieausweis, Rechnungen', 'Anbau oder Dachausbau ohne Genehmigung'],
], { zebra: true }));
C.push(gap());
C.push(box('Tipps & Tricks', [
  '✓ Nimm zur Besichtigung eine Taschenlampe, ein Feuchtemessgerät, einen Zollstock, eine Kamera und einen Spannungsprüfer mit. Fotografiere jeden Raum, Zählerschrank, Heizung, Dachboden und Keller.',
  '✓ Lass dir die Grundleitung per Kamera befahren (ca. 200–500 €). Eine marode Grundleitung unter der Bodenplatte kann 10.000–20.000 € kosten.',
  '✓ Bei Baujahr bis 1993 und Verdacht auf Asbest: Materialproben vor dem Kauf ins Labor geben (ca. 50–100 € pro Probe). Der Befund ist ein starkes Verhandlungsargument.',
  '✓ Bewerte jeden Mangel sofort mit einem Preis. Am Ende der Besichtigung steht eine Summe, nicht ein Gefühl.',
], 'tip'));
C.push(gap());
C.push(box('Achtung', [
  '⚠ Besichtigungen im Sommer verstecken Heizungs- und Feuchteprobleme, frisch gestrichene Kellerwände verstecken Nässe. Immer messen, nie nur schauen.',
  '⚠ Wenn der Verkäufer keinen Zugang zu Dachboden oder Keller ermöglicht: Preis massiv nach unten oder Finger weg.',
], 'warn'));
C.push(H2('Go oder No-Go'));
C.push(BL('Sanierungskosten nach Gewerken aufschreiben, 15–20 % Puffer dazu, in den Flip-Rechner eintragen.'));
C.push(BL('Liegt der Kaufpreis über dem „Max. Kaufpreis“: nachverhandeln – mit den konkreten Mängeln als Begründung – oder gehen.'));
C.push(BL([b('No-Go beim ersten Projekt: '), t('großflächiger Schwamm oder Hausbock im Dachstuhl, Setzungsrisse, nasser Keller ohne klare Ursache, Asbest in großen Flächen, Denkmalschutz, ungenehmigte Bauteile.')]));

// ---------- 3 Phase B ----------
br();
C.push(H1('3. Phase B – Planung vor dem ersten Hammerschlag'));
C.push(P('Die Zeit zwischen Notartermin und Übergabe (meist 4–8 Wochen) nutzt du, um alles vorzubereiten. Am Tag der Schlüsselübergabe sollen die Handwerker schon Termine haben.'));
C.push(H2('Die Planungs-Checkliste'));
[
  [b('Aufmaß und Raumbuch: '), t('Jeden Raum ausmessen (Boden, Wände, Fenster, Türen). Pro Raum aufschreiben, was gemacht wird. Daraus entstehen deine Mengen und die Ausschreibung.')],
  [b('Leistungsbeschreibung pro Gewerk: '), t('Was genau soll gemacht werden, mit welchem Material, in welcher Qualität? Nur dann sind Angebote vergleichbar.')],
  [b('Drei Angebote je Gewerk: '), t('Nicht nur auf den Preis schauen: Termin, Referenzen, Festpreis oder Stundenlohn, Zahlungsplan, Gewährleistung.')],
  [b('Festpreis vereinbaren, wo es geht: '), t('Pauschalpreise mit klar beschriebener Leistung. Stundenlohnarbeiten nur mit Obergrenze.')],
  [b('Bauzeitenplan: '), t('Welche Firma wann, in welcher Reihenfolge, mit Puffer (siehe Kapitel 5).')],
  [b('Genehmigungen: '), t('Tragende Wände, Anbauten, Dachausbau, Nutzungsänderung, Aufteilung in Wohnungen → Bauamt. Container oder Gerüst auf öffentlicher Straße → Sondernutzungserlaubnis bei der Stadt.')],
  [b('Förderung: '), t('Energieberater einbinden. Förderanträge (KfW/BAFA) müssen in der Regel vor Auftragsvergabe bzw. Baubeginn gestellt werden – sonst ist die Förderung weg.')],
  [b('Lüftungskonzept: '), t('Werden viele Fenster getauscht (Faustregel: mehr als ein Drittel), braucht es ein Lüftungskonzept nach DIN 1946-6. Sonst droht Schimmel, weil das Haus plötzlich dicht ist.')],
  [b('Versicherungen: '), t('Wohngebäudeversicherung über Leerstand und Umbau informieren, Bauherrenhaftpflicht abschließen, ggf. Bauleistungsversicherung. Private Helfer (Freunde, Familie) müssen bei der BG BAU angemeldet werden.')],
  [b('Baustrom, Bauwasser, Container: '), t('Strom und Wasser auf dich ummelden, ggf. Baustromverteiler. Container für Bauschutt und Holz/Mischabfall bestellen.')],
  [b('Material: '), t('Lange Lieferzeiten früh bestellen (Fenster, Haustür, Wärmepumpe oft 6–12 Wochen). Den Flip-Standard festlegen.')],
].forEach(x => C.push(NL(x, 'num2')));
C.push(gap());
C.push(box('Tipps & Tricks', [
  '✓ Gute Handwerker sind im Oldenburger Münsterland Monate ausgebucht. Sprich sie schon an, wenn du ein Objekt ernsthaft prüfst – nicht erst nach dem Notartermin.',
  '✓ Ein Generalunternehmer (GU) kostet 10–20 % Aufschlag, nimmt dir aber die Koordination ab. Für das erste Projekt kann das sinnvoll sein; ab dem zweiten koordinierst du Einzelgewerke selbst.',
  '✓ Frag jeden Handwerker: „Was würden Sie an meiner Stelle weglassen?“ Gute Handwerker sparen dir Geld.',
], 'tip'));

// ---------- 4 Phase C ----------
br();
C.push(H1('4. Phase C – Die Sanierung in 13 Schritten'));
C.push(P('Jeder Schritt hat einen Steckbrief, einen Ablauf, Tipps und Warnungen. Die Kosten sind Richtwerte für ein Haus mit ca. 120 m² Wohnfläche.'));

step(1, 'Sichern, dokumentieren, entrümpeln',
  { ziel: 'Haus übernehmen, Ist-Zustand festhalten, leer und sicher machen', wer: 'Selbst (+ Entrümpelungsfirma bei viel Inventar)', dauer: '2–7 Tage', kosten: 'Container 300–600 € je Stück, Entrümpelung 1.500–5.000 €' },
  [
    'Schlösser tauschen (Schließzylinder), Zählerstände (Strom, Wasser, Gas) fotografieren und ummelden.',
    'Jeden Raum fotografieren und filmen: Vorher-Bilder für Dokumentation, Versicherung und spätere Vermarktung („vorher/nachher“).',
    'Wasser im Winter: Heizung auf Frostschutz oder Leitungen entleeren.',
    'Entrümpeln, Sperrmüll und Wertstoffe trennen. Brauchbares (Möbel, Geräte) verkaufen oder verschenken.',
    'Abfall getrennt sammeln: Bauschutt, Holz, Metall, Gipskarton, Mischabfall – gemischt ist es am teuersten.',
  ],
  ['Alte Heizkörper, Kupfer und Metall bringen beim Schrotthändler Geld.', 'Ein abschließbarer Raum oder Container für Werkzeug und Material spart Ärger.'],
  ['Bei Leerstand in der kalten Jahreszeit verlangen Gebäudeversicherungen meist, dass geheizt oder das Wasser abgesperrt und entleert wird. Ein Frostschaden ohne diese Vorsorge wird oft nicht bezahlt.']);

step(2, 'Schadstoffe prüfen und entfernen',
  { ziel: 'Gefahrstoffe sicher erkennen und fachgerecht entfernen, bevor Staub entsteht', wer: 'Labor für Proben; Entfernung nur durch Fachfirma mit Sachkunde (TRGS 519)', dauer: '0–2 Wochen', kosten: 'Proben 50–100 € je Stück; Asbestsanierung stark mengenabhängig (einige Tausend bis über 20.000 €)' },
  [
    'Verdachtsmaterialien auflisten: Floor-Flex-Platten und schwarze Kleber unter Böden, Eternit (Dach, Fassade, Blumenkästen), Nachtspeicheröfen, Fensterbänke, Spachtel und Putze, alte Mineralwolle.',
    'Proben nehmen lassen und ins Labor geben, bevor gebohrt, geschliffen oder abgerissen wird.',
    'Positive Befunde: Angebot einer zugelassenen Fachfirma, Entsorgung mit Nachweis.',
  ],
  ['Bei Baujahr bis 1993 grundsätzlich mit Schadstoffen rechnen und das in der Kalkulation berücksichtigen.', 'Der Laborbefund gehört in die Verkaufsmappe – Käufer sehen: hier wurde sauber gearbeitet.'],
  [[b('Asbest nie selbst schleifen, flexen, bohren oder brechen. '), t('Die Fasern verursachen Jahrzehnte später schwere Krankheiten, verseuchen das ganze Haus und können eine teure Dekontamination auslösen. Unsachgemäßes Arbeiten mit Asbest ist zudem verboten.')],
    'Alte gelbliche Mineralwolle (vor ca. 2000) kann krebserregend sein: nur mit Schutzausrüstung (FFP3, Overall) und staubarm ausbauen, verpackt entsorgen.']);

step(3, 'Rückbau und Entkernung',
  { ziel: 'Alles raus, was erneuert wird – bis auf den Teil der Substanz, der bleibt', wer: 'Gut in Eigenleistung (mit Helfern); Abbruch tragender Teile nur nach Statik', dauer: '1–2 Wochen', kosten: 'Eigenleistung: Container + Mietgeräte ca. 1.500–4.000 €; Abbruchfirma ca. 3.000–10.000 €' },
  [
    'Strom in den Rückbau-Bereichen abschalten lassen (Elektriker), Wasser absperren.',
    'Von innen nach außen, von oben nach unten: Tapeten, Böden und Fliesen, alte Bäder und Küche, Deckenverkleidungen.',
    'Nicht tragende Wände erst entfernen, wenn klar ist, dass sie nicht tragend sind (Statiker).',
    'Nach dem Rückbau: zweite Begehung mit dem Sachverständigen – jetzt sieht man, was hinter den Verkleidungen ist. Kostenschätzung anpassen.',
  ],
  ['Mietgeräte statt Kaufen: Abbruchhammer, Fliesenstripper, Wandschlitzfräse. Für einen Tag kosten sie oft unter 100 €.', 'Fenster und Türöffnungen beim Rückbau mit Folie schützen, wenn sie bleiben.'],
  [[b('Tragende Wände: '), t('Eine entfernte tragende Wand ohne Statik und Abfangung kann Decken einstürzen lassen. Statiker + Genehmigung, Stahlträger ggf. 5.000–20.000 €.')],
    'Schornsteine, Lüftungsschächte und Leitungen in Wänden vorher orten (Leitungssucher). Eine angebohrte Gasleitung ist lebensgefährlich.']);

step(4, 'Gebäudehülle: Dach, Fenster, Abdichtung',
  { ziel: 'Das Haus ist dicht – von oben, von der Seite und von unten', wer: 'Fachbetriebe (Dachdecker, Fensterbauer, Abdichtung)', dauer: '1–4 Wochen (+ Lieferzeit Fenster 6–12 Wochen)', kosten: 'Fenster ca. 12.000–22.500 € (15 Stück); Dach 150–300 €/m² Dachfläche; Kellerabdichtung 15.000–40.000 €' },
  [
    'Dach zuerst: undichte Stellen, Dachrinnen und Fallrohre. Nur wenn nötig neu eindecken – dann gleich dämmen.',
    'Fenster und Haustür einbauen (luftdicht nach RAL-Montage). Fensterbänke innen/außen gleich mitplanen.',
    'Feuchtigkeit an der Ursache bekämpfen: Außenabdichtung, Drainage, Horizontalsperre – nicht nur innen Sperrputz.',
    'Bauteile trocknen lassen, ggf. Bautrockner (mieten) einsetzen.',
  ],
  ['Dämmung der obersten Geschossdecke ist eine der günstigsten Maßnahmen für eine bessere Energieklasse (ca. 30–60 €/m²) und geht in Eigenleistung.', 'Bei neuen Fenstern an das Lüftungskonzept denken (siehe Phase B).'],
  ['Neue, dichte Fenster in einem alten Haus ohne Lüftung = Schimmel im ersten Winter. Das ist einer der häufigsten Sanierungsschäden.', 'Feuchte Wände innen nur zu überputzen, ohne die Ursache zu beheben, ist beim Verkauf arglistig, wenn du davon weißt – Haftungsrisiko.']);

step(5, 'Rohinstallation: Abwasser, Wasser, Heizung, Elektro',
  { ziel: 'Alle Leitungen liegen, sind geprüft und dokumentiert, bevor Wände und Böden zugehen', wer: 'Fachbetriebe: Sanitär/Heizung (SHK), Elektriker. Schlitze stemmen ggf. Eigenleistung nach Absprache', dauer: '2–3 Wochen', kosten: 'Elektrik 90–140 €/m² Wohnfläche; Bad-Rohinstallation in den Badkosten; Wärmepumpe 25.000–40.000 € abzgl. Förderung' },
  [
    'Reihenfolge: Abwasser zuerst (braucht Gefälle), dann Trinkwasser und Heizung, zuletzt Elektro (flexibel zu verlegen).',
    'Elektro: neuer Zählerschrank/Unterverteilung nach aktuellen Regeln, FI-Schutzschalter, ausreichend Steckdosen, Netzwerk/Leerrohre. Zählerschrank-Änderungen beim Netzbetreiber früh anmelden (Wochen Vorlauf).',
    'Trinkwasser: alte Blei- oder verzinkte Leitungen komplett ersetzen. Danach Druckprüfung mit Protokoll.',
    'Heizung: Heizlast berechnen lassen, bevor eine Wärmepumpe bestellt wird. Heizkörper ggf. größer, Vorlauftemperatur senken, hydraulischer Abgleich.',
    'Alles fotografieren und vermaßen, bevor es verschlossen wird (Leitungsplan für die Verkaufsmappe).',
  ],
  ['Steckdosen sind billig im Rohbau und teuer danach: in jedem Raum lieber zwei mehr. Käufer zählen sie.', 'LED-Deckenspots in Küche, Flur und Bad wirken hochwertig und kosten wenig, wenn die Leitungen jetzt verlegt werden.', 'Leerrohre zu Dach (PV), Garage (Wallbox) und jedem Raum (Netzwerk) sind ein starkes Verkaufsargument.'],
  ['Elektro nur teilweise erneuern („Mischanlage“) ist ein Brandrisiko und fällt bei jeder Prüfung auf. Wenn, dann richtig.', 'Gas- und Elektroanschlüsse sind ausschließlich Sache eingetragener Fachbetriebe – auch wegen Versicherung und Haftung.', 'Eine Wärmepumpe ohne Heizlastberechnung läuft im alten Haus oft ineffizient – hohe Stromkosten sprechen sich beim Käufer herum.']);

step(6, 'Innenausbau grob: Trockenbau, Dämmung, Putz, Estrich',
  { ziel: 'Wände und Decken geschlossen, Böden eben – bereit für Fliesen und Oberflächen', wer: 'Trockenbau und Putz teils Eigenleistung; Estrich Fachbetrieb', dauer: '2–4 Wochen + Trocknungszeit', kosten: 'Trockenbau 40–80 €/m² Wandfläche; Estrich 25–45 €/m²' },
  [
    'Schlitze verschließen, Trockenbauwände stellen, Vorwände für WC und Waschtisch.',
    'Dachschrägen und oberste Geschossdecke dämmen, Dampfbremse luftdicht verkleben.',
    'Putz ausbessern oder neu, Wände für den Maler vorbereiten.',
    'Estrich einbringen bzw. Unebenheiten mit Ausgleichsmasse egalisieren.',
  ],
  ['Ausgleichsmasse statt neuem Estrich, wo es reicht: schneller und billiger.', 'Während der Trocknung parallel außen (Garten, Fassade) arbeiten – so geht keine Zeit verloren.'],
  [[b('Estrich braucht Zeit: '), t('Faustregel Zementestrich ca. 1 Woche pro cm Dicke (bis 4 cm), danach deutlich länger. Vor dem Bodenbelag Restfeuchte messen lassen (CM-Messung). Zu früh verlegt = Schimmel, gewölbte Böden, Ablösung.')],
    'Eine undichte Dampfbremse im Dach führt zu Feuchte in der Dämmung – unsichtbar, bis es schimmelt.']);

step(7, 'Bad und Fliesen',
  { ziel: 'Ein zeitloses, dichtes Bad – der wichtigste Raum für den Verkaufspreis', wer: 'SHK-Betrieb + Fliesenleger (Fliesen in Eigenleistung nur mit Übung)', dauer: '1–2 Wochen pro Bad', kosten: '12.000–25.000 € pro Bad' },
  [
    'Vorwandelemente setzen (WC, Waschtisch), Rohinstallation prüfen.',
    'Abdichtung im Spritzwasserbereich und unter der Dusche nach DIN 18534: Dichtschlämme bzw. Flüssigabdichtung, Dichtbänder in allen Ecken, Manschetten an Rohrdurchführungen.',
    'Fliesen verlegen (große Formate, z. B. 60 × 60 cm, matt, hellgrau/beige), Fugen, Silikon.',
    'Bodengleiche Dusche mit Glaswand statt Wanne – wenn genug Platz ist, zusätzlich Wanne.',
  ],
  ['Bodengleiche Dusche + großformatige Fliesen + LED-Spiegel + Handtuchheizkörper = „neues Bad“-Effekt für überschaubares Geld.', 'Weiße Keramik, Armaturen einer soliden Mittelklasse-Marke – keine Billigware, die in einem Jahr tropft.', 'Abdichtung vor dem Fliesen fotografieren (für die Verkaufsmappe).'],
  [[b('Fehlende oder fehlerhafte Abdichtung ist der Klassiker unter den Wasserschäden: '), t('Wasser läuft unbemerkt in Wand und Decke, der Schaden zeigt sich Monate später – oft erst beim Käufer.')],
    'Silikonfugen sind Wartungsfugen, keine Abdichtung.']);

step(8, 'Feininstallation und Prüfungen',
  { ziel: 'Alles funktioniert und ist geprüft', wer: 'Elektriker, SHK-Betrieb', dauer: 'ca. 1 Woche', kosten: 'in den Gewerken enthalten' },
  [
    'Schalter, Steckdosen, Leuchten, Rauchmelder (Pflicht in Schlafräumen, Kinderzimmern und Fluren).',
    'WC, Waschtisch, Armaturen, Duschwand, Heizkörper, Thermostate.',
    'Elektro: Prüf- bzw. Übergabeprotokoll vom Elektriker anfordern. Heizung: hydraulischer Abgleich, Inbetriebnahmeprotokoll.',
  ],
  ['Ein einheitliches Schalterprogramm in Reinweiß wirkt sauber und modern.', 'Alle Protokolle und Rechnungen sammeln – das ist die halbe Verkaufsmappe.'],
  ['Ohne Prüfprotokoll keine Abnahme und keine Schlusszahlung an den Elektriker.']);

step(9, 'Maler und Innentüren',
  { ziel: 'Helle, glatte, frische Räume', wer: 'Gut in Eigenleistung; große Flächen mit Airless-Gerät (mieten)', dauer: '1–2 Wochen', kosten: 'Material ca. 8–15 €/m² Wohnfläche; Innentür mit Zarge 250–500 € je Stück' },
  [
    'Abkleben und abdecken, Untergrund prüfen, grundieren (Tiefengrund).',
    'Spachteln und schleifen bis zur gewünschten Glätte (meist Q2, sichtbare Wände Q3).',
    'Zwei Anstriche in hellem Weiß oder sehr hellem Grau/Greige, Decken weiß.',
    'Innentüren und Zargen (weiß, CPL) montieren, Drückergarnituren einheitlich.',
  ],
  ['Eine einzige Wandfarbe für das ganze Haus: einfacher, günstiger, Nachbessern ohne Farbunterschiede.', 'Alte, aber intakte Türen lackieren und neue Drücker montieren statt alles zu tauschen.'],
  ['Auf feuchte oder versalzene Wände streichen bringt Abplatzungen und Flecken – Ursache vorher beheben.']);

step(10, 'Bodenbeläge',
  { ziel: 'Ein durchgehender, heller, robuster Boden', wer: 'Eigenleistung gut möglich (Klick-Vinyl)', dauer: 'ca. 1 Woche', kosten: 'Material 25–50 €/m², Verlegung Fachbetrieb 15–30 €/m²' },
  [
    'Untergrund eben, trocken (CM-Messung) und sauber.',
    'Durchgehend einen Belag im Wohnbereich verlegen (z. B. Klick-Vinyl/Rigid in heller Eiche), nur Bad und ggf. Flur gefliest.',
    'Sockelleisten, Übergangsprofile, Silikon/Acryl an Anschlüssen.',
  ],
  ['Rigid-Vinyl ist wasserresistent, leise und schnell verlegt – für Flips ideal.', 'Gleicher Boden in allen Räumen lässt das Haus größer wirken.'],
  ['Auf Fußbodenheizung nur freigegebene Beläge und Unterlagen verwenden.', 'Dehnungsfugen zu Wänden nicht vergessen – sonst wölbt sich der Boden.']);

step(11, 'Außenbereich und erster Eindruck',
  { ziel: 'Der Käufer verliebt sich schon an der Straße', wer: 'Eigenleistung + ggf. Gärtner/Maler', dauer: '3–7 Tage (parallel zu Innenarbeiten)', kosten: '2.000–8.000 €' },
  [
    'Garten zurückschneiden, Rasen, Beete mulchen, Unkraut entfernen.',
    'Pflaster und Fassade reinigen (Hochdruckreiniger), Fassade ggf. streichen.',
    'Haustür (neu oder lackiert), Hausnummer, Briefkasten, Außenleuchte, Klingel einheitlich.',
    'Garage/Carport aufräumen und streichen.',
  ],
  ['Fotos und Besichtigungen beginnen draußen. 2.000 € für den Eingangsbereich bringen oft mehr als 2.000 € im Wohnzimmer.'],
  []);

step(12, 'Abschluss: Reinigung, Abnahme, Dokumentation',
  { ziel: 'Mängelfrei, sauber, alles belegt', wer: 'Selbst + Reinigungsfirma; Energieberater', dauer: 'ca. 1 Woche', kosten: 'Bauendreinigung 3–6 €/m²; Bedarfsausweis 300–600 €' },
  [
    'Abnahme mit jedem Gewerk: gemeinsam begehen, Mängelliste mit Fristen, erst danach Schlusszahlung.',
    'Bauendreinigung (Fenster, Böden, Sanitär).',
    'Neuen Energieausweis (Bedarfsausweis) ausstellen lassen – die bessere Klasse gehört ins Inserat.',
    [b('Verkaufsmappe anlegen: '), t('Vorher-Nachher-Fotos, Rechnungen, Protokolle (Druckprüfung, Elektro, Heizung), Leitungsfotos, Laborbefunde, Grundrisse, Energieausweis.')],
  ],
  ['Die Verkaufsmappe schafft Vertrauen, beschleunigt die Finanzierung des Käufers und schützt dich vor späteren Streitigkeiten.'],
  ['Schlusszahlung vor der Abnahme = du verlierst dein wichtigstes Druckmittel für Mängelbeseitigung.']);

step(13, 'Home Staging und Fotos',
  { ziel: 'Die bestmögliche Präsentation für Inserat und Besichtigung', wer: 'Selbst oder Home-Staging-Anbieter; Fotograf', dauer: '2–3 Tage', kosten: '1.000–3.000 €' },
  [
    'Wohnzimmer, Küche/Essen und Schlafzimmer andeuten (Leihmöbel, Teppich, Pflanzen, Licht).',
    'Profi-Fotos bei Tageslicht, Grundriss-Grafik, ggf. 360°-Rundgang.',
    'Inserat mit Fokus auf Energieklasse, neue Technik, neues Bad, sofort beziehbar.',
  ],
  ['Leere Räume wirken kleiner und kälter. Schon wenige Möbel zeigen, wie viel Platz da ist.'],
  []);

// ---------- 5 Baustellenmanagement ----------
br();
C.push(H1('5. Baustellen-Management'));
C.push(H2('Beispiel-Bauzeitenplan (120 m², mittlere Sanierung)'));
C.push(table([2600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600], [
  ['Gewerk', 'W1', 'W2', 'W3', 'W4', 'W5', 'W6', 'W7', 'W8', 'W9', 'W10', 'W11', 'W12'],
  ['Entrümpeln, Rückbau', '■', '■', '', '', '', '', '', '', '', '', '', ''],
  ['Fenster / Dach', '', '', '■', '■', '', '', '', '', '', '', '', ''],
  ['Elektro Rohinstallation', '', '', '■', '■', '', '', '', '', '', '', '', ''],
  ['Sanitär/Heizung Rohinstallation', '', '', '■', '■', '■', '', '', '', '', '', '', ''],
  ['Trockenbau, Dämmung, Putz', '', '', '', '', '■', '■', '', '', '', '', '', ''],
  ['Estrich/Ausgleich + Trocknung', '', '', '', '', '', '■', '■', '', '', '', '', ''],
  ['Bad: Abdichtung, Fliesen', '', '', '', '', '', '', '■', '■', '', '', '', ''],
  ['Feininstallation', '', '', '', '', '', '', '', '', '■', '', '', ''],
  ['Maler, Türen', '', '', '', '', '', '', '', '■', '■', '', '', ''],
  ['Böden', '', '', '', '', '', '', '', '', '', '■', '', ''],
  ['Außen, Garten', '', '', '', '', '', '■', '■', '', '', '', '', ''],
  ['Abnahme, Reinigung, Staging', '', '', '', '', '', '', '', '', '', '', '■', ''],
  ['Puffer', '', '', '', '', '', '', '', '', '', '', '', '■'],
]));
C.push(H2('So führst du die Baustelle'));
[
  [b('Wöchentliche Baubesprechung '), t('(z. B. montags 7:30 Uhr): Was ist fertig, was kommt, wo hakt es?')],
  [b('Bautagebuch: '), t('Jeden Tag kurz: wer war da, was wurde gemacht, Fotos, Probleme. Unverzichtbar bei Streit und für die Verkaufsmappe.')],
  [b('Zahlungen nur nach Leistung: '), t('Abschläge nach Baufortschritt, keine Vorkasse. Bei größeren Aufträgen 5–10 % bis zur Mängelbeseitigung einbehalten (vertraglich vereinbaren).')],
  [b('Nachträge nur schriftlich: '), t('Jede Zusatzleistung vorher mit Preis bestätigen lassen – „Machen Sie mal“ ist ein Blankoscheck.')],
  [b('Abnahme je Gewerk: '), t('Mit Protokoll. Ab der Abnahme läuft die Gewährleistung (bei Bauwerken nach BGB in der Regel 5 Jahre, bei VOB/B-Verträgen meist 4 Jahre).')],
  [b('Kostenkontrolle: '), t('Eine Tabelle: Budget je Gewerk, beauftragt, bezahlt, offen. Jede Woche aktualisieren.')],
].forEach(x => C.push(BL(x)));

// ---------- 6 Eigenleistung ----------
C.push(H1('6. Eigenleistung: Was du selbst machst – und was nie'));
C.push(table([3200, 3200, 3200], [
  ['Gut selbst machbar', 'Mit Erfahrung / Anleitung', 'Nie selbst'],
  ['Entrümpeln, Rückbau (nicht tragend)', 'Trockenbau, Dämmung oberste Geschossdecke', 'Gas, Heizungsanlage, Schornstein'],
  ['Tapeten entfernen, Spachteln, Streichen', 'Fliesen verlegen (Wand), Silikon', 'Elektro-Anschlüsse, Zählerschrank'],
  ['Klick-Vinyl, Sockelleisten', 'Innentüren montieren', 'Statik, tragende Wände'],
  ['Garten, Pflaster reinigen, Fassade streichen', 'Fensterbänke, kleinere Putzarbeiten', 'Asbest und andere Schadstoffe'],
  ['Möbel- und Küchenmontage, Endreinigung', 'Ausgleichsmasse', 'Abdichtung im Bad (Gewährleistung!), Dach'],
], { zebra: true }));
C.push(P([t('Faustregel: Eigenleistung spart bei einem 120-m²-Haus schnell 10.000–20.000 €. Aber: Deine Zeit ist auch Haltedauer. Wenn du nur abends und am Wochenende arbeitest, kann ein Profi, der es in einer Woche macht, unterm Strich billiger sein.', { italics: true })], { spacing: { before: 100 } }));

// ---------- 7 Tipps ----------
C.push(H1('7. Tipps und Tricks der Profis'));
[
  [b('Flip-Standard-Paket: '), t('Lege einmal fest: Wandfarbe, Boden, Badfliese, Sanitärserie, Armaturen, Schalterprogramm, Innentür, Leuchten. Bei jedem Projekt gleich bestellen.')],
  [b('Großhändler statt Baumarkt: '), t('Über deine Handwerker oder ein eigenes Gewerbekonto beim regionalen Baustoff-, Sanitär- und Elektrogroßhandel kaufst du oft deutlich günstiger und bekommst Lieferung auf die Baustelle.')],
  [b('Just-in-time liefern lassen: '), t('Material erst kurz vor Einbau – weniger Diebstahl, weniger Schäden, mehr Platz.')],
  [b('Mieten statt kaufen: '), t('Teure Spezialmaschinen, die du nur einmal brauchst, immer mieten (siehe Kapitel 10).')],
  [b('Licht verkauft: '), t('Helle Wände, LED-Spots, große Leuchtmittel mit warmweißem Licht – auf Fotos und bei Besichtigungen der größte Unterschied.')],
  [b('Neutral statt persönlich: '), t('Weiß, Hellgrau, Greige, helles Holz. Der Käufer soll sich sein Leben vorstellen, nicht deinen Geschmack.')],
  [b('Vorher-Nachher als Marketing: '), t('Dokumentiere jedes Projekt. Das hilft beim Verkauf, bei der nächsten Bankfinanzierung und für deinen Ruf in Damme.')],
  [b('Handwerker gut behandeln: '), t('Pünktlich zahlen, Kaffee hinstellen, Baustelle vorbereitet übergeben. Wer gut zahlt, bekommt beim nächsten Projekt schneller einen Termin.')],
  [b('Immer eine Exit-Variante: '), t('Plane die Ausstattung so, dass das Haus auch als Mietobjekt funktioniert (robuste Böden, pflegeleichtes Bad).')],
].forEach(x => C.push(BL(x)));

// ---------- 8 Katastrophen ----------
br();
C.push(H1('8. Die 16 größten Fehler und Katastrophen'));
C.push(P('Die meisten Flips scheitern nicht an einer großen Fehlentscheidung, sondern an einer Kette kleiner Fehler. Diese Liste solltest du vor jedem Projekt einmal durchgehen.'));
C.push(table([2400, 3300, 3900], [
  ['Was passiert', 'Folge', 'So vermeidest du es'],
  ['Schimmel nach der Sanierung', 'Neue dichte Fenster, nasser Estrich, keine Lüftung → Schimmel beim Käufer, Haftung', 'Lüftungskonzept, CM-Messung vor Bodenbelag, Bautrockner, Ursache vor Optik'],
  ['Asbest beim Rückbau entdeckt', 'Baustopp, Fachfirma, Dekontamination, Wochen Verzug, Gesundheitsgefahr', 'Proben vor dem Kauf (Baujahr bis 1993), nie selbst flexen'],
  ['Tragende Wand entfernt', 'Risse, Einsturzgefahr, Statiker, Stahlträger, Haftung', 'Statiker vorher, Genehmigung, Abfangung'],
  ['Wasserschaden im Bad', 'Fehlende Abdichtung → Schäden in Wand/Decke, oft erst beim Käufer', 'Abdichtung nach Norm, Fotos, Druckprüfung, Fachbetrieb'],
  ['Frost im leeren Haus', 'Geplatzte Leitungen, Wasser in allen Etagen, Versicherung zahlt evtl. nicht', 'Heizen oder Leitungen entleeren, Versicherung informieren, regelmäßig kontrollieren'],
  ['Diebstahl auf der Baustelle', 'Werkzeug, Kupfer, Geräte weg; Verzug', 'Abschließen, Material just in time, Werkzeug abends mitnehmen, ggf. Kamera'],
  ['Handwerker-Insolvenz nach Vorkasse', 'Geld weg, Arbeit nicht gemacht, neue Firma teurer', 'Keine Vorkasse, Zahlung nach Leistung, Referenzen prüfen'],
  ['Schwarzarbeit', 'Keine Gewährleistung, keine Versicherung bei Unfall/Schaden, Strafen, keine Rechnung für die Steuer', 'Nur mit Rechnung, Fachbetriebe für Technik'],
  ['„Wenn wir schon dabei sind …“', 'Budget um 30–50 % überschritten, Marge weg', 'Leistungsumfang vorher festlegen, Änderungen nur mit neuem Rechner-Ergebnis'],
  ['Gewerke-Stau und Lieferzeiten', 'Fenster nach 12 Wochen, Elektriker wartet auf Netzbetreiber, jeder Monat kostet', 'Früh bestellen, Bauzeitenplan mit Puffer, Handwerker vor Kauf ansprechen'],
  ['Marode Grundleitung', '10.000–20.000 € für Aufgraben und Erneuern', 'Kamerabefahrung vor dem Kauf'],
  ['Halbe Elektrik', 'Mischanlage, Brandgefahr, fällt bei Prüfung/Gutachten des Käufers auf', 'Komplett erneuern oder fachgerecht prüfen und dokumentieren'],
  ['Falsche Wärmepumpe', 'Hohe Stromkosten, kalte Räume, unzufriedener Käufer', 'Heizlastberechnung, Heizkörper prüfen, Fachplaner'],
  ['Überkapitalisierung', 'Luxusausstattung, die der Markt in Damme nicht bezahlt', 'Nachbarschafts-Niveau, Flip-Standard, Vergleichsobjekte'],
  ['Unfall auf der Baustelle', 'Verletzung, Ausfall, Haftung; private Helfer ohne Versicherungsschutz', 'Schutzausrüstung, sichere Leitern/Gerüst, Helfer bei der BG BAU anmelden'],
  ['Fehlende Genehmigung', 'Rückbauverfügung, Bußgeld, Käufer springt ab', 'Bauamt fragen, bevor gebaut wird'],
], { zebra: true }));
C.push(gap());
C.push(box('Die drei Fragen vor jeder Entscheidung auf der Baustelle', [
  '1. Bezahlt der Käufer in Damme dafür mehr, als es kostet?',
  '2. Verlängert es die Bauzeit – und was kostet das?',
  '3. Kann es später einen Schaden oder eine Haftung auslösen, wenn ich es weglasse oder billig mache?',
], 'key'));

// ---------- 9 Arbeitsschutz ----------
C.push(H1('9. Arbeitsschutz'));
C.push(BL([b('Staub: '), t('FFP2/FFP3-Masken beim Schleifen, Stemmen, Abbruch; Bausauger mit Staubklasse M statt Besen.')]));
C.push(BL([b('Augen, Ohren, Hände, Füße: '), t('Schutzbrille, Gehörschutz bei Bohrhammer und Flex, Schnittschutzhandschuhe, Sicherheitsschuhe S3.')]));
C.push(BL([b('Höhe: '), t('Stabile Stehleiter, kein Arbeiten auf wackeligen Stühlen, für Fassade/Dach Gerüst oder Fachbetrieb.')]));
C.push(BL([b('Strom: '), t('Vor dem Bohren Leitungen orten, Stromkreise abschalten und spannungsfrei prüfen.')]));
C.push(BL([b('Notfall: '), t('Erste-Hilfe-Kasten, Feuerlöscher, Handy geladen, nie allein schwere Abbrucharbeiten.')]));

// ---------- 10 Werkzeuge & Material ----------
C.push(H1('10. Zusammenfassung: Werkzeuge und Materialien'));
C.push(P('Die vollständige Liste mit Preisspannen, Kaufen/Leihen-Empfehlung und Abhakfeld steht in der Excel-Datei „Werkzeug_und_Materialliste.xlsx“. Hier die Kurzfassung.'));
C.push(H2('Werkzeuge – Grundausstattung (kaufen)'));
C.push(table([2600, 7000], [
  ['Bereich', 'Werkzeuge'],
  ['Akku-System (eine Marke, 18 V)', 'Akku-Bohrschrauber, Schlagschrauber, Bohrhammer SDS-plus, Winkelschleifer 125 mm, Stichsäge, Multitool (oszillierend), Handkreissäge, LED-Baustrahler, 3–4 Akkus + Schnellladegerät'],
  ['Netzgeräte', 'Bausauger Staubklasse M, Rührwerk für Kleber/Spachtel, Exzenterschleifer, Kappsäge (Böden, Leisten), Heißluftföhn'],
  ['Messen & Prüfen', 'Kreuzlinienlaser, Wasserwaage 60/120 cm, Richtlatte 2 m, Zollstock, Maßband, Winkel, Leitungsortungsgerät, zweipoliger Spannungsprüfer, Feuchtemessgerät'],
  ['Handwerkzeug', 'Hammer, Fäustel, Meißel-Set, Kuhfuß/Brechstange, Cuttermesser, VDE-Schraubendreher, Zangen-Set, Spachtel-Set, Kartuschenpresse, Kelle, Gummihammer, Zahnspachtel'],
  ['Maler & Boden', 'Farbrollen + Teleskopstange, Pinsel, Abdeckvlies, Malerkrepp, Zugeisen, Schlagklotz, Abstandskeile, Cutter für Vinyl'],
  ['Transport & Ordnung', 'Schubkarre, Eimer, Mörtelkübel, Besen, Schaufel, Leiter, Arbeitsböcke, Verlängerungskabel/Kabeltrommel, Werkzeugkiste'],
  ['Schutz', 'FFP3-Masken, Schutzbrille, Gehörschutz, Handschuhe, Sicherheitsschuhe S3, Knieschoner, Erste-Hilfe-Kasten, Feuerlöscher'],
], { boldFirst: true }));
C.push(P([t('Startausstattung inklusive Schutzausrüstung laut Liste ca. 2.200–5.100 € – je nach Marke und Qualität. Lohnt sich ab dem ersten Projekt, weil die Eigenleistung mehr spart.', { italics: true })], { spacing: { before: 80 } }));
C.push(H2('Werkzeuge – leihen (Baumaschinenverleih)'));
C.push(BL('Abbruchhammer, Wandschlitzfräse, Fliesen-/Bodenstripper, Nass-Fliesenschneider, Bautrockner/Luftentfeuchter, Estrich-/Bodenschleifer, Airless-Farbspritzgerät, Dampf-Tapetenablöser, Hochdruckreiniger (oder kaufen), Gerüst, Anhänger, Container.'));
C.push(H2('Materialien – dein Flip-Standard'));
C.push(table([2400, 7200], [
  ['Gewerk', 'Standard-Material (Beispiel)'],
  ['Maler', 'Tiefengrund, Spachtelmasse, Malervlies (bei schlechten Wänden), Dispersionsfarbe weiß (Nassabrieb Klasse 2), Acryl'],
  ['Boden', 'Rigid-/Klick-Vinyl helle Eiche (Nutzungsklasse 33) mit integrierter Trittschalldämmung, weiße Sockelleisten, Übergangsprofile'],
  ['Bad', 'Feinsteinzeug 60 × 60 cm matt hellgrau/beige, Flexkleber C2, Abdichtungs-Set (Flüssigabdichtung, Dichtband, Manschetten), Sanitärsilikon, Vorwandelemente, weiße Keramik, Mittelklasse-Armaturen, Duschrinne + Glaswand, LED-Spiegel, Handtuchheizkörper'],
  ['Trockenbau & Dämmung', 'Gipskartonplatten (im Bad imprägniert), CW/UW-Profile, Mineralwolle oder Holzfaser (WLS 035), Dampfbremse + Klebeband, Fugenspachtel, Kantenschutz'],
  ['Elektro (über Elektriker)', 'Schalterprogramm reinweiß, LED-Spots, Rauchmelder, Leerrohre'],
  ['Türen & Fenster', 'Innentüren weiß CPL mit Zargen, einheitliche Drückergarnituren, Fensterbänke'],
  ['Kleinmaterial', 'Schrauben-/Dübel-Sortiment, Bauschaum, Montagekleber, Abdeckfolie, Malerkrepp, Big Bags, Müllsäcke, Schleifpapier, Trennscheiben, Bohrer-Set'],
], { boldFirst: true }));

const doc = new Document({
  creator: 'Sanierungs-Handbuch', title: 'Fix & Flip – Sanierungs-Handbuch', styles, numbering, features: { updateFields: true },
  sections: [{
    properties: { page: { size: { width: 11906, height: 16838 }, margin: { top: 1134, bottom: 1134, left: 1134, right: 1134 } } },
    footers: { default: new Footer({ children: [new Paragraph({ alignment: AlignmentType.RIGHT, children: [t('Fix & Flip – Sanierungs-Handbuch · Seite ', { size: 16, color: '888888' }), new TextRun({ children: [PageNumber.CURRENT], font: FONT, size: 16, color: '888888' })] })] }) },
    children: C,
  }],
});
Packer.toBuffer(doc).then(buf => { fs.writeFileSync(OUT, buf); console.log('geschrieben:', OUT); });
