// Erstellt den Fix-&-Flip-Plan als Word-Dokument.
// Aufruf: NODE_PATH=$(npm root -g) node fix_flip_plan.js deals.json AUSGABE.docx
const fs = require('fs');
const {
  Document, Packer, Paragraph, TextRun, HeadingLevel, AlignmentType, LevelFormat, Table, TableRow, TableCell,
  WidthType, ShadingType, BorderStyle, PageBreak, Footer, PageNumber, TableOfContents,
} = require('docx');

const deals = JSON.parse(fs.readFileSync(process.argv[2], 'utf8'));
const OUT = process.argv[3];
const D = (name) => deals[name];
const eur = (n) => (n < 0 ? '−' : '') + Math.round(Math.abs(n)).toLocaleString('de-DE') + ' €';
const pct = (n) => (n < 0 ? '−' : '') + Math.abs(n * 100).toLocaleString('de-DE', { minimumFractionDigits: 1, maximumFractionDigits: 1 }) + ' %';
const k = (n) => Math.round(n / 1000).toLocaleString('de-DE') + ' T€';

const FONT = 'Calibri';
const ACC = '9E3B24';
const W = 9638; // Inhaltsbreite A4 mit 2 cm Rändern

const t = (text, o = {}) => new TextRun({ text, font: FONT, size: 21, ...o });
const b = (text, o = {}) => t(text, { bold: true, ...o });
const P = (runs, o = {}) => new Paragraph({ spacing: { after: 120, line: 276 }, ...o, children: (Array.isArray(runs) ? runs : [runs]).map(r => typeof r === 'string' ? t(r) : r) });
const H1 = (text) => new Paragraph({ heading: HeadingLevel.HEADING_1, spacing: { before: 360, after: 160 }, children: [new TextRun({ text, font: FONT, size: 32, bold: true, color: '1C2430' })] });
const H2 = (text) => new Paragraph({ heading: HeadingLevel.HEADING_2, spacing: { before: 240, after: 100 }, children: [new TextRun({ text, font: FONT, size: 25, bold: true, color: ACC })] });
const BL = (runs, lvl = 0) => new Paragraph({ numbering: { reference: 'bul', level: lvl }, spacing: { after: 60, line: 264 }, children: (Array.isArray(runs) ? runs : [runs]).map(r => typeof r === 'string' ? t(r) : r) });
const NL = (runs, ref = 'num') => new Paragraph({ numbering: { reference: ref, level: 0 }, spacing: { after: 60, line: 264 }, children: (Array.isArray(runs) ? runs : [runs]).map(r => typeof r === 'string' ? t(r) : r) });

const border = { style: BorderStyle.SINGLE, size: 4, color: 'D5D2CC' };
const borders = { top: border, bottom: border, left: border, right: border };
function table(widths, rows, opts = {}) {
  const sum = widths.reduce((a, c) => a + c, 0);
  const scale = W / sum;
  const ws = widths.map(w => Math.round(w * scale));
  ws[ws.length - 1] += W - ws.reduce((a, c) => a + c, 0);
  return new Table({
    width: { size: W, type: WidthType.DXA },
    columnWidths: ws,
    rows: rows.map((r, ri) => new TableRow({
      tableHeader: ri === 0,
      children: r.map((cell, ci) => {
        const head = ri === 0;
        const fill = head ? ACC : (opts.zebra && ri % 2 === 0 ? 'F6F4F1' : (opts.fills && opts.fills[ri] && opts.fills[ri][ci]) || undefined);
        const txt = Array.isArray(cell) ? cell : [cell];
        return new TableCell({
          borders, width: { size: ws[ci], type: WidthType.DXA },
          shading: fill ? { fill, type: ShadingType.CLEAR, color: 'auto' } : undefined,
          margins: { top: 60, bottom: 60, left: 90, right: 90 },
          children: [new Paragraph({
            alignment: (opts.right && opts.right.includes(ci) && !head) ? AlignmentType.RIGHT : AlignmentType.LEFT,
            children: txt.map(x => typeof x === 'string' ? t(x, { size: 18, bold: head, color: head ? 'FFFFFF' : undefined }) : x),
          })],
        });
      }),
    })),
  });
}
const box = (title, lines, fill = 'F5E6E0') => new Table({
  width: { size: W, type: WidthType.DXA }, columnWidths: [W],
  rows: [new TableRow({ children: [new TableCell({
    borders: { top: { style: BorderStyle.SINGLE, size: 18, color: ACC }, bottom: border, left: border, right: border },
    shading: { fill, type: ShadingType.CLEAR, color: 'auto' }, margins: { top: 120, bottom: 120, left: 160, right: 160 },
    children: [new Paragraph({ spacing: { after: 80 }, children: [b(title, { size: 22 })] }),
      ...lines.map(l => new Paragraph({ spacing: { after: 60 }, children: (Array.isArray(l) ? l : [l]).map(r => typeof r === 'string' ? t(r) : r) }))],
  })] })],
});
const gap = () => new Paragraph({ spacing: { after: 80 }, children: [] });

// ---- Kennzahlen aus dem Rechner ----
const bun = D('Bungalow Damme-West'), dhh = D('Doppelhaushälfte Damme'), loh = D('Einfamilienhaus Lohne'),
  lem = D('Haus Lemförde'), etw = D('Wohnung Damme-Ort'), live = D('Live-in-Flip: DHH Damme');

const C = [];
// ---------- Titel ----------
C.push(new Paragraph({ spacing: { before: 1800, after: 200 }, children: [new TextRun({ text: 'FIX & FLIP', font: FONT, size: 26, bold: true, color: ACC, characterSpacing: 60 })] }));
C.push(new Paragraph({ spacing: { after: 200 }, children: [new TextRun({ text: 'Strategieplan für Damme und Umgebung', font: FONT, size: 52, bold: true, color: '1C2430' })] }));
C.push(P([t('Ältere Häuser und Wohnungen kaufen, gezielt sanieren und mit Gewinn verkaufen – mit allen Zahlen, Regeln und Risiken, die dafür entscheidend sind, und den ersten Deals aus der Region.', { size: 24, color: '555555' })]));
C.push(new Paragraph({ spacing: { before: 600, after: 60 }, children: [t('Stand: November 2026 · Begleitdatei: ', { color: '555555' }), b('Fix_Flip_Rechner_Damme.xlsx', { color: '555555' })] }));
C.push(P([t('Keine Steuer- oder Rechtsberatung. Alle Marktwerte sind Schätzungen aus Online-Inseraten und Ratgebern; vor jeder Entscheidung mit Steuerberater, Notar, Bank und Bausachverständigem prüfen.', { size: 18, italics: true, color: '777777' })]));
C.push(new Paragraph({ children: [new PageBreak()] }));

// ---------- Inhalt ----------
C.push(H1('Inhalt'));
C.push(new TableOfContents('Inhalt', { hyperlink: true, headingStyleRange: '1-2' }));
C.push(P([t('Hinweis: Das Inhaltsverzeichnis füllt Word beim Öffnen (ggf. Rechtsklick → „Feld aktualisieren“).', { size: 18, italics: true, color: '777777' })]));
C.push(new Paragraph({ children: [new PageBreak()] }));

// ---------- 1 Kurzfassung ----------
C.push(H1('1. Das Wichtigste auf einer Seite'));
C.push(box('Ehrliches Fazit', [
  [b('Fix & Flip kann in Damme funktionieren – aber nicht zu den Preisen, die heute online stehen. '), t('Von fünf aktuellen Angeboten bringt zum Angebotspreis nur eines (Lohne) einen Gewinn, und auch das nur knapp.')],
  [b('Der Grund: '), t('Zwischen unsaniert (ca. 2.000–2.100 €/m²) und gut saniert (ca. 2.600–2.800 €/m²) liegen in Damme nur etwa 500–700 €/m². Kauf- und Verkaufsnebenkosten, Zinsen und Sanierung fressen diese Spanne fast vollständig auf.')],
  [b('Der Gewinn entsteht im Einkauf. '), t('Ein Flip lohnt sich hier nur, wenn du 20–30 % unter dem Angebotspreis kaufst – also abseits der Portale (Erbfälle, Zwangsversteigerungen, Privatverkäufe, deine Anschreiben) – oder einen großen Werthebel findest (Energieklasse H → C, Aufteilung in zwei Wohnungen, Grundstücksteilung).')],
  [b('Kapitalbedarf: '), t(`Bei klassischer Flip-Finanzierung (70 % Bank) brauchst du pro Haus etwa 100.000–140.000 € Eigenkapital. Das ist für den Anfang sehr viel.`)],
]));
C.push(gap());
C.push(H2('Meine Empfehlung in drei Schritten'));
C.push(NL([b('Start als Live-in-Flip: '), t('Ein renovierungsbedürftiges Haus kaufen, selbst einziehen, Schritt für Schritt (mit viel Eigenleistung) sanieren und nach gut 2 Jahren verkaufen. Vorteile: Gewinn steuerfrei, bis zu 90 % Finanzierung zu Eigenheim-Zinsen, höhere Förderung, gesparte Miete.')]));
C.push(NL([b('Parallel Deals jagen, die echte Flips sind: '), t('Ziel 1–2 Objekte pro Jahr, die du 25 % unter Marktwert bekommst. Jedes Angebot nur bis zum „Max. Kaufpreis“ aus dem Rechner.')]));
C.push(NL([b('Immer mit Plan B: '), t('Jedes Flip-Objekt muss sich auch als Vermietung tragen („Flip or Hold“). Findet sich kein Käufer, vermietest du – das passt zu deiner Kapitalanlage-Strategie.')]));
C.push(H2('Die Deals im Überblick'));
const dealRow = (name, d, note) => [name, eur(d.angebot), eur(d.gewinn), pct(d.marge), eur(d.maxK), note];
C.push(table([2300, 1300, 1400, 1000, 1400, 2200], [
  ['Objekt', 'Angebot', 'Gewinn vor Steuern', 'Marge', 'Max. Kaufpreis', 'Einschätzung'],
  dealRow('Einfamilienhaus Lohne (1946, 224 m²)', loh, 'Bester Kandidat – besichtigen, max. ca. 175.000 € bieten'),
  dealRow('Bungalow Damme-West (139,5 m²)', bun, 'Nur bei ca. 180.000 € interessant'),
  dealRow('Doppelhaushälfte Damme (1984)', dhh, 'Als Flip zu teuer; als Live-in-Flip prüfen'),
  dealRow('Wohnung Damme-Ort (1993)', etw, 'Kein Sanierungsfall – Gegenbeispiel'),
  dealRow('Haus Lemförde (Schimmel)', lem, 'Nur über Grundstücksteilung denkbar'),
], { right: [1, 2, 3, 4] }));
C.push(P([t('Annahmen: Sanierung je nach Zustand 300–900 €/m² plus 15–20 % Puffer, 70 % Bankfinanzierung zu 4,5 %, Verkauf über Makler (3,57 %), Zielmarge 15 %. Details in Kapitel 9 und im Excel-Rechner.', { size: 18, italics: true, color: '666666' })], { spacing: { before: 80 } }));
C.push(new Paragraph({ children: [new PageBreak()] }));

// ---------- 2 Grundprinzip ----------
C.push(H1('2. So funktioniert Fix & Flip'));
C.push(P('Beim Fix & Flip kaufst du eine Immobilie unter Marktwert, weil sie einen sichtbaren Mangel hat (alt, unsaniert, Erbfall, Zeitdruck des Verkäufers). Durch eine gezielte Sanierung hebst du sie auf das Niveau, das Käufer in der Gegend erwarten, und verkaufst sie zum Marktpreis für sanierte Objekte. Dein Gewinn ist die Differenz – abzüglich aller Kosten.'));
C.push(H2('Die Formel'));
C.push(box('Gewinn = Verkaufspreis − (Kaufpreis + Kaufnebenkosten + Sanierung + Haltekosten + Verkaufskosten + Steuern)', [
  [b('Verkaufspreis (ARV, „after repair value“): '), t('Was vergleichbare, sanierte Häuser in derselben Lage wirklich bringen – nicht der Wunschpreis.')],
  [b('Kaufnebenkosten: '), t('Grunderwerbsteuer 5 % (Niedersachsen; NRW 6,5 %), Notar und Grundbuch ca. 2 %, ggf. Makler 3,57 % → bis zu 10,6 % des Kaufpreises.')],
  [b('Haltekosten: '), t('Zinsen, Grundsteuer, Versicherung, Strom/Heizung im Leerstand, bei Wohnungen Hausgeld – jeden Monat, bis der Käufer bezahlt hat.')],
  [b('Verkaufskosten: '), t('Makler (Verkäuferanteil bei Häusern an Privatleute meist 3,57 %), Energieausweis, Fotos, Anzeigen.')],
], 'F6F4F1'));
C.push(H2('Die sieben Phasen eines Flips'));
C.push(table([1500, 3700, 2600, 1800], [
  ['Phase', 'Was passiert', 'Ergebnis', 'Dauer'],
  ['1. Vorbereitung', 'Markt verstehen, Vergleichspreise sammeln, Team aufbauen, Finanzierungsrahmen klären', 'Du weißt, was saniert in Damme kostet', '1–3 Monate (einmalig)'],
  ['2. Akquise', 'Objekte finden: Portale, Zwangsversteigerungen, Erbfälle, Anschreiben, Netzwerk', 'Laufend neue Objekte', 'laufend'],
  ['3. Prüfung', 'Besichtigung mit Bausachverständigem, Sanierungskosten schätzen, Rechner ausfüllen', 'Höchstgebot', '1–2 Wochen je Objekt'],
  ['4. Kauf', 'Verhandeln, Finanzierung, Notar, Kaufpreiszahlung, Übergabe', 'Schlüssel in der Hand', '1–2 Monate'],
  ['5. Sanierung', 'Planung, Angebote, Bauzeitenplan, Baustelle steuern, Kosten kontrollieren', 'Verkaufsfertiges Objekt', '2–5 Monate'],
  ['6. Vermarktung', 'Energieausweis, Home Staging, Profi-Fotos, Preis festlegen, Besichtigungen', 'Käufer mit Finanzierungsbestätigung', '1–2 Monate'],
  ['7. Verkauf', 'Notar, Zahlung, Übergabe, Steuer, Auswertung („Was lerne ich daraus?“)', 'Gewinn auf dem Konto', '1–2 Monate'],
], { zebra: true }));
C.push(P([t('Realistische Gesamtdauer pro Objekt: 6–12 Monate. Jeder zusätzliche Monat kostet Zinsen und Nebenkosten – Zeit ist beim Flip Geld.', { italics: true })], { spacing: { before: 80 } }));

// ---------- 3 Zahlen ----------
C.push(H1('3. Die Zahlen: Kalkulation und Faustregeln'));
C.push(H2('Die fünf Faustregeln'));
C.push(NL([b('Höchstgebot ≈ 70 % des Verkaufspreises minus Sanierungskosten. '), t('Beispiel: saniert 300.000 €, Sanierung 60.000 € → maximal ca. 150.000 € bieten. Der Excel-Rechner berechnet das genauer („Max. Kaufpreis für die Zielmarge“).')], 'num2'));
C.push(NL([b('Zielmarge mindestens 15 % des Verkaufspreises vor Steuern. '), t('Darunter reicht ein einziger Kostenausreißer, und der Gewinn ist weg.')], 'num2'));
C.push(NL([b('15–20 % Puffer auf die Sanierung, 2 Monate Puffer auf die Zeit. '), t('Bei alten Häusern taucht fast immer etwas auf (Leitungen, Feuchte, Asbest).')], 'num2'));
C.push(NL([b('Sanieren auf Nachbarschafts-Niveau, nicht auf Luxus-Niveau. '), t('Ein Marmorbad in Damme-West bringt nicht mehr Verkaufspreis, nur weniger Gewinn.')], 'num2'));
C.push(NL([b('Erst Verkaufspreis, dann Kaufpreis. '), t('Du rechnest immer rückwärts vom realistischen Verkaufspreis. Nie vom Angebotspreis aus.')], 'num2'));
C.push(H2('Was die Spanne in Damme bedeutet'));
C.push(table([3800, 2900, 2900], [
  ['Kennzahl (Stand Okt./Nov. 2026)', 'Wert', 'Quelle / Hinweis'],
  ['Häuser Damme, Durchschnitt', 'ca. 2.420 €/m²', 'immowelt, Okt. 2026'],
  ['Unsanierte Doppelhaushälfte Damme', 'ca. 2.090 €/m²', 'Inserat 235.000 € / 112,6 m²'],
  ['Neubau-Niveau, Energieklasse A+', 'ca. 3.480 €/m²', 'Inserat Damme-Süd 417.500 € / 120 m²'],
  ['Gut saniertes älteres Haus (Annahme)', '2.600–2.800 €/m²', 'vorsichtig; mit Gutachterausschuss prüfen'],
  ['Häuser Lohne, Durchschnitt', 'ca. 2.360–2.460 €/m²', 'immowelt Juli 2026, Engel & Völkers'],
  ['Wohnungen Damme, Durchschnitt', 'ca. 2.310 €/m²', 'Portal-Durchschnitt Nov. 2026'],
], { zebra: true }));
C.push(P([t('Rechenbeispiel Doppelhaushälfte Damme (113 m²): saniert ca. 293.800 € – das Inserat verlangt 235.000 €. Die Spanne von knapp 59.000 € reicht nicht einmal für die Sanierung (ca. 58.500 € inkl. Puffer), geschweige denn für Nebenkosten (ca. 24.800 €), Zinsen und Verkauf. Ergebnis: ', {}), b(eur(dhh.gewinn) + ' Verlust.'), t(` Erst bei einem Kaufpreis von ca. ${eur(dhh.maxK)} würde die Zielmarge erreicht.`)], { spacing: { before: 100 } }));
C.push(H2('Wo der Hebel wirklich liegt'));
C.push(BL([b('Energieklasse: '), t('Laut immowelt (2026) werden Häuser der Klasse G ca. 10 % und Klasse H ca. 17 % günstiger angeboten als Klasse D; das IW misst gegenüber A/A+ sogar 20–30 % Abschlag für G/H. Ein Sprung von H/G auf D/C (Dämmung oberste Geschossdecke, Fenster, Heizung) ist oft der stärkste Wertsprung pro investiertem Euro.')]));
C.push(BL([b('Fläche besser nutzen: '), t('Ein großes Haus in zwei Wohnungen aufteilen (Baugenehmigung!) oder einen Dachboden ausbauen. Zwei kleinere Einheiten verkaufen oder vermieten sich in der Region oft besser als ein 224-m²-Haus.')]));
C.push(BL([b('Grundstück: '), t('Große Grundstücke (z. B. 3.136 m² in Lemförde) können teilbar sein. Ein abgetrennter Bauplatz kann mehr wert sein als das alte Haus. Vorher beim Bauamt klären (Bebauungsplan, § 34 BauGB).')]));
C.push(BL([b('Optik & erster Eindruck: '), t('Helle Böden, frische Wände, neues Bad, gepflegter Garten und gute Fotos kosten wenig und entscheiden über den Preis.')]));

// ---------- 4 Steuern & Recht ----------
C.push(H1('4. Steuern und Recht'));
C.push(H2('Steuern beim Verkauf'));
C.push(table([2600, 3900, 3100], [
  ['Situation', 'Was gilt', 'Folge für dich'],
  ['Verkauf innerhalb von 10 Jahren (Privatperson)', 'Privates Veräußerungsgeschäft nach § 23 EStG: Gewinn wird mit deinem persönlichen Steuersatz versteuert.', 'Bei 30–42 % Steuersatz bleiben vom Gewinn nur 58–70 % übrig.'],
  ['Du hast selbst darin gewohnt', 'Steuerfrei, wenn du das Objekt im Verkaufsjahr und in den beiden Kalenderjahren davor selbst bewohnt hast (oder durchgehend seit dem Kauf).', 'Live-in-Flip: z. B. Einzug Dezember 2026, Verkauf frühestens Januar 2028. Genau mit dem Steuerberater planen.'],
  ['Mehr als 3 Objekte in ca. 5 Jahren gekauft und verkauft', '„Drei-Objekt-Grenze“ aus der Rechtsprechung: Dann gilt man als gewerblicher Grundstückshändler. Das Finanzamt kann das im Einzelfall auch früher annehmen.', 'Alle Gewinne steuerpflichtig (auch nach 10 Jahren) + Gewerbesteuer. Ab einer Serie von Flips ist eine GmbH meist sinnvoller – Steuerberater fragen.'],
  ['Umsatzsteuer', 'Der Verkauf von Grundstücken ist in der Regel umsatzsteuerfrei.', 'Die Mehrwertsteuer auf Handwerkerrechnungen kannst du als Privatperson nicht zurückholen – rechne immer brutto.'],
], { zebra: true }));
C.push(H2('Recht und Pflichten'));
C.push(BL([b('Gewährleistung beim Verkauf: '), t('Bei gebrauchten Immobilien wird die Haftung für Mängel im Notarvertrag üblicherweise ausgeschlossen. Das gilt nicht, wenn du einen Mangel kennst und verschweigst (Arglist). Bei sehr umfassenden Sanierungen kann eine Haftung ähnlich wie bei einem Neubau drohen. Deshalb: Mängel offenlegen, alle Rechnungen und Fotos dokumentieren, Vertrag vom Notar/Anwalt gestalten lassen.')]));
C.push(BL([b('Baugenehmigung: '), t('Nötig bei tragenden Wänden, Anbauten, Dachausbau oder Aufteilung in zwei Wohnungen. Für den Verkauf einzelner Wohnungen braucht es zusätzlich eine Abgeschlossenheitsbescheinigung und eine Teilungserklärung beim Notar.')]));
C.push(BL([b('Gebäudeenergierecht: '), t('Das Gebäudemodernisierungsgesetz hat im Sommer 2026 das GEG abgelöst; die 65-%-Regel für neue Heizungen ist entfallen. Die genauen Nachrüstpflichten nach einem Eigentümerwechsel mit dem Energieberater klären. Ein Energieausweis ist im Inserat Pflicht.')]));
C.push(BL([b('Makler beim Verkauf: '), t('Verkaufst du ein Haus oder eine Wohnung an Privatleute und beauftragst einen Makler, zahlst du mindestens die Hälfte der Provision (Halbteilungsgrundsatz). Ein Privatverkauf spart diese Kosten, kostet aber Zeit.')]));
C.push(BL([b('Gewerbe: '), t('Wer regelmäßig flippt, muss ein Gewerbe anmelden. Für Bauträgertätigkeit (Verkauf vor Fertigstellung) gelten weitere Pflichten (§ 34c GewO) – das ist beim normalen Flip nicht der Fall.')]));
C.push(BL([b('Versicherung: '), t('Leerstand und Bauphase der Gebäudeversicherung melden, Bauherrenhaftpflicht abschließen.')]));

// ---------- 5 Finanzierung ----------
C.push(H1('5. Finanzierung und Förderung'));
C.push(H2('Wie Banken einen Flip sehen'));
C.push(BL('Banken finanzieren lieber Eigenheime und Vermietung als kurzfristigen Handel. Erwarte für einen reinen Flip etwa 60–80 % Finanzierung von Kauf und Sanierung – die Nebenkosten und meist einen Teil der Sanierung zahlst du aus Eigenkapital.'));
C.push(BL('Wähle ein Darlehen, das du ohne hohe Vorfälligkeitsentschädigung ablösen kannst: variable Zinsen, kurze Zinsbindung oder ein vereinbartes Ablöserecht bei Verkauf. Eine 10-jährige Zinsbindung passt nicht zu einem 8-Monats-Projekt.'));
C.push(BL('Zeige der Bank immer Plan B: Wie trägt sich das Objekt, wenn du es vermietest? Das macht die Finanzierung deutlich leichter.'));
C.push(BL('Vereinbare eine Auszahlung in Tranchen für die Sanierung (nach Baufortschritt) – so zahlst du nur Zinsen auf das, was du wirklich brauchst.'));
C.push(H2('Was dich ein Projekt an Eigenkapital kostet'));
C.push(table([3400, 2100, 2100, 2000], [
  ['Objekt', 'Investition gesamt', 'Eigenkapital-bedarf', 'Bankanteil'],
  ['Bungalow Damme-West', eur(bun.inv), eur(bun.ek), '70 %'],
  ['Doppelhaushälfte Damme', eur(dhh.inv), eur(dhh.ek), '70 %'],
  ['Einfamilienhaus Lohne', eur(loh.inv), eur(loh.ek), '70 %'],
  ['Live-in-Flip Doppelhaushälfte', eur(live.inv), eur(live.ek), '85 %'],
], { zebra: true, right: [1, 2, 3] }));
C.push(P([t('Das Eigenkapital enthält Nebenkosten, den nicht finanzierten Teil der Sanierung sowie Zinsen und laufende Kosten während der Haltedauer.', { size: 18, italics: true, color: '666666' })], { spacing: { before: 80 } }));
C.push(H2('Förderung (Stand: BEG-Reform 21.07.2026)'));
C.push(BL([b('Heizungstausch: '), t('Für vermietete bzw. nicht selbst genutzte Objekte gibt es nach aktuellen Angaben nur die Grundförderung von 30 %. Der Klimageschwindigkeits- und der Einkommensbonus gelten nur für Selbstnutzer – ein weiterer Vorteil des Live-in-Flips. Ab 2027 sind Änderungen angekündigt; vor jeder Maßnahme bei der KfW prüfen.')]));
C.push(BL([b('Energieberater einbinden: '), t('Ein Energieberater klärt, welche Programme aktuell gelten (Gebäudehülle, Effizienzhaus-Kredite), und stellt nach der Sanierung den neuen Energieausweis aus. Förderanträge müssen meist vor Beginn der Maßnahme gestellt werden.')]));

// ---------- 6 Sanierung ----------
C.push(H1('6. Was sanieren – und was nicht'));
C.push(P('Die Grundregel: Investiere dort, wo Käufer zuerst hinschauen und wo die Energieklasse springt. Alles andere ist Liebhaberei.'));
C.push(table([3000, 1900, 1700, 3000], [
  ['Maßnahme', 'Kosten (Richtwert 2026)', 'Wirkung', 'Empfehlung'],
  ['Streichen, Tapeten, Böden', '15–30 €/m² + 40–80 €/m²', 'hoch', 'Immer – auch in Eigenleistung'],
  ['Bad komplett', '12.000–25.000 €', 'sehr hoch', 'Bei altem Bad immer'],
  ['Elektrik', '90–140 €/m²', 'Sicherheit', 'Bei alter Anlage ohne FI Pflicht'],
  ['Dämmung oberste Geschossdecke', '30–60 €/m²', 'Energieklasse', 'Fast immer – günstigster Klassensprung'],
  ['Fenster', '12.000–22.500 € (ca. 15 Stück)', 'hoch', 'Bei Einfach-/alter Verglasung'],
  ['Heizung/Wärmepumpe', '25.000–40.000 € (−30 % Förderung)', 'Energieklasse', 'Je nach Objekt'],
  ['Dach neu', '150–300 €/m² Dachfläche', 'notwendig', 'Nur wenn nötig – vor Kauf prüfen'],
  ['Fassadendämmung', '150–250 €/m² Fassade', 'mittel', 'Selten beim Flip'],
  ['Küche', '5.000–12.000 €', 'gering', 'Meist weglassen'],
  ['Garten, Eingang, Home Staging, Fotos', '3.000–11.000 €', 'hoch', 'Immer'],
], { zebra: true }));
C.push(P([t('Ausführliche Liste mit Hinweisen im Excel-Rechner, Blatt „Sanierungskosten“. Die Spannen stammen aus Ratgebern 2026 und weichen stark voneinander ab – für jedes Objekt mindestens zwei Angebote je Gewerk aus der Region einholen.', { size: 18, italics: true, color: '666666' })], { spacing: { before: 80 } }));
C.push(H2('Sanierungsstufen pro m² Wohnfläche'));
C.push(BL([b('Auffrischung 200–400 €/m²: '), t('Optik, Böden, Bad, kleine Technik (z. B. Wohnung von 1993).')]));
C.push(BL([b('Mittlere Sanierung 400–700 €/m²: '), t('zusätzlich Elektrik, Fenster, Teile der Heizung (z. B. Haus von 1984).')]));
C.push(BL([b('Komplettsanierung 800–1.400 €/m²: '), t('alles bis auf den Rohbau (z. B. Haus von 1946 mit Klasse H oder Wasserschaden). Für Einsteiger zu riskant.')]));
C.push(H2('Eigenleistung'));
C.push(P('Als junger, handwerklich interessierter Käufer kannst du viel selbst machen: Abbruch, Entrümpelung, Malerarbeiten, Böden, Garten. Das spart schnell 10.000–20.000 € pro Objekt. Elektrik, Gas, Heizung und Statik gehören immer in die Hände von Fachbetrieben – auch wegen Haftung und Versicherung.'));

// ---------- 7 Markt ----------
C.push(H1('7. Der Markt in Damme und Umgebung'));
C.push(BL([b('Käufer: '), t('Vor allem Familien und Paare, die in der Region arbeiten (Industrie, Landmaschinen, Agrar- und Ernährungswirtschaft). Sie suchen Einfamilienhäuser, Doppelhaushälften und Bungalows mit 110–140 m², 4–5 Zimmern, Garten und vernünftiger Energieklasse.')]));
C.push(BL([b('Ältere Käufer: '), t('Ebenerdige Bungalows sind gefragt bei Menschen, die sich verkleinern wollen. Ein sanierter Bungalow spricht zwei Käufergruppen an.')]));
C.push(BL([b('Preisniveau: '), t('Damme liegt mit ca. 2.420 €/m² für Häuser im Mittelfeld des Landkreises Vechta. Lohne ist ähnlich, Lemförde und Stemwede (NRW) günstiger, aber dort sind auch die Verkaufspreise niedriger.')]));
C.push(BL([b('Konkurrenz: '), t('Familien, die selbst sanieren wollen, zahlen für unsanierte Häuser oft mehr, als ein Flip verträgt. Gute Flip-Objekte findest du deshalb selten über Portale, sondern dort, wo Verkäufer schnell und unkompliziert verkaufen wollen.')]));
C.push(BL([b('Wichtigste Hausaufgabe: '), t('Kaufe den Grundstücksmarktbericht des Gutachterausschusses (für den Landkreis Vechta: Gutachterausschuss Oldenburg-Cloppenburg) und sammle 20 Vergleichsobjekte sanierter Häuser mit Preis, Baujahr, Fläche und Energieklasse. Das ist die Grundlage für jeden Verkaufspreis im Rechner.')]));

// ---------- 8 Deals finden ----------
C.push(H1('8. Wo du Flip-Objekte findest'));
C.push(table([2600, 4400, 2600], [
  ['Kanal', 'So gehst du vor', 'Rabatt-Chance'],
  ['Zwangsversteigerungen', 'Amtsgerichte Vechta, Diepholz, Osnabrück, Rahden (NRW) auf zvg-portal.de wöchentlich prüfen. Gutachten lesen, Objekt von außen ansehen, Finanzierung vorher klären (10 % Sicherheitsleistung im Termin).', 'hoch, aber ohne Innenbesichtigung'],
  ['Erbengemeinschaften', 'Erben wollen oft schnell und ohne Aufwand verkaufen, besonders wenn sie weit weg wohnen. Am besten über Bekannte, Vereine und eine eigene Anzeige „Suche Haus, auch renovierungsbedürftig, schnelle und unkomplizierte Abwicklung“ – immer respektvoll und ohne Druck.', 'hoch'],
  ['Deine Anschreiben', 'Dieselbe Methode wie bei der Rückmietung – mit einer zweiten Variante: Kauf mit Auszug, schnelle Abwicklung, keine Besichtigungen für Fremde.', 'mittel bis hoch'],
  ['Lange inserierte Objekte', 'Häuser, die seit über 3 Monaten online sind, mit Wasserschaden, Schimmel, schlechter Energieklasse. Dort ist Verhandlungsspielraum.', 'mittel'],
  ['Banken & Makler vor Ort', 'Immobilienabteilungen der Volksbanken und Sparkassen, lokale Makler (z. B. Völker, Haskamp): sag ihnen genau, was du suchst und dass du schnell entscheidest.', 'mittel'],
  ['Kleinanzeigen & Suchaufträge', 'Suchaufträge mit „renovierungsbedürftig“ und Preis bis 220.000 € in Damme + 20 km.', 'gering bis mittel'],
], { zebra: true }));

// ---------- 9 Deals ----------
C.push(H1('9. Aktuelle Deals aus der Region'));
C.push(P([t('Alle Zahlen stammen aus dem Excel-Rechner (Annahmen dort änderbar). Die Inserate sind Stand November 2026 und können bereits verkauft oder geändert sein. ', {}), b('Vor jedem Gebot: Exposé anfordern, besichtigen, Bausachverständigen mitnehmen.')]));
const dealBlock = (title, d, facts, verdict, next) => {
  C.push(H2(title));
  C.push(P([t(facts, { color: '444444' })]));
  C.push(table([2400, 2400, 2400, 2400], [
    ['Verkaufspreis saniert', 'Gesamtkosten', 'Gewinn vor Steuern', 'Max. Kaufpreis'],
    [eur(d.arv), eur(d.gesamt), eur(d.gewinn), `${eur(d.maxK)} (${pct(-d.abschlag)})`],
  ], { right: [0, 1, 2, 3] }));
  C.push(P([b('Einschätzung: '), t(verdict)], { spacing: { before: 100, after: 60 } }));
  C.push(P([b('Nächster Schritt: '), t(next)]));
};
dealBlock('Deal 1: Einfamilienhaus Lohne – bester Kandidat', loh,
  'Lohne, Baujahr 1946, 224 m² Wohnfläche + 121 m² Nutzfläche, 10 Zimmer, 524 m² Grundstück, Energieklasse H, Gasheizung von 2024. Angebot 195.000 € + 3,57 % Käuferprovision (Haskamp Immobilien). Nur ca. 870 €/m².',
  `Der Quadratmeterpreis ist extrem niedrig. Mit einer umfassenden Sanierung (750 €/m², inkl. Puffer ca. ${eur(loh.reno)}) und einem vorsichtigen Verkaufspreis von 2.200 €/m² bleibt ein Gewinn von ca. ${eur(loh.gewinn)} vor Steuern (${pct(loh.marge)}). Das ist knapp unter der Zielmarge – bei einem Kaufpreis von ca. 175.000 € wäre sie erreicht. Großer Hebel: Aufteilung in zwei Wohnungen (je ca. 110 m²) – zwei Einheiten verkaufen oder vermieten sich leichter. Risiko: hoher Eigenkapitalbedarf (ca. ${eur(loh.ek)}) und große Baustelle.`,
  'Exposé, Grundrisse und Energieausweis anfordern; mit Bausachverständigem besichtigen; beim Bauamt Lohne fragen, ob eine Aufteilung in zwei Wohnungen genehmigungsfähig ist; Angebot um 170.000–175.000 €.');
dealBlock('Deal 2: Bungalow Damme-West', bun,
  'Damme-West, 139,5 m², 4 Zimmer, 748 m² Grundstück, Baujahr 1968 bzw. 1979 (Angaben widersprüchlich), „Entwicklungspotenzial“. Angebot 250.000 €.',
  `Bungalows sind in Damme gefragt (Familien und ältere Käufer). Bei mittlerer Sanierung (600 €/m²) und 2.700 €/m² Verkaufspreis ergibt sich zum Angebotspreis aber ein Verlust von ca. ${eur(Math.abs(bun.gewinn))}. Interessant wird er erst bei ca. ${eur(bun.maxK)}.`,
  'Baujahr, Zustand von Dach, Heizung und Elektrik klären. Wenn das Objekt länger als 3 Monate online ist: Angebot um 185.000 € mit Begründung (Sanierungskosten belegen).');
dealBlock('Deal 3: Doppelhaushälfte Damme (Glückauf)', dhh,
  'Damme, Baujahr 1984, ca. 113 m², 5 Zimmer, 581 m² Grundstück, renovierungsbedürftig. Angebot 235.000 € + 3,57 % Käuferprovision (Immobilien Völker, ImmoNr 261698).',
  `Als reiner Flip deutlich zu teuer (Verlust ca. ${eur(Math.abs(dhh.gewinn))}). Als Live-in-Flip (du ziehst ein, sanierst mit viel Eigenleistung für 350 €/m², verkaufst nach 2 Jahren steuerfrei und privat) wird daraus bei 215.000 € Kaufpreis ein kleines Plus von ca. ${eur(live.ergebnis)} inklusive gesparter Miete – bei nur ca. ${eur(live.ek)} Eigenkapital. Mit ca. 185.000 € Kaufpreis würde der Live-in-Flip die Zielmarge erreichen.`,
  'Gut geeignet als erstes eigenes Projekt, wenn du ohnehin ausziehen möchtest. Preisgespräch mit dem Makler, Sanierungsbedarf (Dach, Fenster, Heizung von 1984?) prüfen lassen.');
dealBlock('Deal 4: Haus in Lemförde – nur mit Grundstücksidee', lem,
  'Lemförde (Landkreis Diepholz), Bahnhofstraße 76, 149 m², 3.136 m² Grundstück. Angebot 236.000 €. Laut Inserat Wasserschaden und Schimmelbefall, Modernisierung nur mit erheblichem Aufwand.',
  `Als Haus-Flip ein klares Verlustgeschäft (ca. ${eur(Math.abs(lem.gewinn))}). Interessant ist allein das große Grundstück: Lässt sich ein Teil als Bauplatz abtrennen und verkaufen, ändert sich die Rechnung komplett. Bodenrichtwerte in Lemförde liegen laut Portalen grob zwischen 18 und 105 €/m² – unbedingt amtlich (BORIS Niedersachsen) prüfen.`,
  'Nur weiterverfolgen, wenn das Bauamt eine Teilung bestätigt. Sonst Finger weg.');
dealBlock('Gegenbeispiel: Wohnung Damme-Ort', etw,
  'Damme-Ort, 83 m², 3 Zimmer, Baujahr ca. 1993, Angebot 190.000 € (Aggregator-Inserat).',
  `Gepflegte Wohnungen zum Marktpreis sind kein Flip: Die Sanierung hebt den Wert kaum, die Nebenkosten bleiben. Ergebnis ca. ${eur(etw.gewinn)}. Dieses Beispiel zeigt, was du nicht kaufen solltest.`,
  'Nicht weiterverfolgen. Wohnungen nur, wenn sie deutlich unsaniert und günstig sind (unter ca. 1.600 €/m²).');
C.push(H2('Beobachtungsliste: Zwangsversteigerungen'));
C.push(BL([b('Steinfeld, Handorfer Straße 81: '), t('Einfamilienhaus, ca. 160 m², 755 m² Grundstück, Verkehrswert 370.000 € (AG Vechta). Der Termin am 24.04.2026 ist vorbei – Status und neuen Termin beim Amtsgericht erfragen.')]));
C.push(BL([b('Dinklage: '), t('Einfamilienhaus, Verkehrswert 495.000 €, Termin laut einem Portal am 27.11.2026 (Az. 10 K 1/25, AG Vechta) – Angabe unbestätigt, im amtlichen Portal prüfen. Für den Start zu groß.')]));
C.push(BL([b('Hüde/Lemförde (AG Diepholz): '), t('Mehrere Termine mit lückenhaften Daten. Wöchentlich auf zvg-portal.de nachsehen.')]));
C.push(P([t('Die Angaben stammen aus Versteigerungsportalen, die sich teils widersprechen. Maßgeblich ist nur die amtliche Bekanntmachung mit Gutachten.', { size: 18, italics: true, color: '666666' })]));

// ---------- 10 Team ----------
C.push(H1('10. Dein Team'));
C.push(table([2600, 4300, 2700], [
  ['Wer', 'Wofür', 'Wann'],
  ['Bausachverständiger', 'Zustand, Schäden, Sanierungskosten vor dem Kauf', 'Vor jedem Gebot'],
  ['Energieberater', 'Energieklasse heute/später, Förderung, Bedarfsausweis', 'Vor Kauf und nach Sanierung'],
  ['Handwerker (2–3 je Gewerk)', 'Elektro, Sanitär/Heizung, Maler/Boden, Trockenbau, Dach, Fenster', 'Netzwerk jetzt aufbauen'],
  ['Steuerberater', 'Privat oder GmbH, § 23 EStG, Drei-Objekt-Grenze, Live-in-Planung', 'Vor dem ersten Kauf'],
  ['Bank', 'Finanzierungsrahmen, Ablöserecht, Tranchen', 'Vor dem ersten Gebot'],
  ['Notar / Anwalt', 'Kaufvertrag, Haftungsausschluss beim Verkauf, ggf. Teilung', 'Kauf und Verkauf'],
  ['Makler / Gutachter vor Ort', 'Realistischer Verkaufspreis, Vermarktung', 'Vor Kauf (Preisschätzung) und beim Verkauf'],
], { zebra: true }));
C.push(P('Tipp: Gute Handwerker sind in der Region knapp. Baue die Beziehungen auf, bevor du ein Haus hast – z. B. indem du für kleine Aufträge zuverlässig und pünktlich zahlst. Wer schnell zahlt, bekommt beim nächsten Mal schneller einen Termin.', { spacing: { before: 100 } }));

// ---------- 11 Risiken ----------
C.push(H1('11. Risiken und wie du sie begrenzt'));
C.push(table([2700, 3300, 3600], [
  ['Risiko', 'Was passiert', 'Gegenmaßnahme'],
  ['Sanierung wird teurer', 'Versteckte Schäden (Feuchte, Leitungen, Asbest) – der Gewinn schmilzt.', 'Bausachverständiger vor Kauf, 15–20 % Puffer, Festpreisangebote'],
  ['Bauzeit verzögert sich', 'Jeder Monat kostet Zinsen und Nebenkosten.', 'Bauzeitenplan, Handwerker früh binden, 2 Monate Zeitpuffer'],
  ['Verkaufspreis sinkt', 'Zinsen steigen, Käufer bekommen schlechtere Finanzierung.', 'Vorsichtiger Verkaufspreis, Plan B Vermietung, nicht zu groß starten'],
  ['Unverkäuflich', 'Objekt passt nicht zur Nachfrage (zu groß, schlechte Lage).', 'Nur Objekte, die Familien in Damme suchen; Plan B rechnen'],
  ['Steuerliche Einstufung', 'Gewerblicher Grundstückshandel, Gewerbesteuer.', 'Steuerberater vorab, Zahl der Verkäufe planen'],
  ['Haftung nach Verkauf', 'Käufer macht Mängel geltend.', 'Mängel offenlegen, Dokumentation, Notarvertrag sauber gestalten'],
  ['Zu viel Eigenkapital in einem Objekt', 'Kein Geld für Überraschungen oder das nächste Projekt.', 'Klein anfangen (Live-in), Liquiditätsreserve'],
], { zebra: true }));

// ---------- 12 Plan ----------
C.push(H1('12. Dein Startplan für die nächsten 90 Tage'));
C.push(table([1800, 7800], [
  ['Wann', 'Was'],
  ['Woche 1–2', 'Termin beim Steuerberater (Privat vs. GmbH, Live-in-Flip-Zeitplan) und bei zwei Banken (Finanzierungsrahmen, Ablöserecht, Live-in vs. Flip).'],
  ['Woche 2–4', 'Grundstücksmarktbericht kaufen, 20 sanierte Vergleichsobjekte in Damme/Lohne mit Preis, Baujahr, Fläche und Energieklasse in eine Tabelle eintragen. Verkaufspreis-Annahmen im Rechner anpassen.'],
  ['Woche 2–6', 'Team aufbauen: einen Bausachverständigen, einen Energieberater und je 2–3 Handwerker pro Gewerk kennenlernen.'],
  ['Woche 3', 'Einfamilienhaus Lohne: Exposé anfordern, besichtigen, beim Bauamt die Aufteilung klären.'],
  ['ab Woche 3', 'Suchaufträge einrichten, zvg-portal.de jeden Montag prüfen, Anschreiben in einer zweiten Version „Kauf mit Auszug“ verschicken.'],
  ['Woche 4–12', '10 Besichtigungen, 5 vollständige Kalkulationen im Rechner, 3 Angebote abgeben – jedes höchstens zum „Max. Kaufpreis“.'],
  ['Tag 90', 'Entscheidung: Live-in-Flip starten oder erstes Flip-Objekt kaufen – oder bewusst weitersuchen. Kein Kauf ist besser als ein schlechter Kauf.'],
], { zebra: true }));
C.push(gap());
C.push(box('Deine Deal-Kriterien zum Ausschneiden', [
  '✓ Kaufpreis ≤ Max. Kaufpreis aus dem Rechner (≈ 70 % des Verkaufspreises minus Sanierung)',
  '✓ Marge ≥ 15 % vor Steuern – mit vorsichtigem Verkaufspreis',
  '✓ Sanierung überschaubar (400–700 €/m²), keine Statik-, Schimmel- oder Feuchte-Großbaustelle beim ersten Projekt',
  '✓ Objekt, das Familien in Damme und Umgebung suchen (110–140 m², Garten, gute Lage)',
  '✓ Plan B Vermietung trägt sich',
  '✓ Bausachverständiger war drin, Steuerberater und Bank haben zugestimmt',
]));

// ---------- Quellen ----------
C.push(H1('Quellen'));
const src = [
  'immowelt: Immobilienpreise und Inserate Damme, Lohne, Lemförde, Stemwede (Abruf Nov. 2026)',
  'Immobilien Völker (Lohne): Doppelhaushälfte Damme, ImmoNr 261698',
  'Haskamp Immobilien (Lohne): Einfamilienhaus mit Garage und Keller in Lohne',
  'Engel & Völkers: Immobilienpreise Lohne (Oldenburg), Okt. 2026',
  'immowelt Pressemitteilung 2026: „Gute Energiebilanz zahlt sich aus: Bis zu 20 Prozent Preisaufschlag für effiziente Immobilien“',
  'Institut der deutschen Wirtschaft: IW-Wohnindex Q1 2026',
  'Ratgeber zu Sanierungskosten 2026 (u. a. baufivergleich.de, cleverbauen24.de, handwerker-kontakte.de)',
  'Ratgeber zur BEG-Reform 21.07.2026 und zum Gebäudemodernisierungsgesetz (u. a. baukostenplan.de, energie-experten.org, vermieterwelt.de) – Angaben teils widersprüchlich, bei der KfW prüfen',
  'Ratgeber zu Fix & Flip und Steuern (u. a. fix-flip-immobilien.de, schwerd.info); § 23 EStG',
  'versteigerungspool.de, zvgsuche.de, zvg24.net: Zwangsversteigerungen AG Vechta und AG Diepholz',
  'Portale zu Bodenrichtwerten Lemförde (bodenrichtwerte-deutschland.de) – amtlich: BORIS Niedersachsen',
];
src.forEach(s => C.push(BL(t(s, { size: 18 }))));

const doc = new Document({
  creator: 'Fix & Flip Plan',
  title: 'Fix & Flip – Strategieplan Damme',
  styles: {
    default: { document: { run: { font: FONT, size: 21 } } },
    paragraphStyles: [
      { id: 'Heading1', name: 'Heading 1', basedOn: 'Normal', next: 'Normal', quickFormat: true, run: { font: FONT, size: 32, bold: true }, paragraph: { outlineLevel: 0 } },
      { id: 'Heading2', name: 'Heading 2', basedOn: 'Normal', next: 'Normal', quickFormat: true, run: { font: FONT, size: 25, bold: true }, paragraph: { outlineLevel: 1 } },
    ],
  },
  features: { updateFields: true },
  numbering: { config: [
    { reference: 'bul', levels: [{ level: 0, format: LevelFormat.BULLET, text: '•', alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 360, hanging: 260 } } } }] },
    { reference: 'num', levels: [{ level: 0, format: LevelFormat.DECIMAL, text: '%1.', alignment: AlignmentType.LEFT, style: { run: { bold: true, color: ACC }, paragraph: { indent: { left: 360, hanging: 300 } } } }] },
    { reference: 'num2', levels: [{ level: 0, format: LevelFormat.DECIMAL, text: '%1.', alignment: AlignmentType.LEFT, style: { run: { bold: true, color: ACC }, paragraph: { indent: { left: 360, hanging: 300 } } } }] },
  ] },
  sections: [{
    properties: { page: { size: { width: 11906, height: 16838 }, margin: { top: 1134, bottom: 1134, left: 1134, right: 1134 } } },
    footers: { default: new Footer({ children: [new Paragraph({ alignment: AlignmentType.RIGHT, children: [t('Fix & Flip – Strategieplan Damme · Seite ', { size: 16, color: '888888' }), new TextRun({ children: [PageNumber.CURRENT], font: FONT, size: 16, color: '888888' })] })] }) },
    children: C,
  }],
});
Packer.toBuffer(doc).then(buf => { fs.writeFileSync(OUT, buf); console.log('geschrieben:', OUT); });
