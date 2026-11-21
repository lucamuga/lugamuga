// Erstellt den Gebietsplan für Kauf mit Rückmietung (Damme und Umgebung) als Word-Dokument.
// Aufruf: NODE_PATH=$(npm root -g) node gebietsplan.js AUSGABE.docx
const fs = require('fs');
const path = require('path');
const { Document, Packer, Paragraph, TextRun, PageBreak, Footer, PageNumber, AlignmentType } = require('docx');
const { FONT, ACC, t, b, P, H1, H2, BL, NL, gap, table, box, numbering, styles } = require(path.join(__dirname, 'doc_helfer.js'));

const OUT = process.argv[2];
const C = [];

// ---------- Titel ----------
C.push(new Paragraph({ spacing: { before: 600, after: 120 }, children: [new TextRun({ text: 'KAUF MIT RÜCKMIETUNG', font: FONT, size: 24, bold: true, color: ACC, characterSpacing: 60 })] }));
C.push(new Paragraph({ spacing: { after: 160 }, children: [new TextRun({ text: 'Gebietsplan Damme und Umgebung', font: FONT, size: 48, bold: true, color: '1C2430' })] }));
C.push(P([t('Wo du ältere Eigentümer mit passenden Häusern findest – und in welcher Reihenfolge du die Gebiete angehst. Stand November 2026.', { size: 22, color: '555555' })]));
C.push(gap());
C.push(box('Kurzfassung', [
  [b('Die besten Gebiete sind die Siedlungen der 1950er und 1960er Jahre. '), t('Wer damals gebaut hat, ist heute meist zwischen 75 und 95 Jahre alt oder die Häuser sind bereits an Kinder vererbt, die nicht selbst einziehen. Genau dort entsteht der Bedarf, ein Haus zu verkaufen und trotzdem darin wohnen zu bleiben.')],
  [b('Top-Gebiete: '), t('Damme-Glückauf (ehemalige Bergarbeitersiedlung ab 1951) und Clemens-August-Dorf (Siedlung ab 1950). Danach Damme-Zentrum/Süd (Krankenhaus, Ärzte, Einkauf fußläufig) und die Bergarbeitersiedlungen Handorf-Langenberg (Holdorf) und Steinfeld-Süd.')],
  [b('Wichtig: '), t('In Damme ist nur rund jeder Fünfte 65 Jahre oder älter (ca. 3.430 von 17.700 Einwohnern, Zahlenspiegel 2025) – weniger als im Bundesschnitt (ca. 23 %). Die Zielgruppe ist klein, deshalb lohnt sich die gezielte Konzentration auf die alten Siedlungen.')],
]));

// ---------- 1 Kriterien ----------
C.push(H1('1. Was ein gutes Gebiet für Rückmietung ausmacht'));
C.push(table([2600, 4300, 2700], [
  ['Kriterium', 'Warum es wichtig ist', 'Woran du es erkennst'],
  ['Häuser aus 1950–1980', 'Die Erstbesitzer sind heute im Rentenalter; viele Häuser sind abbezahlt, aber renovierungsbedürftig.', 'kleine 1½-geschossige Siedlungshäuser, Klinker, große Gärten, alte Fenster'],
  ['Infrastruktur fußläufig', 'Ältere Mieter bleiben länger und zufriedener wohnen, wenn Arzt, Apotheke, Bäcker und Kirche nah sind.', 'Zentrum, Damme-Süd (Krankenhaus Lindenstraße), Ortskerne'],
  ['Große Grundstücke', 'Wertreserve und spätere Chance (Teilung, Anbau, Verkauf an Familien).', 'Glückauf: Grundstücke um 950–1.050 m² laut aktuellen Inseraten'],
  ['Spätere Nachfrage', 'Nach Auszug des Verkäufers willst du gut vermieten oder verkaufen können.', 'beliebte Stadtteile mit Familien, Schulen, Kitas'],
  ['Moderate Preise', 'Nur dann trägt die Miete die Finanzierung.', 'Häuser unter dem Damme-Schnitt von ca. 2.420 €/m²'],
  ['Einliegerwohnungen / 2 Hausteile', 'Senior wohnt in einem Teil, der andere wird zusätzlich vermietet → besserer Cashflow.', 'zwei Eingänge, zwei Klingeln, Anbauten aus den 1970ern'],
], { zebra: true }));
C.push(H2('Gebiete, die du meiden solltest'));
C.push(BL([b('Neubaugebiete: '), t('junge Familien, keine Verkaufsabsicht, hohe Preise.')]));
C.push(BL([b('Einzelhöfe im Außenbereich (Bauerschaften): '), t('Hofstellen sind schwer zu bewerten und zu finanzieren, bauliche Änderungen sind im Außenbereich stark eingeschränkt, und nach dem Auszug ist die Nachfrage klein.')]));
C.push(BL([b('Ferien- und Wochenendhausgebiete am Dümmer: '), t('Dauerwohnen ist dort oft nicht erlaubt – für eine dauerhafte Vermietung ungeeignet.')]));

// ---------- 2 Ranking ----------
C.push(new Paragraph({ children: [new PageBreak()] }));
C.push(H1('2. Die Gebiete im Überblick'));
C.push(P([t('Einwohnerzahlen laut Zahlenspiegel der Stadt Damme 2025 (Stand 01.01.2025), soweit verfügbar. Mieten und Bodenrichtwerte sind Werte aus Online-Portalen und dienen nur der Orientierung.', { size: 18, italics: true, color: '666666' })]));
C.push(table([1000, 2300, 1100, 2900, 2300], [
  ['Priorität', 'Gebiet', 'Einwohner', 'Warum', 'Achtung'],
  ['A', 'Damme-Glückauf', '1.047', 'Bergarbeitersiedlung ab 1951, bis 1966 erweitert; heute gepflegtes, beliebtes Wohnviertel; große Grundstücke', 'Bodenrichtwert laut Portalen hoch (ca. 240–250 €/m²) → Häuser teurer; Energieklassen oft schlecht'],
  ['A', 'Clemens-August-Dorf', '1.854', 'Siedlung ab 1950 (Kardinal-von-Galen-Siedlungswerk), starkes Wachstum in den 1960ern; größter Siedlungsbereich; Grundschule Sierhausen', 'Mischung aus alten und neueren Häusern – gezielt die älteren Straßenzüge auswählen'],
  ['A–B', 'Damme-Zentrum und Damme-Süd', 'k. A.', 'Krankenhaus St. Elisabeth (Lindenstraße, Damme-Süd), Ärzte, Einkauf, Kirche fußläufig – ideal für ältere Mieter', 'höhere Preise, kleinere Grundstücke'],
  ['B', 'Damme-West', '1.199', 'gemischtes Alter, u. a. Bungalows der 1960er/70er (ebenerdig – sehr seniorengerecht)', 'auch neuere Bebauung'],
  ['B', 'Osterdamme', '1.385', 'größere Ortschaft mit älterem Ortskern; auf Häuser mit Einliegerwohnung achten', 'weiter vom Zentrum entfernt'],
  ['B', 'Rüschendorf (mit Ihlendorf, Kemphausen, Hüde)', '919', 'ältere Ortslagen; Mieten laut Portal ca. 8,59 €/m²', 'Hüde liegt am Dümmer – Ferienhausgebiete ausschließen'],
  ['B', 'Osterfeine', '595', 'dörflich, ältere Häuser; Mieten laut Portal ca. 8,87 €/m²', 'kleine Zielgruppe, längere Wege'],
  ['B', 'Handorf-Langenberg (Gemeinde Holdorf)', 'k. A.', 'Bergarbeitersiedlung 1953–1958 – gleiches Profil wie Glückauf; Preise vor Ort vergleichen', 'außerhalb von Damme; Bahnhof Holdorf in der Nähe'],
  ['B', 'Steinfeld-Süd', 'k. A.', 'Bergarbeiterhäuser von 1957 – gleiches Profil', 'außerhalb von Damme'],
  ['C', 'Borringhausen, Haverbeck, Bexadde, Sierhausen, Holte, Nordhofe', '375 / 256 / 220 / 261 / 52 / 60', 'nur einzelne Häuser in Ortslage interessant', 'überwiegend Außenbereich und Hofstellen'],
  ['–', 'Ferien- und Wochenendhausgebiete am Dümmer (z. B. Teile von Dümmerlohausen, Hüde, Lembruch), Neubaugebiete', '', 'Dauerwohnen oft nicht erlaubt bzw. junge Eigentümer', 'meiden (siehe Kapitel 1) – normale Wohnlagen in Dümmerlohausen einzeln prüfen'],
], { zebra: true }));

// ---------- 3 Steckbriefe ----------
C.push(H1('3. Die Top-Gebiete im Detail'));
C.push(H2('A: Damme-Glückauf'));
C.push(BL('Entstanden ab 1951 als Siedlung für die Bergleute der Grube Damme, mit Mitteln des Marshallplans gefördert; bis 1966 erweitert. Zusammen mit Handorf-Langenberg und Steinfeld-Süd entstanden über 300 Häuser mit mehr als 600 Wohnungen.'));
C.push(BL('Heute ein gefragtes Wohnviertel mit großen Gärten und starker Gemeinschaft (z. B. Schützenfest Glückauf, Knappenkapelle).'));
C.push(BL('Aktuelle Inserate zeigen typische Grundstücke um 950–1.050 m² und Häuser aus 1950–1972 (z. B. das geprüfte Zweifamilienhaus, Baujahr 1950/1972).'));
C.push(BL([b('Ansatz: '), t('Persönliche Briefe an ältere, unmodernisierte Häuser; Kontakt über Schützenverein und Kirchengemeinde. Viele Häuser sind bereits vererbt – auch Erben ansprechen, deren Eltern noch im Haus wohnen.')]));
C.push(H2('A: Clemens-August-Dorf'));
C.push(BL('Nach dem Zweiten Weltkrieg für Vertriebene und Einheimische gegründet (Siedlungswerk Kardinal von Galen), ab den 1960er Jahren stark gewachsen; mit ca. 1.850 Einwohnern der größte Siedlungsbereich in dieser Auswahl.'));
C.push(BL('Ruhige Wohnlage, Grundschule Sierhausen, gute Nachfrage durch Familien – gut für die spätere Vermietung oder den Verkauf.'));
C.push(BL([b('Ansatz: '), t('Zuerst die älteren Straßenzüge (kleine Siedlungshäuser, große Gärten) identifizieren, dann gezielt anschreiben.')]));
C.push(H2('A–B: Damme-Zentrum und Damme-Süd'));
C.push(BL('Krankenhaus St. Elisabeth (247 Betten) in der Lindenstraße, dazu Ärzte, Apotheken und Einkauf fußläufig – für ältere Menschen der attraktivste Wohnort.'));
C.push(BL([b('Ansatz: '), t('Gut für Senioren, die bleiben wollen; wegen der höheren Preise nur, wenn der Rechner aufgeht oder eine Einliegerwohnung vorhanden ist.')]));

// ---------- 4 Akquise-Plan ----------
C.push(new Paragraph({ children: [new PageBreak()] }));
C.push(H1('4. Dein Akquise-Plan in vier Phasen'));
C.push(table([1700, 2600, 5300], [
  ['Zeitraum', 'Gebiet', 'Was du tust'],
  ['Monat 1–2', 'Damme-Glückauf', 'Straßen ablaufen, passende Häuser notieren (Baujahr-Optik, Zustand); 30–40 persönliche Briefe; Aushang/Gespräch über Kirchengemeinde und Seniorentreff'],
  ['Monat 2–3', 'Clemens-August-Dorf', 'gleiches Vorgehen; ältere Straßenzüge zuerst; 30–40 Briefe'],
  ['Monat 3–4', 'Damme-Zentrum / Damme-Süd / Damme-West (Bungalows)', 'Fokus auf Häuser mit Einliegerwohnung und ebenerdige Bungalows; Kontakt zu Hausärzten, Pflegediensten und Seniorenberatung (Flyer auslegen, wenn erlaubt)'],
  ['Monat 4–6', 'Handorf-Langenberg, Steinfeld-Süd, Osterdamme', 'Wiederholung im Umland; Anzeige im Gemeindeblatt / OV'],
  ['laufend', 'alle Gebiete', 'Antworten nachhalten, zweite Welle nach 6 Monaten nur bei Häusern ohne Absage; Empfehlungen zufriedener Verkäufer nutzen'],
], { zebra: true }));
C.push(H2('So findest du die richtigen Häuser'));
C.push(BL('Zu Fuß oder mit dem Rad durch die Straßen gehen – nicht mit dem Auto langsam vorbeifahren (wirkt aufdringlich).'));
C.push(BL('Merkmale: kleines 1½-geschossiges Siedlungshaus, alte Fenster/Haustür, gepflegter, aber älterer Garten, zweiter Eingang oder Anbau, Handlauf an der Eingangstreppe.'));
C.push(BL('Nur die Hausnummer und Auffälligkeiten notieren – keine Fotos von Personen, keine Namen aus dem Briefkasten sammeln.'));
C.push(BL('Brief adressieren „An die Eigentümer des Hauses [Straße Nr.]“, wenn du den Namen nicht kennst.'));
C.push(H2('Kontakt mit Respekt – damit dein Ruf gut bleibt'));
C.push(BL('Nur per Post, nie unangemeldet an der Tür klingeln.'));
C.push(BL('Briefkästen mit „Keine Werbung“ respektieren; wer absagt, wird nie wieder angeschrieben (Liste führen).'));
C.push(BL('Pro Haus höchstens zwei Briefe im Jahr.'));
C.push(BL('Im Gespräch immer anbieten, dass Kinder oder eine Vertrauensperson dabei sind.'));
C.push(gap());
C.push(box('Realistische Erwartung', [
  'Bei persönlichen, ehrlichen Briefen sind Antwortquoten von wenigen Prozent üblich. Aus 100 Briefen entstehen vielleicht 2–5 Gespräche und daraus ein passendes Objekt. Das ist normal – entscheidend ist, dass jedes Gespräch fair verläuft und sich herumspricht.',
], 'info'));

// ---------- 5 Hilfsmittel ----------
C.push(H1('5. Was du dafür schon hast'));
C.push(BL([b('Anschreiben und Beiblatt: '), t('Kauf-mit-Rueckmietung/02_Anschreiben-Verkaeufer')]));
C.push(BL([b('Rechner: '), t('Kauf-mit-Rueckmietung/01_Rechner (Rückmietung) bzw. Kapitalanlage-Vermietung (Objekt-Rechner mit Spannen)')]));
C.push(BL([b('Checkliste Objektprüfung: '), t('Kauf-mit-Rueckmietung/03_Objektpruefung')]));

// ---------- Quellen ----------
C.push(H1('Quellen'));
[
  'Stadt Damme: Zahlenspiegel 2025 (Einwohner der Ortsteile und Altersstruktur, Stand 01.01.2025) – Angaben über Suchergebnisse',
  'Erzbergbau Damme: Bergarbeitersiedlung früher; Kultur & Tradition (erzbergbau-damme.de)',
  'OM online: „Erst die Währungsreform brachte die Lösung“ (Siedlungsbau nach 1950)',
  'Wikipedia: Clemens-August-Dorf; Damme (Dümmer); St. Barbara (Handorf-Langenberg)',
  'Schützenverein Handorf-Langenberg: Geschichte (Siedlung Langenberg 1953–1958)',
  'Krankenhaus St. Elisabeth gGmbH, Lindenstraße 3–7, Damme (Deutsches Krankenhausverzeichnis, Das Örtliche)',
  'Mietpreise Ortsteile: miete-aktuell.de (Rüschendorf, Osterfeine, Dümmerlohausen); Damme gesamt: ImmoScout24, Engel & Völkers, miet-check.de',
  'Bodenrichtwerte: Portalangaben (miete-aktuell.de) – amtlich über BORIS Niedersachsen prüfen',
].forEach(s => C.push(BL(t(s, { size: 18 }))));

const doc = new Document({
  creator: 'Gebietsplan', title: 'Gebietsplan Rückmietung Damme', styles, numbering,
  sections: [{
    properties: { page: { size: { width: 11906, height: 16838 }, margin: { top: 1134, bottom: 1134, left: 1134, right: 1134 } } },
    footers: { default: new Footer({ children: [new Paragraph({ alignment: AlignmentType.RIGHT, children: [t('Gebietsplan Rückmietung Damme · Seite ', { size: 16, color: '888888' }), new TextRun({ children: [PageNumber.CURRENT], font: FONT, size: 16, color: '888888' })] })] }) },
    children: C,
  }],
});
Packer.toBuffer(doc).then(buf => { fs.writeFileSync(OUT, buf); console.log('geschrieben:', OUT); });
