// ============================================================
// ROUTE: POST /generate-sichtbarkeit
// Sichtbarkeits- & Relevanztest – DOCX Generator
// Einbinden in bestehende server.js nach der /generate Route
// ============================================================

app.post('/generate-sichtbarkeit', async (req, res) => {
  try {
    const d = req.body;

    // Eingabefelder
    const vorname    = d.vorname    || '';
    const nachname   = d.nachname   || '';
    const email      = d.email      || '';
    const datum      = d.datum      || new Date().toLocaleDateString('de-AT');
    const klasse     = d.klasse     || 'II';   // I, II oder III
    const score      = parseInt(d.score) || 0;

    // Antworten: { "F01": "ja", "F02": "nein", ... }
    const antworten  = d.antworten  || {};

    // ── FARBEN ──────────────────────────────────────────────
    const DUNKEL  = '0D1B3E';
    const GOLD    = 'C8A951';
    const GOLD_H  = 'F5EDD6';
    const WEISS   = 'FFFFFF';
    const TEXT    = '2C2C2C';
    const GRAU_T  = '666666';
    const GRAU    = 'F5F6FA';
    const BORDER  = 'E0E4EF';
    const ROT_H   = 'FDECEA';
    const ROT_B   = 'E24B4A';
    const AMB_H   = 'FEF3E2';
    const AMB_B   = 'BA7517';
    const GRN_H   = 'EAF3DE';
    const GRN_B   = '3B6D11';

    const klasseHell   = klasse==='I' ? ROT_H  : klasse==='II' ? AMB_H  : GRN_H;
    const klasseBorder = klasse==='I' ? ROT_B  : klasse==='II' ? AMB_B  : GRN_B;

    const klasseText = klasse==='I'
      ? 'Handlungsbedarf'
      : klasse==='II'
      ? 'Konkrete Hebel vorhanden'
      : 'Starke Ausgangslage';

    const headline = klasse==='I'
      ? 'Hier liegt ungenutztes Potenzial \u2013 das l\u00e4sst sich \u00e4ndern.'
      : klasse==='II'
      ? 'Solide Basis \u2013 jetzt kommt der entscheidende Schritt.'
      : 'Ausgezeichnet \u2013 Sie sind nah dran. Jetzt systematisch.';

    // ── KARTEN-TEXTE ────────────────────────────────────────
    const karten = {
      I: {
        bedeutet:   'Keine klare Positionierung, kein erkennbarer Nutzen, kein gef\u00fchrter Auftritt nach au\u00dfen.',
        wirkt:      'Zuf\u00e4llig. Unstrukturiert. Wie ein vernachl\u00e4ssigter Account.',
        konsequenz: 'Ihr Profil arbeitet aktuell gegen Sie. Hier fehlt die Basis.',
        jetzt:      'Basis schaffen. Positionierung kl\u00e4ren. In 15 Minuten zeige ich Ihnen genau wo.'
      },
      II: {
        bedeutet:   'Man erkennt, wer Sie sind \u2013 aber nicht, warum Sie relevant sind.',
        wirkt:      'Solide, aber ohne Wirkung. Besucher verstehen Sie, handeln aber nicht.',
        konsequenz: 'Sie verschenken t\u00e4glich unbemerkt Gespr\u00e4chsanl\u00e4sse und Anfragen.',
        jetzt:      'Den Hebel finden. Wenige Anpassungen, gro\u00dfe Wirkung. Ich zeige Ihnen welche.'
      },
      III: {
        bedeutet:   'Guter Auftritt, klare Struktur \u2013 aber noch kein aktives Vertriebsinstrument.',
        wirkt:      'Professionell, aber passiv. Es \u00fcberzeugt \u2013 l\u00f6st aber keine Gespr\u00e4che aus.',
        konsequenz: 'Hier beginnt gezielte Kundengewinnung. Die Basis ist stark \u2013 jetzt Systematik.',
        jetzt:      'Vom guten Profil zum verl\u00e4sslichen Vertriebsraum. Das ist der n\u00e4chste Schritt.'
      }
    };
    const k = karten[klasse] || karten['II'];

    // ── FRAGEN ──────────────────────────────────────────────
    const fragen = [
      { kat: 'Erster Eindruck',       text: 'Wer & Womit erkennbar',                  key: 'F01' },
      { kat: 'Erster Eindruck',       text: 'Header zeigt Positionierung',             key: 'F02' },
      { kat: 'Erster Eindruck',       text: 'Profil wirkt wie Entscheider',            key: 'F03' },
      { kat: 'Positionierung',        text: 'Slogan zeigt klaren Nutzen',              key: 'F04' },
      { kat: 'Positionierung',        text: 'Problem wird klar benannt',               key: 'F05' },
      { kat: 'Positionierung',        text: 'Keine Floskeln im Profiltext',            key: 'F06' },
      { kat: 'Profil als Instrument', text: 'Klarer CTA vorhanden',                   key: 'F07' },
      { kat: 'Profil als Instrument', text: 'Anfragen in letzten 30 Tagen',           key: 'F08' },
      { kat: 'Profil als Instrument', text: 'W\u00fcrde sich selbst kontaktieren',    key: 'F09' },
    ];

    // ── HILFSFUNKTIONEN ─────────────────────────────────────
    const {
      Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
      AlignmentType, BorderStyle, WidthType, ShadingType, VerticalAlign, PageBreak
    } = require('docx');

    const noBorder  = { style: BorderStyle.NIL, size: 0, color: WEISS };
    const noBorders = { top: noBorder, bottom: noBorder, left: noBorder, right: noBorder };
    const thin = c => ({ style: BorderStyle.SINGLE, size: 1, color: c || BORDER });
    const allB = c => ({ top: thin(c), bottom: thin(c), left: thin(c), right: thin(c) });

    const P = (text, opts = {}) => new Paragraph({
      alignment: opts.align || AlignmentType.LEFT,
      spacing:   { before: opts.before || 0, after: opts.after || 120 },
      children:  [new TextRun({ text, font: 'Arial', size: opts.size || 22,
                    bold: opts.bold || false, color: opts.color || TEXT,
                    italics: opts.italic || false })]
    });
    const SP = n => new Paragraph({ spacing: { before: 0, after: n },
                    children: [new TextRun({ text: '', font: 'Arial', size: 20 })] });

    const C = (children, opts = {}) => new TableCell({
      borders:       opts.b || noBorders,
      shading:       opts.fill ? { fill: opts.fill, type: ShadingType.CLEAR } : undefined,
      width:         opts.w   ? { size: opts.w, type: WidthType.DXA }        : undefined,
      margins:       { top: opts.mt||100, bottom: opts.mb||100, left: opts.ml||160, right: opts.mr||160 },
      verticalAlign: opts.va || VerticalAlign.TOP,
      children
    });
    const TR  = cells => new TableRow({ children: cells });
    const TBL = (rows, colWidths) => new Table({
      width: { size: 9026, type: WidthType.DXA },
      columnWidths: colWidths,
      borders: { top: noBorder, bottom: noBorder, left: noBorder, right: noBorder,
                 insideH: noBorder, insideV: noBorder },
      rows
    });

    // ── BLÖCKE ──────────────────────────────────────────────
    const headerBlock = () => TBL([TR([C([
      SP(120),
      new Paragraph({ spacing: { before: 0, after: 60 }, children: [
        new TextRun({ text: 'SICHTBARKEITS- & RELEVANZTEST', font: 'Arial', size: 18,
                      color: GOLD, bold: true, allCaps: true })]}),
      new Paragraph({ spacing: { before: 0, after: 80 }, children: [
        new TextRun({ text: `Pers\u00f6nliche Auswertung f\u00fcr ${vorname} ${nachname}`,
                      font: 'Arial', size: 26, color: WEISS, bold: true })]}),
      new Paragraph({ spacing: { before: 0, after: 40 }, children: [
        new TextRun({ text: `${datum} \u00b7 Klartext Mittelstand \u00b7 Rudolf Pusterhofer`,
                      font: 'Arial', size: 18, color: '8AACCC' })]}),
      SP(120),
    ], { fill: DUNKEL, b: allB(DUNKEL), w: 9026, ml: 280, mr: 280 })])], [9026]);

    const scoreBlock = () => TBL([TR([
      C([
        new Paragraph({ spacing: { before: 0, after: 40 }, children: [
          new TextRun({ text: String(score), font: 'Arial', size: 96, bold: true, color: DUNKEL })]}),
        P('von 9 positiven Antworten', { size: 20, color: GRAU_T }),
      ], { fill: GRAU, b: allB(BORDER), w: 2800, ml: 200, mr: 200 }),
      C([SP(40)], { fill: WEISS, w: 200 }),
      C([
        SP(40),
        new Paragraph({ spacing: { before: 0, after: 80 }, children: [
          new TextRun({ text: `KLASSE ${klasse} \u00b7 ${klasseText.toUpperCase()}`,
                        font: 'Arial', size: 18, bold: true, color: klasseBorder, allCaps: true })]}),
        new Paragraph({ spacing: { before: 0, after: 60 }, children: [
          new TextRun({ text: headline, font: 'Arial', size: 26, bold: true, color: DUNKEL })]}),
        SP(40),
      ], { fill: klasseHell,
           b: { top: noBorder, bottom: noBorder, right: noBorder,
                left: { style: BorderStyle.SINGLE, size: 16, color: klasseBorder }},
           w: 6026, ml: 240, mr: 200 }),
    ])], [2800, 200, 6026]);

    const karte = (label, text) => C([
      new Paragraph({ spacing: { before: 0, after: 60 }, children: [
        new TextRun({ text: label, font: 'Arial', size: 17, bold: true,
                      color: klasseBorder, allCaps: true })]}),
      P(text, { size: 20, color: TEXT, after: 60 }),
    ], { fill: klasseHell, b: allB(BORDER), w: 4313, ml: 160, mr: 160, mt: 140, mb: 140 });

    const kartenBlock = () => [
      TBL([TR([
        karte('Was das bedeutet',          k.bedeutet),
        C([SP(0)], { fill: WEISS, w: 400 }),
        karte('Wie Ihr Profil wirkt',       k.wirkt),
      ])], [4313, 400, 4313]),
      SP(120),
      TBL([TR([
        karte('Unternehmerische Konsequenz', k.konsequenz),
        C([SP(0)], { fill: WEISS, w: 400 }),
        karte('Was jetzt z\u00e4hlt',      k.jetzt),
      ])], [4313, 400, 4313]),
    ];

    const antwortenBlock = () => {
      const rows = [];
      let lastKat = '';
      fragen.forEach((f, i) => {
        if (f.kat !== lastKat) {
          lastKat = f.kat;
          rows.push(TR([C([new Paragraph({ spacing: { before: 60, after: 60 }, children: [
            new TextRun({ text: f.kat.toUpperCase(), font: 'Arial', size: 18,
                          bold: true, color: GOLD })]})
          ], { fill: DUNKEL, b: allB(DUNKEL), w: 9026, ml: 160 })]));
        }
        const raw = (antworten[f.key] || '').toLowerCase().trim();
        const isJa   = raw === 'ja'   || raw === 'true';
        const isNein = raw === 'nein' || raw === 'false';
        const aText  = isJa ? 'JA \u2714' : isNein ? 'NEIN \u2718' : 'UNSICHER';
        const aColor = isJa ? '2E7D32'    : isNein ? 'C62828'      : '8a6200';
        const aFill  = isJa ? 'E8F5E9'    : isNein ? 'FFEBEE'      : 'FFF8E1';
        rows.push(TR([
          C([P(f.text, { size: 20, color: TEXT, after: 0 })],
            { fill: WEISS, b: { top: noBorder, bottom: thin(BORDER), left: noBorder, right: noBorder },
              w: 7226, ml: 160 }),
          C([new Paragraph({ alignment: AlignmentType.CENTER, spacing: { before: 0, after: 0 },
              children: [new TextRun({ text: aText, font: 'Arial', size: 20,
                          bold: true, color: aColor })]})],
            { fill: aFill, b: allB(BORDER), w: 1800, ml: 80, mr: 80 }),
        ]));
      });
      return new Table({
        width: { size: 9026, type: WidthType.DXA },
        columnWidths: [7226, 1800],
        borders: { top: noBorder, bottom: noBorder, left: noBorder, right: noBorder,
                   insideH: noBorder, insideV: noBorder },
        rows
      });
    };

    const gespraechBlock = () => TBL([TR([C([
      SP(80),
      new Paragraph({ alignment: AlignmentType.CENTER, spacing: { before: 0, after: 80 }, children: [
        new TextRun({ text: 'Was nach unserem Gespr\u00e4ch anders ist',
                      font: 'Arial', size: 24, bold: true, color: DUNKEL })]}),
      new Paragraph({ spacing: { before: 0, after: 60 }, children: [
        new TextRun({ text: '\u2192 Sie wissen genau, welche 2\u20133 Stellen in Ihrem Profil heute Anfragen verhindern',
                      font: 'Arial', size: 20, color: TEXT })]}),
      new Paragraph({ spacing: { before: 0, after: 60 }, children: [
        new TextRun({ text: '\u2192 Sie haben einen klaren Satz f\u00fcr Ihren Slogan \u2013 der Entscheider sofort anspricht',
                      font: 'Arial', size: 20, color: TEXT })]}),
      new Paragraph({ spacing: { before: 0, after: 60 }, children: [
        new TextRun({ text: '\u2192 Sie wissen, ob LinkedIn f\u00fcr Sie der richtige Hebel ist \u2013 oder wo er wirklich liegt',
                      font: 'Arial', size: 20, color: TEXT })]}),
      new Paragraph({ spacing: { before: 0, after: 80 }, children: [
        new TextRun({ text: '\u2192 Sie gehen mit einem konkreten n\u00e4chsten Schritt heraus \u2013 kein Konzept, keine Theorie',
                      font: 'Arial', size: 20, color: TEXT })]}),
      SP(80),
    ], { fill: GOLD_H, b: allB(GOLD), w: 9026, ml: 240, mr: 240 })])], [9026]);

    const ctaBlock = () => TBL([TR([C([
      SP(80),
      new Paragraph({ alignment: AlignmentType.CENTER, spacing: { before: 0, after: 80 }, children: [
        new TextRun({ text: 'Jetzt 15-Minuten-Gespr\u00e4ch buchen',
                      font: 'Arial', size: 22, bold: true, color: WEISS })]}),
      new Paragraph({ alignment: AlignmentType.CENTER, spacing: { before: 0, after: 60 }, children: [
        new TextRun({ text: 'calendly.com/pusterhofer/wahlen-sie-ihren-termin',
                      font: 'Arial', size: 20, bold: true, color: GOLD })]}),
      new Paragraph({ alignment: AlignmentType.CENTER, spacing: { before: 0, after: 80 }, children: [
        new TextRun({ text: 'Kostenfreies Gespr\u00e4ch \u00b7 Kein Verkaufsgespr\u00e4ch',
                      font: 'Arial', size: 18, color: '8AACCC', italics: true })]}),
      SP(80),
    ], { fill: DUNKEL, b: allB(DUNKEL), w: 9026, ml: 240, mr: 240 })])], [9026]);

    const sigBlock = () => TBL([TR([C([
      new Paragraph({ spacing: { before: 0, after: 60 }, children: [
        new TextRun({ text: 'Rudolf Pusterhofer', font: 'Arial', size: 24, bold: true, color: DUNKEL })]}),
      new Paragraph({ spacing: { before: 0, after: 40 }, children: [
        new TextRun({ text: 'Klartext Mittelstand', font: 'Arial', size: 20, color: GOLD, bold: true })]}),
      P('Einordnung f\u00fcr Unternehmer, die Entscheidungen tragen.',
        { size: 19, color: GRAU_T, italic: true, after: 60 }),
      P('ausderpraxis.com \u00b7 +43 664 88 366 140', { size: 18, color: GRAU_T, after: 40 }),
      new Paragraph({ spacing: { before: 0, after: 0 }, children: [
        new TextRun({ text: 'Eisenstadt \u00b7 \u00d6sterreich', font: 'Arial', size: 18, color: GRAU_T })]}),
    ], { fill: WEISS,
         b: { top: thin(GOLD), bottom: noBorder, left: noBorder, right: noBorder },
         w: 9026, ml: 0, mt: 160 }),
    ])], [9026]);

    // ── DOKUMENT ERSTELLEN ──────────────────────────────────
    const doc = new Document({
      numbering: { config: [] },
      styles: { default: { document: { run: { font: 'Arial', size: 22 } } } },
      sections: [{
        properties: {
          page: {
            size: { width: 11906, height: 16838 },
            margin: { top: 1000, right: 1000, bottom: 1000, left: 1000 }
          }
        },
        children: [
          // SEITE 1
          headerBlock(), SP(240),
          scoreBlock(),  SP(240),
          ...kartenBlock(),

          // SEITE 2
          new Paragraph({ children: [new PageBreak()], spacing: { before: 0, after: 0 } }),
          headerBlock(), SP(240),
          new Paragraph({ spacing: { before: 0, after: 120 }, children: [
            new TextRun({ text: 'Ihre Antworten im \u00dcberblick',
                          font: 'Arial', size: 28, bold: true, color: DUNKEL })]}),
          P('Diese 9 Punkte zeigen, wo Ihr LinkedIn-Profil heute steht.',
            { size: 20, color: GRAU_T, after: 200 }),
          antwortenBlock(), SP(300),
          gespraechBlock(), SP(240),
          ctaBlock(),       SP(300),
          sigBlock(),
        ]
      }]
    });

    const buffer = await Packer.toBuffer(doc);
    const filename = `Sichtbarkeits-Auswertung-${nachname}-Klasse${klasse}.docx`;

    res.setHeader('Content-Type',
      'application/vnd.openxmlformats-officedocument.wordprocessingml.document');
    res.setHeader('Content-Disposition', `attachment; filename="${filename}"`);
    res.send(buffer);

  } catch (err) {
    console.error('generate-sichtbarkeit error:', err);
    res.status(500).json({ error: err.message });
  }
});
