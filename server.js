const express = require('express');
const { Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  AlignmentType, BorderStyle, WidthType, ShadingType, VerticalAlign,
  Header, Footer, PageBreak } = require('docx');

const app = express();
app.use(express.json());

const DUNKEL="1A1A2E", GOLD="C8A951", GOLD_HELL="F5EDD6", GRAU="F0F0EA";
const ROT="C0392B", ROT_HELL="FDECEA", ORANGE="D97706", ORANGE_HELL="FEF3E2";
const GRUEN="2E7D5E", GRUEN_HELL="E8F5EE", WEISS="FFFFFF", TEXT="2C2C2C", TEXT_GRAU="666666", BORDER="DDDDCC";

const noBorder={style:BorderStyle.NIL,size:0,color:"FFFFFF"};
const noBorders={top:noBorder,bottom:noBorder,left:noBorder,right:noBorder};
function thin(c){return {style:BorderStyle.SINGLE,size:1,color:c||BORDER};}
function allB(c){return {top:thin(c),bottom:thin(c),left:thin(c),right:thin(c)};}
function leftB(c,s){return {top:noBorder,bottom:noBorder,right:noBorder,left:{style:BorderStyle.SINGLE,size:s||12,color:c||GOLD}};}

function C(children,opts){
  opts=opts||{};
  return new TableCell({
    borders:opts.b||noBorders,
    shading:opts.fill?{fill:opts.fill,type:ShadingType.CLEAR}:undefined,
    width:opts.w?{size:opts.w,type:WidthType.DXA}:undefined,
    verticalAlign:opts.v||VerticalAlign.TOP,
    margins:{top:opts.mt||80,bottom:opts.mb||80,left:opts.ml||160,right:opts.mr||160},
    columnSpan:opts.span,
    children:children
  });
}

function P(txt,opts){
  opts=opts||{};
  return new Paragraph({
    alignment:opts.align||AlignmentType.LEFT,
    spacing:{before:opts.before||0,after:opts.after||60},
    children:[new TextRun({text:String(txt||''),font:'Arial',size:opts.size||22,bold:opts.bold||false,italics:opts.italic||false,color:opts.color||TEXT})]
  });
}

function SP(n){return new Paragraph({spacing:{before:0,after:n||120},children:[new TextRun({text:'',font:'Arial',size:20})]});}

function ampelFarbe(a){
  if(a==='Hier haben wir Nachholbedarf') return ROT;
  if(a==='Bin mir nicht sicher') return ORANGE;
  return GRUEN;
}
function ampelFill(a){
  if(a==='Hier haben wir Nachholbedarf') return ROT_HELL;
  if(a==='Bin mir nicht sicher') return ORANGE_HELL;
  return GRUEN_HELL;
}

function einstufungFarbe(e){
  if(e==='Dringender Handlungsbedarf') return ROT;
  if(e==='Handlungsbedarf erkannt') return ORANGE;
  return GRUEN;
}
function einstufungFill(e){
  if(e==='Dringender Handlungsbedarf') return ROT_HELL;
  if(e==='Handlungsbedarf erkannt') return ORANGE_HELL;
  return GRUEN_HELL;
}

const FRAGEN = [
  "Wissen Sie, welche KI-Tools Ihre Mitarbeiter heute nutzen?",
  "Mitarbeiter gibt Kundenvertrag in ChatGPT ein \u2013 wissen Sie ob das passiert?",
  "Haben Sie Ihren Mitarbeitern klar gesagt was sie mit KI d\u00FCrfen?",
  "Wissen Sie dass Sie als Gesch\u00E4ftsf\u00FChrer pers\u00F6nlich haften?",
  "Ist Ihre Kalkulation oder Marge jemals in ein KI-Tool eingegeben worden?",
  "K\u00F6nnten Sie einem Beh\u00F6rdenvertreter eine KI-Inventarliste vorlegen?",
  "Was passiert in den ersten 24 Stunden nach einem KI-Datenschutzvorfall?",
  "Gibt es eine klare KI-Leitlinie die jeder Mitarbeiter kennt?",
  "Haben Sie KI verboten und glauben damit sicher zu sein?",
  "Was passiert wenn morgen in der Zeitung steht: Mitarbeiter gibt Daten an KI weiter?",
  "Was w\u00FCrde Ihr wichtigster Kunde sagen wenn er w\u00FCsste was mit seinen Daten passiert?",
  "Was w\u00E4re anders wenn Sie morgen eine klare KI-Haltung h\u00E4tten?"
];

function berichtAbsaetze(text){
  if(!text) return [P('',{size:21})];
  var result = [];
  var lines = text.split('\n');
  var i = 0;
  var handlungsfelderGezeigt = false;

  while(i < lines.length){
    var line = lines[i].trim();
    if(!line){ result.push(SP(80)); i++; continue; }

    // Vor erstem HANDLUNGSFELD: Überschrift "Ihre 3 wichtigsten Handlungsfelder" einfügen
    if(line.match(/^HANDLUNGSFELD 1:/) && !handlungsfelderGezeigt){
      handlungsfelderGezeigt = true;
      result.push(SP(160));
      result.push(new Table({width:{size:9026,type:WidthType.DXA},columnWidths:[200,8826],borders:noBorders,
        rows:[new TableRow({children:[
          C([],{fill:GOLD,w:200,mt:0,mb:0}),
          C([P('Ihre 3 wichtigsten Handlungsfelder',{size:20,bold:true,color:DUNKEL,before:80,after:80})],{fill:GOLD_HELL,w:8826,ml:200})
        ]})]
      }));
      result.push(SP(40));
    }

    if(line.match(/^HANDLUNGSFELD \d+:/)){
      var m = line.match(/^(HANDLUNGSFELD \d+:)\s*(.*)/);
      var titelNr = m ? m[1] : 'HANDLUNGSFELD:';
      var titelText = m ? m[2] : line;
      // Titelzeile: dunkler Hintergrund, goldene Nummer, weisser Titeltext
      result.push(new Table({width:{size:9026,type:WidthType.DXA},columnWidths:[9026],borders:noBorders,
        rows:[new TableRow({children:[C([
          new Paragraph({spacing:{before:80,after:80},children:[
            new TextRun({text:titelNr+' ',font:'Arial',size:20,bold:true,color:GOLD}),
            new TextRun({text:titelText,font:'Arial',size:22,bold:true,color:WEISS})
          ]})
        ],{fill:DUNKEL,w:9026,ml:240,mr:240})]})]
      }));
      i++;
      var bodyLines = [];
      while(i < lines.length){
        var nx = lines[i].trim();
        if(nx.match(/^(HANDLUNGSFELD \d+:|BLICK NACH VORNE|WAS NACH UNSEREM GESPR)/)) break;
        if(nx) bodyLines.push(nx);
        i++;
      }
      if(bodyLines.length > 0){
        result.push(new Table({width:{size:9026,type:WidthType.DXA},columnWidths:[9026],borders:noBorders,
          rows:[new TableRow({children:[C(
            bodyLines.map(function(bl){return P(bl,{size:21,color:TEXT,after:80});}),
            {fill:GOLD_HELL,w:9026,ml:240,mr:240,b:{top:noBorder,bottom:thin(GOLD),left:{style:BorderStyle.SINGLE,size:12,color:GOLD},right:noBorder}}
          )]})]
        }));
      }
      result.push(SP(60));
      continue;
    }

    // Seitenumbruch VOR Blick nach vorne
    if(line.match(/^BLICK NACH VORNE/)){
      result.push(new Paragraph({children:[new PageBreak()]}));
      var blickRows = [];
      i++;
      while(i < lines.length){
        var nx = lines[i].trim();
        if(nx.match(/^WAS NACH UNSEREM GESPR/)) break;
        if(nx) blickRows.push(P(nx,{size:21,color:TEXT,after:80}));
        i++;
      }
      result.push(new Table({width:{size:9026,type:WidthType.DXA},columnWidths:[9026],borders:noBorders,
        rows:[new TableRow({children:[C([
          new Paragraph({spacing:{before:80,after:60},children:[new TextRun({text:'Blick nach vorne',font:'Arial',size:22,bold:true,color:DUNKEL})]}),
          ...blickRows
        ],{fill:GRAU,w:9026,b:leftB(GOLD,8),ml:240,mr:240,mt:80,mb:80})]})]
      }));
      result.push(SP(120));
      continue;
    }

    if(line.match(/^WAS NACH UNSEREM GESPR/)){
      var wasRows = [];
      i++;
      while(i < lines.length){
        var nx = lines[i].trim();
        if(nx) wasRows.push(P(nx,{size:21,color:TEXT,after:80}));
        i++;
      }
      result.push(new Table({width:{size:9026,type:WidthType.DXA},columnWidths:[9026],borders:noBorders,
        rows:[new TableRow({children:[C([
          new Paragraph({spacing:{before:80,after:60},children:[new TextRun({text:'Was nach unserem Gespräch anders ist',font:'Arial',size:22,bold:true,color:DUNKEL})]}),
          ...wasRows
        ],{fill:GOLD_HELL,w:9026,b:leftB(GOLD,8),ml:240,mr:240,mt:80,mb:80})]})]
      }));
      continue;
    }

    result.push(P(line,{size:21,color:TEXT,before:0,after:80}));
    i++;
  }
  return result;
}

app.get('/', function(req,res){
  res.json({status:'ok',service:'Schatten-KI Dokument Generator'});
});

app.post('/generate', async function(req,res){
  try {
    const d = req.body;
    const vorname = d.vorname || d.Vorname || '';
    const nachname = d.nachname || d.Nachname || '';
    const email = d.email||'';
    const datum = d.datum||new Date().toLocaleDateString('de-AT');
    const rot = Number(d.nachholbedarf || d.Nachholbedarf || 0) || 0;
    const gelb = Number(d.unsicher || d.Unsicher || 0) || 0;
    const gruen = Number(d.geregelt || d.Geregelt || 0) || 0;
    const gesamt = rot + gelb + gruen;
    const einstufung = d.einstufung||'Handlungsbedarf erkannt';
    const ki_bericht = (d.ki_bericht||'').replace(/"/g, "'");
    const anrede = ki_bericht.match(/Sehr geehrte Frau/) ? 'Sehr geehrte Frau' : 'Sehr geehrter Herr';

    const ec = einstufungFarbe(einstufung);
    const ef = einstufungFill(einstufung);

    // ✅ FIX: Antworten kommen als verschachteltes Objekt von Make.com
    const antworten = d.Antworten || d.antworten || {};

    const antwortenRows = [
      new TableRow({children:[
        C([P('Nr.',{size:17,bold:true,color:WEISS})],{fill:DUNKEL,w:400,b:allB(DUNKEL)}),
        C([P('Frage',{size:17,bold:true,color:WEISS})],{fill:DUNKEL,w:5626,b:allB(DUNKEL)}),
        C([P('Ihre Einsch\u00E4tzung',{size:17,bold:true,color:WEISS,align:AlignmentType.CENTER})],{fill:DUNKEL,w:3000,b:allB(DUNKEL)}),
      ]}),
      ...[1,2,3,4,5,6,7,8,9,10,11,12].map(function(i){
        const key = 'F'+(i<10?'0':'')+i;
        const antwort = antworten[key] || '';  // ✅ FIX: aus antworten-Objekt holen
        const ac = ampelFarbe(antwort);
        const af = ampelFill(antwort);
        return new TableRow({children:[
          C([P(String(i<10?'0'+i:i),{size:18,bold:true,color:TEXT_GRAU,align:AlignmentType.CENTER})],{fill:GRAU,w:400,b:allB(BORDER)}),
          C([P(FRAGEN[i-1],{size:17,color:TEXT})],{fill:WEISS,w:5626,b:allB(BORDER)}),
          C([P(antwort,{size:15,bold:true,color:ac,align:AlignmentType.CENTER})],{fill:af,w:3000,b:allB(BORDER),v:VerticalAlign.CENTER}),
        ]});
      })
    ];

    const doc = new Document({
      styles:{default:{document:{run:{font:'Arial',size:22,color:TEXT}}}},
      sections:[{
        properties:{page:{size:{width:11906,height:16838},margin:{top:1134,right:1134,bottom:1134,left:1134}}},
        headers:{default:new Header({children:[new Paragraph({
          border:{bottom:{style:BorderStyle.SINGLE,size:6,color:DUNKEL,space:1}},
          spacing:{before:0,after:80},
          children:[new TextRun({text:'Rudolf Pusterhofer  \u00B7  Schatten-KI stoppen  \u00B7  Klartext f\u00FCr Gesch\u00E4ftsf\u00FChrer  \u00B7  ausderpraxis.com',font:'Arial',size:16,color:TEXT_GRAU})]
        })]})},
        footers:{default:new Footer({children:[new Paragraph({
          border:{top:{style:BorderStyle.SINGLE,size:3,color:BORDER,space:1}},
          spacing:{before:80,after:0},
          children:[new TextRun({text:'Rudolf Pusterhofer  \u00B7  +43 664 88 366 140  \u00B7  ausderpraxis.com  \u00B7  Vertraulich',font:'Arial',size:16,color:TEXT_GRAU})]
        })]})},
        children:[
          new Table({width:{size:9026,type:WidthType.DXA},columnWidths:[9026],borders:noBorders,
            rows:[new TableRow({children:[C([
              P('Aus der Praxis \u00B7 Klartext f\u00FCr Gesch\u00E4ftsf\u00FChrer',{size:16,color:GOLD,before:160,after:60}),
              new Paragraph({spacing:{before:0,after:80},children:[new TextRun({text:'Schatten-KI Risiko-Check',font:'Arial',size:34,bold:true,color:WEISS})]}),
              P('Pers\u00F6nliche Auswertung \u00B7 '+datum+' \u00B7 ausderpraxis.com',{size:16,color:"AAAAAA",after:160}),
            ],{fill:DUNKEL,w:9026,ml:300,mr:300})]})]
          }),
          SP(180),
          P('Pers\u00F6nlich & Vertraulich',{size:17,italic:true,color:TEXT_GRAU,after:40}),
          new Paragraph({spacing:{before:0,after:60},children:[new TextRun({text:vorname+' '+nachname,font:'Arial',size:26,bold:true,color:DUNKEL})]}),
          P(email,{size:18,color:TEXT_GRAU,after:200}),
          new Table({width:{size:9026,type:WidthType.DXA},columnWidths:[2942,100,2942,100,2942],borders:noBorders,
            rows:[new TableRow({children:[
              C([new Paragraph({alignment:AlignmentType.CENTER,spacing:{before:120,after:40},children:[new TextRun({text:String(rot),font:'Arial',size:72,bold:true,color:ROT})]}),P('Hier haben wir Nachholbedarf',{size:16,color:ROT,align:AlignmentType.CENTER,after:120})],{fill:ROT_HELL,w:2942,b:allB(ROT),mt:0,mb:0,ml:80,mr:80}),
              C([P('')],{w:100}),
              C([new Paragraph({alignment:AlignmentType.CENTER,spacing:{before:120,after:40},children:[new TextRun({text:String(gelb),font:'Arial',size:72,bold:true,color:ORANGE})]}),P('Bin mir nicht sicher',{size:16,color:ORANGE,align:AlignmentType.CENTER,after:120})],{fill:ORANGE_HELL,w:2942,b:allB(ORANGE),mt:0,mb:0,ml:80,mr:80}),
              C([P('')],{w:100}),
              C([new Paragraph({alignment:AlignmentType.CENTER,spacing:{before:120,after:40},children:[new TextRun({text:String(gruen),font:'Arial',size:72,bold:true,color:GRUEN})]}),P('Bei uns geregelt',{size:16,color:GRUEN,align:AlignmentType.CENTER,after:120})],{fill:GRUEN_HELL,w:2942,b:allB(GRUEN),mt:0,mb:0,ml:80,mr:80}),
            ]})]
          }),
          SP(140),
          new Table({width:{size:9026,type:WidthType.DXA},columnWidths:[9026],borders:noBorders,
            rows:[new TableRow({children:[C([
              new Paragraph({spacing:{before:100,after:60},children:[new TextRun({text:einstufung,font:'Arial',size:26,bold:true,color:ec})]}),
              P('Das ist kein theoretisches Risiko. Das ist die Situation in Ihrem Unternehmen \u2013 heute, in diesem Moment.',{size:20,color:TEXT,after:100}),
            ],{fill:ef,w:9026,b:leftB(ec,14),ml:240,mr:240})]})]
          }),
          SP(200),
          new Table({width:{size:9026,type:WidthType.DXA},columnWidths:[200,8826],borders:noBorders,
            rows:[new TableRow({children:[C([],{fill:GOLD,w:200,mt:0,mb:0}),C([P('Ihre 12 Antworten im \u00DCberblick',{size:20,bold:true,color:DUNKEL,before:80,after:80})],{fill:GOLD_HELL,w:8826,ml:200})]})]
          }),
          SP(60),
          new Table({width:{size:9026,type:WidthType.DXA},columnWidths:[400,5626,3000],borders:noBorders,rows:antwortenRows}),
          new Paragraph({children:[new PageBreak()]}),
          new Table({width:{size:9026,type:WidthType.DXA},columnWidths:[200,8826],borders:noBorders,
            rows:[new TableRow({children:[C([],{fill:GOLD,w:200,mt:0,mb:0}),C([P('Ihr pers\u00F6nlicher KI-Bericht',{size:20,bold:true,color:DUNKEL,before:80,after:80})],{fill:GOLD_HELL,w:8826,ml:200})]})]
          }),
          SP(60),
          ...berichtAbsaetze(ki_bericht),
          new Paragraph({children:[new PageBreak()]}),
          SP(200),
          P('Eisenstadt, '+datum,{size:20,color:TEXT_GRAU,after:160}),
          new Paragraph({spacing:{before:0,after:80},children:[new TextRun({text:anrede+' '+nachname+',',font:'Arial',size:23,color:TEXT})]}),
          SP(60),
          P('Ihr Check-Ergebnis liegt mir vor. Ich m\u00F6chte Ihnen anbieten, was ich in 30 Jahren immer als wirksamsten ersten Schritt erlebt habe:',{size:21,color:TEXT,after:100}),
          SP(100),
          new Table({width:{size:9026,type:WidthType.DXA},columnWidths:[9026],borders:noBorders,
            rows:[new TableRow({children:[C([
              new Paragraph({alignment:AlignmentType.CENTER,spacing:{before:140,after:80},children:[new TextRun({text:'Ein Gespr\u00E4ch. 30 Minuten. Kostenlos.',font:'Arial',size:26,bold:true,color:DUNKEL})]}),
              P('Kein Verkaufsgespr\u00E4ch. Keine Verpflichtung. Keine Pr\u00E4sentation.',{size:18,italic:true,color:TEXT_GRAU,align:AlignmentType.CENTER,after:80}),
              P('Ich zeige Ihnen, welche drei Schritte in Ihrem Unternehmen sofort den gr\u00F6\u00DFten Unterschied machen.',{size:20,color:TEXT,align:AlignmentType.CENTER,after:100}),
              new Paragraph({alignment:AlignmentType.CENTER,spacing:{before:80,after:140},children:[new TextRun({text:'calendly.com/pusterhofer',font:'Arial',size:22,bold:true,color:TEXT})]}),
            ],{fill:GOLD_HELL,w:9026,b:{top:thin(GOLD),bottom:thin(GOLD),left:{style:BorderStyle.SINGLE,size:12,color:GOLD},right:thin(GOLD)},ml:300,mr:300})]})]
          }),
          SP(200),
          P('Ich freue mich auf das Gespr\u00E4ch.',{size:21,color:TEXT,after:200}),
          new Paragraph({spacing:{before:0,after:40},children:[new TextRun({text:'Rudolf Pusterhofer',font:'Arial',size:24,bold:true,color:DUNKEL})]}),
          new Paragraph({spacing:{before:0,after:40},children:[new TextRun({text:'Schatten-KI stoppen \u00B7 Klartext f\u00FCr Gesch\u00E4ftsf\u00FChrer',font:'Arial',size:20,color:TEXT_GRAU})]}),
          P('ausderpraxis.com \u00B7 +43 664 88 366 140',{size:18,color:TEXT_GRAU}),
        ]
      }]
    });

    const buf = await Packer.toBuffer(doc);
    const filename = 'Schatten-KI-Auswertung-'+nachname+'-'+vorname+'.docx';
    res.setHeader('Content-Type','application/vnd.openxmlformats-officedocument.wordprocessingml.document');
    res.setHeader('Content-Disposition','attachment; filename="'+filename+'"');
    res.send(buf);
  } catch(e) {
    console.error(e);
    res.status(500).json({error:e.message});
  }
});

const PORT = process.env.PORT || 3000;
app.listen(PORT, function(){
  console.log('Server läuft auf Port '+PORT);
});
