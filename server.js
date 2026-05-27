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

function HF(titel, text){
  return [
    new Paragraph({spacing:{before:180,after:60},children:[new TextRun({text:titel,font:'Arial',size:24,bold:true,color:DUNKEL})]}),
    new Paragraph({spacing:{before:0,after:100},children:[new TextRun({text:text,font:'Arial',size:21,color:TEXT})]})
  ];
}

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
  return text.split('\n').map(function(line){
    line=line.trim();
    if(!line) return SP(80);
    if(line.match(/^HANDLUNGSFELD \d+:/)){
      return new Paragraph({spacing:{before:180,after:60},children:[new TextRun({text:line,font:'Arial',size:24,bold:true,color:DUNKEL})]});
    }
    if(line.match(/^(BLICK NACH VORNE|WAS NACH UNSEREM GESPR)/)){
      return new Paragraph({spacing:{before:180,after:60},children:[new TextRun({text:line,font:'Arial',size:22,bold:true,color:DUNKEL})]});
    }
    return P(line,{size:21,color:TEXT,before:0,after:80});
  });
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
    const rot = parseInt(d.nachholbedarf || d.Nachholbedarf || 0);
    const gelb = parseInt(d.unsicher || d.Unsachen || 0);
    const gruen = parseInt(d.geregelt || d.Geregelt || 0);
    const einstufung = d.einstufung||'Handlungsbedarf erkannt';
    const ki_bericht = (d.ki_bericht||'').replace(/"/g, "'").replace(/\n/g, ' ');
    const antworten = d.antworten||{};

    const ec = einstufungFarbe(einstufung);
    const ef = einstufungFill(einstufung);

    const antwortenRows = [
      new TableRow({children:[
        C([P('Nr.',{size:17,bold:true,color:WEISS})],{fill:DUNKEL,w:400,b:allB(DUNKEL)}),
        C([P('Frage',{size:17,bold:true,color:WEISS})],{fill:DUNKEL,w:5626,b:allB(DUNKEL)}),
        C([P('Ihre Einsch\u00E4tzung',{size:17,bold:true,color:WEISS,align:AlignmentType.CENTER})],{fill:DUNKEL,w:3000,b:allB(DUNKEL)}),
      ]}),
      ...[1,2,3,4,5,6,7,8,9,10,11,12].map(function(i){
        const key = 'F'+(i<10?'0':'')+i;
       const antworten = {
  F01: d.F01 || '',
  F02: d.F02 || '',
  F03: d.F03 || '',
  F04: d.F04 || '',
  F05: d.F05 || '',
  F06: d.F06 || '',
  F07: d.F07 || '',
  F08: d.F08 || '',
  F09: d.F09 || '',
  F10: d.F10 || '',
  F11: d.F11 || '',
  F12: d.F12 || ''
};
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
          new Table({width:{size:9026,type:WidthType.DXA},columnWidths:[9026],borders:noBorders,
            rows:[new TableRow({children:[C(berichtAbsaetze(ki_bericht),{fill:GRAU,w:9026,b:leftB(GOLD,8),ml:240,mr:240,mt:100,mb:100})]})]
          }),
          new Paragraph({children:[new PageBreak()]}),
          SP(200),
          P('Eisenstadt, '+datum,{size:20,color:TEXT_GRAU,after:160}),
          new Paragraph({spacing:{before:0,after:80},children:[new TextRun({text:'Sehr geehrte/r Herr/Frau '+nachname+',',font:'Arial',size:23,color:TEXT})]}),
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
