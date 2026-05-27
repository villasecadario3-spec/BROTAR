const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  AlignmentType, WidthType, BorderStyle, ShadingType, VerticalAlign,
  ImageRun, PageOrientation, SectionType, Footer, Header
} = require('docx');
const XLSX = require('xlsx');
const JSZip = require('jszip');
 
const SERVIU_B64 = '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';
 
exports.handler = async (event) => {
  if (event.httpMethod === 'OPTIONS') {
    return { statusCode: 200, headers: corsHeaders() };
  }
  if (event.httpMethod !== 'POST') {
    return { statusCode: 405, body: 'Method Not Allowed' };
  }
 
  try {
    const body = JSON.parse(event.body);
    const { excelBase64, config, batchStart = 0, batchSize = 25 } = body;
 
    if (!excelBase64) return { statusCode: 400, body: JSON.stringify({ error: 'No se recibió el archivo Excel' }) };
 
    // Leer Excel
    const excelBuf = Buffer.from(excelBase64, 'base64');
    const wb = XLSX.read(excelBuf, { type: 'buffer' });
    const ws = wb.Sheets[wb.SheetNames[0]];
    const data = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' });
    const allRows = data.slice(1).map(r => r.map(v => v === null || v === undefined ? '' : String(v).trim()));
    const totalRows = allRows.length;
 
    // Procesar solo el lote solicitado
    const batchEnd = Math.min(batchStart + batchSize, totalRows);
    const rows = allRows.slice(batchStart, batchEnd);
 
    const logoBuffer = Buffer.from(SERVIU_B64, 'base64');
    const zip = new JSZip();
 
    function normStr(s) {
      return s.trim().normalize('NFD').replace(/[\u0300-\u036f]/g, '')
        .replace(/[ñÑ]/g, 'n').replace(/[^a-zA-Z0-9 ]/g, '').trim();
    }
 
    for (let i = 0; i < rows.length; i++) {
      const row = rows[i];
      const folio = batchStart + i + 1;
      const nombreSafe = normStr(val(row[1]));
      const apellidoSafe = normStr(val(row[2]));
      const rutSafe = (val(row[4]) || `sin_rut_${folio}`).replace(/[^a-zA-Z0-9]/g, '');
      const filename = `${folio}.FichaDiagnostica_${nombreSafe}_${apellidoSafe}_${rutSafe}.docx`;
 
      const doc = buildDoc(row, config, folio, logoBuffer);
      const buffer = await Packer.toBuffer(doc);
      zip.file(filename, buffer);
    }
 
    const zipBuffer = await zip.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE', compressionOptions: { level: 4 } });
    const zipBase64 = zipBuffer.toString('base64');
 
    return {
      statusCode: 200,
      headers: { ...corsHeaders(), 'Content-Type': 'application/json' },
      body: JSON.stringify({
        zipBase64,
        batchStart,
        batchEnd,
        totalRows,
        done: batchEnd >= totalRows
      })
    };
 
  } catch (err) {
    console.error(err);
    return {
      statusCode: 500,
      headers: corsHeaders(),
      body: JSON.stringify({ error: err.message })
    };
  }
};
 
function corsHeaders() {
  return {
    'Access-Control-Allow-Origin': '*',
    'Access-Control-Allow-Headers': 'Content-Type',
    'Access-Control-Allow-Methods': 'POST, OPTIONS'
  };
}
 
// ─── Helpers ─────────────────────────────────────────────────────────────────
function val(v) {
  if (!v) return '';
  const s = String(v).trim();
  return (s === 'nan' || s === 'NaN') ? '' : s;
}
function capitalize(s) { if (!s) return ''; return s.charAt(0).toUpperCase() + s.slice(1).toLowerCase(); }
function xFor(expected, actual) { return actual === expected ? 'X' : ''; }
function xOrEmpty(v) { const s = val(v); return (s === 'X' || s === 'x') ? 'X' : ''; }
function siNo(v) { return val(v).toLowerCase() === 'si' || val(v).toLowerCase() === 'sí' ? 'SI' : 'NO'; }
 
// ─── Bordes ───────────────────────────────────────────────────────────────────
const BD = { style: BorderStyle.SINGLE, size: 6, color: '000000' };
const BORDERS = { top: BD, bottom: BD, left: BD, right: BD };
const NB = { style: BorderStyle.NONE, size: 0, color: 'FFFFFF' };
const NO_BORDERS = { top: NB, bottom: NB, left: NB, right: NB };
const MARGIN = 720;
const CONTENT_L = 18720 - MARGIN * 2;
const CONTENT_P = 12240 - MARGIN * 2;
 
function cell(text, opts = {}) {
  const { bold=false, width=1000, borders=BORDERS, shade=null, colspan=1, align=AlignmentType.LEFT, fontSize=18, vAlign=VerticalAlign.CENTER, italic=false } = opts;
  return new TableCell({
    columnSpan: colspan, width: { size: width, type: WidthType.DXA }, borders,
    verticalAlign: vAlign,
    shading: shade ? { fill: shade, type: ShadingType.CLEAR } : undefined,
    margins: { top: 60, bottom: 60, left: 100, right: 100 },
    children: [new Paragraph({ alignment: align, children: [new TextRun({ text: String(text || ''), bold, italic, size: fontSize, font: 'Arial' })] })]
  });
}
function hCell(text, opts = {}) { return cell(text, { bold: true, shade: 'D9D9D9', ...opts }); }
function p(text, opts = {}) {
  const { bold=false, size=20, align=AlignmentType.LEFT, before=80, italic=false } = opts;
  return new Paragraph({ alignment: align, spacing: { before, after: 0 }, children: [new TextRun({ text: String(text||''), bold, italic, size, font: 'Arial' })] });
}
function sp(before = 120) { return new Paragraph({ spacing: { before, after: 0 }, children: [] }); }
 
function makeHeader(logoBuffer) {
  return { default: new Header({ children: [new Paragraph({ children: [new ImageRun({ data: logoBuffer, type: 'png', transformation: { width: 100, height: 82 } })] })] }) };
}
function makeFooter() {
  return { default: new Footer({ children: [new Paragraph({ alignment: AlignmentType.LEFT, children: [new TextRun({ text: 'Formatos 2025 SERVIU REGION OHIGGINS res. Ex N°195 del 05.02.2025', size: 16, font: 'Calibri' })] })] }) };
}
function landscapeProps(type = SectionType.NEXT_PAGE) {
  return { type, page: { size: { width: 12240, height: 18720, orientation: PageOrientation.LANDSCAPE }, margin: { top: MARGIN, right: MARGIN, bottom: MARGIN, left: MARGIN } } };
}
function portraitProps(type = SectionType.NEXT_PAGE) {
  return { type, page: { size: { width: 12240, height: 18720 }, margin: { top: MARGIN, right: MARGIN, bottom: MARGIN, left: MARGIN } } };
}
 
function extractMembers(row) {
  const members = [{
    nombre: val(row[1]), apellido1: val(row[2]), apellido2: val(row[3]),
    rut: val(row[4]), estadoCivil: val(row[5]), genero: val(row[6]),
    parentesco: 'Postulante', edad: val(row[7]), adultoMayor: val(row[8]),
    estudios: val(row[9]), situacionLaboral: val(row[10]),
    puebloOriginario: val(row[13]), nacionalidad: val(row[14]),
    enfermedad: val(row[16]), enfermedadCual: val(row[17]),
    discapacidad: val(row[18]), discapacidadCual: val(row[20]), valech: val(row[22]),
  }];
  [24, 48, 72, 96, 120, 144].forEach(b => {
    const nombre = val(row[b]), ap1 = val(row[b + 1]);
    if (!nombre && !ap1) return;
    members.push({
      nombre, apellido1: ap1, apellido2: val(row[b + 2]),
      rut: val(row[b + 3]), estadoCivil: val(row[b + 4]), genero: val(row[b + 5]),
      parentesco: val(row[b + 6]), edad: val(row[b + 7]), adultoMayor: val(row[b + 8]),
      estudios: val(row[b + 9]), situacionLaboral: val(row[b + 10]),
      puebloOriginario: val(row[b + 13]), nacionalidad: val(row[b + 14]),
      enfermedad: val(row[b + 16]), enfermedadCual: val(row[b + 17]),
      discapacidad: val(row[b + 18]), discapacidadCual: val(row[b + 20]), valech: val(row[b + 22]),
    });
  });
  return members;
}
 
function buildDoc(row, pd, folio, logoBuffer) {
  const members = extractMembers(row);
  const W_L = CONTENT_L;
 
  // ── Encabezado ──
  const logoW = Math.round(W_L * 0.08), tituloW = Math.round(W_L * 0.52), folioW = W_L - logoW - tituloW;
  const fW2 = Math.round(folioW * 0.45);
 
  const headerTable = new Table({
    width: { size: W_L, type: WidthType.DXA }, columnWidths: [logoW, tituloW, folioW],
    rows: [new TableRow({ children: [
      new TableCell({ width: { size: logoW, type: WidthType.DXA }, borders: NO_BORDERS, margins: { top: 0, bottom: 0, left: 0, right: 0 }, children: [new Paragraph({ children: [] })] }),
      new TableCell({ width: { size: tituloW, type: WidthType.DXA }, borders: NO_BORDERS, verticalAlign: VerticalAlign.CENTER, margins: { top: 40, bottom: 40, left: 100, right: 40 }, children: [
        sp(40),
        new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: 'FICHA DIAGNÓSTICA FAMILIA POSTULANTE', bold: true, size: 24, font: 'Arial' })] }),
        new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: 'D.S. N° 49, V. y U., DE 2011', bold: true, size: 20, font: 'Arial' })] }),
      ]}),
      new TableCell({ width: { size: folioW, type: WidthType.DXA }, borders: NO_BORDERS, verticalAlign: VerticalAlign.BOTTOM, margins: { top: 40, bottom: 40, left: 80, right: 80 }, children: [
        new Table({ width: { size: folioW - 160, type: WidthType.DXA }, columnWidths: [fW2, folioW - 160 - fW2],
          rows: [new TableRow({ children: [
            new TableCell({ width: { size: fW2, type: WidthType.DXA }, borders: BORDERS, shading: { fill: 'D9D9D9', type: ShadingType.CLEAR }, margins: { top: 60, bottom: 60, left: 80, right: 80 }, children: [new Paragraph({ children: [new TextRun({ text: 'N° Folio Ficha', bold: true, size: 18, font: 'Arial' })] })] }),
            new TableCell({ width: { size: folioW - 160 - fW2, type: WidthType.DXA }, borders: BORDERS, margins: { top: 60, bottom: 60, left: 80, right: 80 }, children: [new Paragraph({ children: [new TextRun({ text: String(folio), size: 18, font: 'Arial' })] })] }),
          ]})]
        })
      ]})
    ]})]
  });
 
  // ── Sección I ──
  const LEFT_W = 7479, RIGHT_W = W_L - LEFT_W;
  const cws = [2943, 709, 1418, 708, 1701];
  const v4 = cws[1] + cws[2] + cws[3] + cws[4];
 
  function hcL(txt) {
    return new TableCell({ width: { size: cws[0], type: WidthType.DXA }, borders: BORDERS, shading: { fill: 'D9D9D9', type: ShadingType.CLEAR }, margins: { top: 60, bottom: 60, left: 100, right: 100 },
      children: [new Paragraph({ children: [new TextRun({ text: txt, bold: true, size: 18, font: 'Arial' })] })]
    });
  }
  function vcL(txt, w) {
    return new TableCell({ width: { size: w, type: WidthType.DXA }, borders: BORDERS, margins: { top: 60, bottom: 60, left: 100, right: 100 },
      children: [new Paragraph({ children: [new TextRun({ text: String(txt || ''), size: 18, font: 'Arial' })] })]
    });
  }
  function vcL4(txt) {
    return new TableCell({ columnSpan: 4, width: { size: v4, type: WidthType.DXA }, borders: BORDERS, margins: { top: 60, bottom: 60, left: 100, right: 100 },
      children: [new Paragraph({ children: [new TextRun({ text: String(txt || ''), size: 18, font: 'Arial' })] })]
    });
  }
 
  const leftTable = new Table({ width: { size: LEFT_W, type: WidthType.DXA }, columnWidths: cws,
    rows: [
      new TableRow({ children: [
        new TableCell({ width: { size: cws[0], type: WidthType.DXA }, borders: BORDERS, shading: { fill: 'D9D9D9', type: ShadingType.CLEAR }, margins: { top: 60, bottom: 60, left: 100, right: 100 }, children: [
          new Paragraph({ children: [new TextRun({ text: 'Tipología de Proyecto (*)', bold: true, size: 18, font: 'Arial' })] }),
          new Paragraph({ children: [new TextRun({ text: '(encerrar en un círculo,', size: 16, font: 'Arial' })] }),
          new Paragraph({ children: [new TextRun({ text: 'la tipología que corresponda)', size: 16, font: 'Arial' })] }),
        ]}),
        vcL('CNT', cws[1]), vcL('', cws[2]), vcL('MP', cws[3]), vcL('', cws[4]),
      ]}),
      new TableRow({ children: [hcL('Nombre del Proyecto'), vcL4(pd.nombreProyecto)] }),
      new TableRow({ children: [hcL('Nombre del Grupo'), vcL4(pd.nombreGrupo)] }),
      new TableRow({ children: [hcL('N° de Familias Postulantes'), vcL4(pd.nFamilias)] }),
      new TableRow({ children: [hcL('Entidad Patrocinante'), vcL4(pd.entidad)] }),
      new TableRow({ children: [hcL('Región'), vcL4(pd.region)] }),
      new TableRow({ children: [hcL('Provincia'), vcL4(pd.provincia)] }),
      new TableRow({ children: [hcL('Comuna'), vcL4(pd.comuna)] }),
    ]
  });
 
  const rightCell = new TableCell({ width: { size: RIGHT_W, type: WidthType.DXA }, borders: NO_BORDERS, verticalAlign: VerticalAlign.TOP, margins: { top: 60, bottom: 60, left: 300, right: 60 },
    children: [new Table({ width: { size: RIGHT_W - 360, type: WidthType.DXA }, columnWidths: [RIGHT_W - 360],
      rows: [new TableRow({ children: [new TableCell({ width: { size: RIGHT_W - 360, type: WidthType.DXA }, borders: BORDERS, margins: { top: 100, bottom: 100, left: 150, right: 150 },
        children: [new Paragraph({ children: [new TextRun({ text: '(*) CNT: Construcción en Nuevos Terrenos; MP: Megaproyecto;', size: 18, font: 'Arial' })] })]
      })] })]
    })]
  });
 
  const secITable = new Table({ width: { size: W_L, type: WidthType.DXA }, columnWidths: [LEFT_W, RIGHT_W],
    rows: [new TableRow({ children: [
      new TableCell({ width: { size: LEFT_W, type: WidthType.DXA }, borders: NO_BORDERS, margins: { top: 0, bottom: 0, left: 0, right: 0 }, children: [leftTable] }),
      rightCell,
    ]})]
  });
 
  // ── Sección II ──
  const cwBase = [390,1155,1140,990,990,825,1080,1080,585,840,1140,1155,1245,960,630,705,765,615,960];
  const baseSum = cwBase.reduce((a,b)=>a+b,0);
  const ff = W_L / baseSum;
  const adjF = cwBase.map(w => Math.round(w * ff));
  adjF[adjF.length-1] += W_L - adjF.reduce((a,b)=>a+b,0);
 
  function famCell(txt, w, opts = {}) {
    const { bold=false, shade=null, colspan=1 } = opts;
    return new TableCell({ columnSpan: colspan, width: { size: w, type: WidthType.DXA }, borders: BORDERS,
      shading: shade ? { fill: shade, type: ShadingType.CLEAR } : undefined,
      margins: { top: 40, bottom: 40, left: 60, right: 60 },
      children: [new Paragraph({ children: [new TextRun({ text: String(txt||''), bold, size: 16, font: 'Arial' })] })]
    });
  }
  function fhc(txt, w, cs=1) { return famCell(txt, w, { bold: true, shade: 'D9D9D9', colspan: cs }); }
 
  const header1 = new TableRow({ children: [
    fhc('N°',adjF[0]), fhc('Nombre',adjF[1]), fhc('Primer Apellido',adjF[2]), fhc('Segundo Apellido',adjF[3]),
    fhc('RUT',adjF[4]), fhc('Estado civil',adjF[5]), fhc('Género',adjF[6]), fhc('Parentesco',adjF[7]),
    fhc('Edad',adjF[8]), fhc('Ad. Mayor',adjF[9]), fhc('Estudios',adjF[10]), fhc('Sit. Laboral',adjF[11]),
    fhc('Ascend. Indígena', adjF[12]+adjF[13], 2), fhc('Extranjero', adjF[14]+adjF[15], 2),
    fhc('Enf. Cat.', adjF[16]+adjF[17], 2), fhc('Valech', adjF[18]),
  ]});
  const header2 = new TableRow({ children: [
    famCell('',adjF[0]),famCell('',adjF[1]),famCell('',adjF[2]),famCell('',adjF[3]),famCell('',adjF[4]),
    famCell('',adjF[5]),famCell('',adjF[6]),famCell('',adjF[7]),famCell('',adjF[8]),famCell('',adjF[9]),
    famCell('',adjF[10]),famCell('',adjF[11]),
    fhc('CUAL',adjF[12]), fhc('Nac.',adjF[13]),
    fhc('SI/NO',adjF[14]), fhc('CUAL',adjF[15]),
    fhc('SI/NO',adjF[16]), fhc('CUAL',adjF[17]),
    fhc('Valech',adjF[18]),
  ]});
 
  const dataRows = [];
  for (let i = 0; i < 7; i++) {
    const m = members[i] || {};
    const has = !!(m.nombre || m.apellido1);
    const enfSN = has ? (val(m.enfermedad).toLowerCase()==='si'?'SI':'NO') : '';
    const valech = has ? val(m.valech).toUpperCase() : '';
    const esExtranjero = has && val(m.nacionalidad) && !val(m.nacionalidad).toLowerCase().startsWith('chilen') ? 'SI' : (has?'NO':'');
    dataRows.push(new TableRow({ children: [
      famCell(String(i+1),adjF[0]), famCell(capitalize(val(m.nombre)),adjF[1]),
      famCell(capitalize(val(m.apellido1)),adjF[2]), famCell(capitalize(val(m.apellido2)),adjF[3]),
      famCell(val(m.rut),adjF[4]), famCell(val(m.estadoCivil),adjF[5]),
      famCell(val(m.genero),adjF[6]), famCell(val(m.parentesco),adjF[7]),
      famCell(val(m.edad),adjF[8]), famCell(has?val(m.adultoMayor).toUpperCase():'',adjF[9]),
      famCell(val(m.estudios),adjF[10]), famCell(val(m.situacionLaboral),adjF[11]),
      famCell(val(m.puebloOriginario),adjF[12]), famCell(capitalize(val(m.nacionalidad)),adjF[13]),
      famCell(esExtranjero,adjF[14]), famCell(val(m.discapacidadCual),adjF[15]),
      famCell(enfSN,adjF[16]), famCell(val(m.enfermedadCual),adjF[17]),
      famCell(valech,adjF[18]),
    ]}));
  }
 
  const secIITable = new Table({ width: { size: W_L, type: WidthType.DXA }, columnWidths: adjF,
    rows: [header1, header2, ...dataRows]
  });
 
  // ── Sección III ──
  const sitHab=val(row[167]),tipoViv=val(row[173]),sistAgua=val(row[179]),servHig=val(row[184]);
 
  function optRow(label, cols) {
    const lW = Math.round(W_L*0.18), rem = W_L-lW, cW = Math.floor(rem/(cols.length*2));
    const cells = [hCell(label, { width: lW, fontSize: 16 })];
    cols.forEach(c => { cells.push(cell(c.label,{width:cW,fontSize:16})); cells.push(cell(c.x,{width:cW,align:AlignmentType.CENTER,bold:true,fontSize:16})); });
    return new TableRow({ children: cells });
  }
 
  const secIIITable = new Table({ width: { size: W_L, type: WidthType.DXA }, rows: [
    optRow('1.1.- Situación Habitacional:',[{label:'Allegado/a',x:xFor('Allegado/a',sitHab)},{label:'Arrendatario/a',x:xFor('Arrendatario/a',sitHab)},{label:'En Situación de calle',x:xFor('En situación de calle',sitHab)},{label:'En Campamento',x:xFor('En campamento',sitHab)},{label:'Otra:',x:!['Allegado/a','Arrendatario/a','En situación de calle','En campamento'].includes(sitHab)&&sitHab?sitHab:''}]),
    optRow('1.2.- Tipo Vivienda:',[{label:'Departamento',x:xFor('Departamento',tipoViv)},{label:'Casa',x:xFor('Casa',tipoViv)},{label:'Mediagua',x:xFor('Mediagua',tipoViv)},{label:'Choza',x:xFor('Choza',tipoViv)},{label:'Otra:',x:!['Departamento','Casa','Mediagua','Choza'].includes(tipoViv)&&tipoViv?tipoViv:''}]),
    optRow('1.3.- Tipo de Sistema Agua:',[{label:'Con llave dentro',x:xFor('Con llave de agua dentro de la vivienda',sistAgua)},{label:'Con llave fuera',x:xFor('Con llave de agua fuera de vivienda',sistAgua)},{label:'No tiene/acarrea',x:xFor('No tiene o la acarrea',sistAgua)},{label:'Otra:',x:!['Con llave de agua dentro de la vivienda','Con llave de agua fuera de vivienda','No tiene o la acarrea'].includes(sistAgua)&&sistAgua?sistAgua:''}]),
    optRow('1.4.- Tipo Serv. Higiénico:',[{label:'WC alcantarillado',x:xFor('WC conectado al alcantarillado',servHig)},{label:'Letrina/pozo negro',x:xFor('Letrina o pozo negro',servHig)},{label:'No posee',x:xFor('No posee servicio higiénico',servHig)},{label:'Otra:',x:!['WC conectado al alcantarillado','Letrina o pozo negro','No posee servicio higiénico'].includes(servHig)&&servHig?servHig:''}]),
  ]});
 
  const c1d=Math.round(W_L*0.35),c2d=W_L-c1d;
  const distribTable = new Table({ width:{size:W_L,type:WidthType.DXA}, columnWidths:[c1d,c2d], rows:[
    new TableRow({children:[hCell('2.1.- N° de Dormitorios',{width:c1d}),cell(val(row[189]),{width:c2d})]}),
    new TableRow({children:[hCell('2.2.- N° de Camas',{width:c1d}),cell(val(row[190]),{width:c2d})]}),
    new TableRow({children:[hCell('2.3.- Posee Calefont',{width:c1d}),cell(val(row[191]).toLowerCase()==='si'?'SI':'NO',{width:c2d})]}),
  ]});
 
  // ── Sección IV ──
  const cSI=800,cNO=800,cStmt=W_L-cSI-cNO;
  const pregs=[
    {text:'1.- Los problemas los solucionamos conversando y llegando a acuerdos.',r:val(row[192])},
    {text:'2.- La mayor parte de las personas del barrio cumplimos las reglas que acordamos.',r:val(row[195])},
    {text:'3.- Es de mi interés conocer y compartir con las personas del barrio.',r:val(row[198])},
    {text:'4.- Existen varias organizaciones comunitarias en el barrio.',r:val(row[201])},
    {text:'5.- Las personas del barrio siempre están dispuestas a colaborar si alguien tiene problemas.',r:val(row[204])},
  ];
  const secIVTable = new Table({ width:{size:W_L,type:WidthType.DXA}, columnWidths:[cStmt,cSI,cNO], rows:[
    new TableRow({children:[hCell('Afirmación',{width:cStmt}),hCell('SI',{width:cSI,align:AlignmentType.CENTER}),hCell('NO',{width:cNO,align:AlignmentType.CENTER})]}),
    ...pregs.map(pg=>new TableRow({children:[
      cell(pg.text,{width:cStmt,fontSize:18}),
      cell(pg.r.toLowerCase()==='si'||pg.r.toLowerCase()==='sí'?'X':'',{width:cSI,align:AlignmentType.CENTER,bold:true}),
      cell(pg.r.toLowerCase()==='no'?'X':'',{width:cNO,align:AlignmentType.CENTER,bold:true}),
    ]}))
  ]});
 
  // ── Sección V ──
  const ATRIB=[
    {l:'Con espacios suficientes',c:211},{l:'Con buena iluminación',c:212},{l:'Con buena ventilación',c:213},
    {l:'Segura',c:214},{l:'Sólida',c:215},{l:'Una bonita fachada',c:216},{l:'Con jardín',c:217},
    {l:'Cercanía con servicios básicos',c:218},{l:'Un bonito barrio',c:219},{l:'Una vivienda ampliable',c:220},
    {l:'El equipamiento urbano del barrio (plazas, juegos, etc.)',c:221},{l:'Otros: (cuáles?)',c:222},
  ];
  const EQUIP=[
    {l:'Áreas verdes adicionales',c:228},{l:'Plaza Activa (Máquinas de ejercicio)',c:229},
    {l:'Contenedores de basura diferenciado por tipo de desecho',c:230},{l:'Huertos Urbanos',c:231},
    {l:'Multicancha',c:232},{l:'Ciclo vías',c:233},{l:'Estacionamientos de Bicicletas',c:234},
    {l:'Escenario o anfiteatro',c:235},{l:'Estacionamientos de Automóviles',c:236},
    {l:'Juegos infantiles',c:237},{l:'Juegos de mesa',c:238},{l:'Luminarias',c:239},
    {l:'Mesa de ping-pong',c:240},{l:'Sala multiuso',c:241},{l:'Canil para mascotas',c:242},
    {l:'Otros. ¿Cuál/es?',c:243},
  ];
  const EQUSR={228:244,229:245,230:246,231:247,232:248,233:249,234:250,235:251,236:252,237:253,238:254,239:255,240:256,241:257,242:260,243:261};
 
  const cX=1200,cL=W_L-cX;
  const atribTable = new Table({width:{size:W_L,type:WidthType.DXA},columnWidths:[cL,cX],rows:[
    new TableRow({children:[hCell('Atributos de la vivienda y el barrio',{width:cL}),hCell('X',{width:cX,align:AlignmentType.CENTER})]}),
    ...ATRIB.map(a=>new TableRow({children:[cell(a.l,{width:cL,fontSize:18}),cell(xOrEmpty(val(row[a.c])),{width:cX,align:AlignmentType.CENTER,bold:true})]}))
  ]});
 
  const equipTable = new Table({width:{size:W_L,type:WidthType.DXA},columnWidths:[cL,cX],rows:[
    new TableRow({children:[hCell('Equipamiento comunitario',{width:cL}),hCell('X',{width:cX,align:AlignmentType.CENTER})]}),
    ...EQUIP.map(e=>new TableRow({children:[cell(e.l,{width:cL,fontSize:18}),cell(xOrEmpty(val(row[e.c])),{width:cX,align:AlignmentType.CENTER,bold:true})]}))
  ]});
 
  function parseUsr(s){if(!s)return{n:false,j:false,m:false,h:false,a:false};const lo=s.toLowerCase();return{n:lo.includes('niñ')||lo.includes('nin'),j:lo.includes('jóven')||lo.includes('joven'),m:lo.includes('mujer'),h:lo.includes('hombre'),a:lo.includes('mayor')};}
  const cE=Math.round(W_L*0.30),cU=Math.floor((W_L-cE)/5),adjU=[cE,cU,cU,cU,cU,W_L-cE-cU*4];
  const allEquip=[...EQUIP,{l:'Bibliotecas',c:null},{l:'Infocentro',c:null},{l:'Otros',c:null}];
  const usrTable = new Table({width:{size:W_L,type:WidthType.DXA},columnWidths:adjU,rows:[
    new TableRow({children:[hCell('Equipamiento',{width:adjU[0],fontSize:16}),hCell('Niñas/niños',{width:adjU[1],align:AlignmentType.CENTER,fontSize:16}),hCell('Jóvenes',{width:adjU[2],align:AlignmentType.CENTER,fontSize:16}),hCell('Mujeres',{width:adjU[3],align:AlignmentType.CENTER,fontSize:16}),hCell('Hombres',{width:adjU[4],align:AlignmentType.CENTER,fontSize:16}),hCell('Ad. Mayores',{width:adjU[5],align:AlignmentType.CENTER,fontSize:16})]}),
    ...allEquip.map(e=>{const usrRaw=e.c&&EQUSR[e.c]?val(row[EQUSR[e.c]]):'';const u=parseUsr(usrRaw);return new TableRow({children:[cell(e.l,{width:adjU[0],fontSize:16}),cell(u.n?'X':'',{width:adjU[1],align:AlignmentType.CENTER,bold:true}),cell(u.j?'X':'',{width:adjU[2],align:AlignmentType.CENTER,bold:true}),cell(u.m?'X':'',{width:adjU[3],align:AlignmentType.CENTER,bold:true}),cell(u.h?'X':'',{width:adjU[4],align:AlignmentType.CENTER,bold:true}),cell(u.a?'X':'',{width:adjU[5],align:AlignmentType.CENTER,bold:true})]});})
  ]});
 
  // ── Sección VI ──
  const TEMAS=[
    {a:'Vivienda y Barrio',t:'Ampliación de la vivienda',c:266},{a:'',t:'Cuidado del medio ambiente y reciclaje',c:267},
    {a:'',t:'Cuidado y mantención de la vivienda',c:268},{a:'',t:'Mejoramiento de la vivienda y el entorno',c:269},
    {a:'Seguridad emergencias',t:'Recomendaciones para emergencias y catástrofes naturales',c:270},
    {a:'',t:'Recomendaciones para la seguridad en el hogar',c:271},
    {a:'Seguridad Ciudadana',t:'Organización vecinal para la seguridad en el barrio',c:272},
    {a:'',t:'Red Institucional que aborda temas de seguridad',c:273},
    {a:'Salud',t:'Alcoholismo y/o drogadicción',c:274},{a:'',t:'Alimentación saludable',c:275},
    {a:'',t:'Red asistencial de salud en el territorio',c:276},{a:'',t:'Salud mental',c:277},
    {a:'',t:'Sexualidad y salud reproductiva',c:278},{a:'',t:'Tipos de previsión en salud (FONASA – ISAPRE)',c:279},
    {a:'Educación y Cultura',t:'Formación y Capacitación Laboral',c:280},{a:'',t:'Nivelación de Estudios',c:281},
    {a:'',t:'Oferta cultural existente en el territorio',c:282},{a:'',t:'Organización vecinal para el desarrollo cultural',c:283},
    {a:'',t:'Pueblos originarios',c:284},{a:'',t:'Personas Inmigrantes',c:285},
    {a:'Familia',t:'Divorcio y separaciones de hecho',c:286},{a:'',t:'Ley de herencia',c:287},
    {a:'',t:'Patrimonio reservado de la mujer casada (art. 150 código civil).',c:288},
    {a:'',t:'Pensión alimenticia',c:289},{a:'',t:'Violencia intrafamiliar',c:290},
    {a:'Participación Comunitaria',t:'Capacitación en elaboración de proyectos',c:291},
    {a:'',t:'Convivencia vecinal',c:292},{a:'',t:'Ley de copropiedad y comité de administración',c:293},
    {a:'',t:'Roles y funciones directiva y tipos de liderazgo',c:294},{a:'',t:'Tipos de organización y su legislación',c:295},
    {a:'Adulto Mayor',t:'Características de esta etapa de vida',c:296},
    {a:'',t:'Recomendaciones para un envejecimiento activo',c:297},
    {a:'',t:'Recomendaciones para familias que asumen cuidados de adulto mayor no Valente',c:298},
    {a:'',t:'Red social especializada en Adulto Mayor',c:299},
    {a:'Mujeres',t:'Fortalecimiento de liderazgo',c:300},{a:'',t:'Violencia hacia la mujer',c:301},
    {a:'Otros',t:'',c:302},
  ];
  const cA=Math.round(W_L*0.22),cT=W_L-cA-cX;
  const temasTable = new Table({width:{size:W_L,type:WidthType.DXA},columnWidths:[cA,cT,cX],rows:[
    new TableRow({children:[hCell('Áreas',{width:cA}),hCell('Temas',{width:cT}),hCell('X',{width:cX,align:AlignmentType.CENTER})]}),
    ...TEMAS.map(t=>new TableRow({children:[cell(t.a,{width:cA,fontSize:16,bold:!!t.a}),cell(t.t,{width:cT,fontSize:16}),cell(xOrEmpty(val(row[t.c])),{width:cX,align:AlignmentType.CENTER,bold:true})]}))
  ]});
 
  // ── Secciones ──
  const notas = [
    p('(*) Estado civil (Casado, Viudo, Soltero, Divorciado, igualmente marcar la condición de separado de hecho, si corresp.)',{size:16}),
    p('(*) Parentesco con Postulante (Hijo/a; cónyuge; conviviente; padre; madre; tío/a; sobrino/a; nieto/a; otros)',{size:16}),
    p('(*) Edad (N° de años; para niños menores de 1 año, indicar N° de meses)',{size:16}),
    p('(*) Situación Laboral (Dependiente-Independiente-Cesante-Jubilado o Pensionado-Trabajo dentro del Hogar-Estudiante)',{size:16}),
    p('(*) Ascendencia indígena (SI/NO). Al marcar SI, identificar Pueblo Indígena.',{size:16}),
    p('(*) Persona Extranjera (SI/NO). Al marcar SI, identificar nacionalidad.',{size:16}),
    p('(*) Discapacidad: mencionar cual, según credencial de discapacidad.',{size:16}),
  ];
 
  const firmas = [
    sp(300), p('NOMBRE Y FIRMA PERSONA POSTULANTE',{bold:true}), sp(700),
    new Paragraph({border:{bottom:{style:BorderStyle.SINGLE,size:6,color:'000000',space:1}},children:[]}),
    sp(40), p('RUT:',{bold:true}), sp(500),
    p('NOMBRE Y FIRMA PROFESIONAL RESPONSABLE PAS – E. P.',{bold:true}), sp(700),
    new Paragraph({border:{bottom:{style:BorderStyle.SINGLE,size:6,color:'000000',space:1}},children:[]}),
    sp(40), p('RUT:',{bold:true}), sp(200), p('TIMBRE',{bold:true}),
  ];
 
  const hf = { headers: makeHeader(logoBuffer), footers: makeFooter() };
 
  return new Document({
    styles: { default: { document: { run: { font: 'Arial', size: 20 } } } },
    sections: [
      { properties: landscapeProps(SectionType.NEXT_PAGE), ...hf, children: [
        headerTable, sp(100),
        p('I.- IDENTIFICACIÓN COMITÉ POSTULANTE Y PROYECTO',{bold:true}), sp(60),
        secITable, sp(120),
        p('II.- IDENTIFICACIÓN GRUPO FAMILIAR POSTULANTE: Información sobre el/la postulante y su grupo familiar.',{bold:true}), sp(60),
        secIITable, sp(40), ...notas, sp(120),
        p('III.- SITUACIÓN HABITACIONAL Y DEL BARRIO ACTUAL',{bold:true}), sp(60),
        p('1.- Situación Habitacional Actual (Marque con una X)',{bold:true}), sp(60),
        secIIITable, sp(100),
        p('2.- Distribución de la vivienda actual',{bold:true}), sp(60),
        distribTable, sp(120),
        p('IV.- Convivencia Comunitaria en su actual lugar de residencia.',{bold:true}),
        p('Favor conteste qué tan de acuerdo se encuentra usted con las siguientes afirmaciones (marcar con un X):',{italic:true,size:18}), sp(60),
        secIVTable,
      ]},
      { properties: landscapeProps(SectionType.NEXT_PAGE), ...hf, children: [
        p('V.- DE LA VIVIENDA Y DEL EQUIPAMIENTO COMUNITARIO DEL PROYECTO HABITACIONAL',{bold:true}), sp(80),
        p('1.- Respecto a la futura Vivienda y Barrio, indique con una X los 3 atributos más valorados:',{bold:true}), sp(60),
        atribTable, sp(100),
        p('2.- ¿Qué Equipamiento Comunitario Adicional le gustaría que fuera parte del Proyecto Habitacional? Marque con una X sus 4 preferencias:',{bold:true}), sp(60),
        equipTable, sp(100),
        p('2.1.- Según su opinión, ¿quiénes serían las personas que más ocuparían ese equipamiento?',{bold:true}), sp(60),
        usrTable,
      ]},
      { properties: landscapeProps(SectionType.NEXT_PAGE), ...hf, children: [
        p('VI.- NECESIDADES E INTERESES DEL GRUPO POSTULANTE',{bold:true}),
        p('Queremos conocer sus intereses. Favor elija un máximo de 3 temas en total:',{italic:true,size:18}), sp(80),
        temasTable,
      ]},
      { properties: portraitProps(SectionType.NEXT_PAGE), ...hf, children: firmas },
    ]
  });
}
