const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  Header, Footer, AlignmentType, BorderStyle, WidthType,
  ShadingType, VerticalAlign, PageNumber, PageBreak
} = require('docx');
const fs = require('fs');

const PAGE_W = 12240, PAGE_H = 15840, MARGIN = 864;
const CONTENT_W = PAGE_W - 2 * MARGIN; // 10512

// Columns: Line Item | Y/Y $ | Y/Y % | Q1 2026 | Q4 2025 | Q1 2025
const COLS = [3600, 1100, 1000, 1604, 1604, 1604]; // sum = 10512

const NAVY="1A1A2E", BLUE="1F4E79", ACCENT="2E75B6", L_BLUE="D6E4F0";
const WHITE="FFFFFF", GRAY="F5F7FA", MUTED="888888", TEXT="1A1A1A";
const GREEN_B="00B050", RED_B="C00000";

function fill(c){return{fill:c,type:ShadingType.CLEAR};}
function thin(color="CCCCCC"){const b={style:BorderStyle.SINGLE,size:1,color};return{top:b,bottom:b,left:b,right:b};}
function noH(){const n={style:BorderStyle.NIL,size:0,color:"FFFFFF"};const v={style:BorderStyle.SINGLE,size:1,color:"CCCCCC"};return{top:n,bottom:n,left:v,right:v};}

function txt(text,opts={}){
  return new TextRun({text:String(text),font:"Arial",size:opts.size||18,
    bold:opts.bold||false,italics:opts.italic||false,color:opts.color||TEXT,
    underline:opts.underline});
}
function para(children,opts={}){
  return new Paragraph({
    children:Array.isArray(children)?children:[children],
    spacing:{before:opts.before||0,after:opts.after||0,line:240},
    alignment:opts.align||AlignmentType.LEFT,
    shading:opts.shading,indent:opts.indent,border:opts.border,
  });
}
function mkCell(content,bg,width,opts={}){
  const runs=typeof content==='string'?[txt(content,opts)]:content;
  return new TableCell({
    children:[new Paragraph({children:runs,alignment:opts.align||AlignmentType.LEFT,spacing:{before:0,after:0},indent:opts.indent})],
    width:{size:width,type:WidthType.DXA},
    shading:bg?fill(bg):undefined,
    borders:opts.fullBorder?thin():noH(),
    margins:{top:opts.tight?60:90,bottom:opts.tight?60:90,left:120,right:100},
    verticalAlign:VerticalAlign.CENTER,
    columnSpan:opts.span||1,
  });
}
function mkRow(cells,isHeader=false){return new TableRow({children:cells,tableHeader:isHeader});}
function mkTable(rows){
  const n={style:BorderStyle.NIL,size:0,color:"FFFFFF"};
  const o={style:BorderStyle.SINGLE,size:4,color:"CCCCCC"};
  return new Table({rows,width:{size:CONTENT_W,type:WidthType.DXA},columnWidths:COLS,
    borders:{top:o,bottom:o,left:o,right:o,insideH:n,insideV:o}});
}

// Color by sign
function colorSign(val){
  if(!val||val===""||val==="N/A"||val==="—")return[null,TEXT];
  const v=String(val).trim();
  if(v.startsWith("+"))return[GREEN_B,WHITE];
  if(v.startsWith("−")||v.startsWith("-"))return[RED_B,WHITE];
  return[null,TEXT];
}

// Row types
const T_NORMAL   = "normal";
const T_SUBTOTAL = "subtotal";
const T_TOTAL    = "total";
const T_MARGIN   = "margin";
const T_SECTION  = "section";
const T_SPACER   = "spacer";

// DATA: [label, yy$, yy%, q1_2026, q4_2025, q1_2025, type, indent, yycolor]
const ROWS = [
  ["Net revenue",                                      "+$2,815M", "+37.9%", "$10,253", "$10,270", "$7,438",  T_TOTAL,   false, "green"],
  ["Cost of sales",                                    "+$1,125M", "+32.6%", "$4,576",  "$4,433",  "$3,451",  T_NORMAL,  false, "green"],
  ["Amortization of acquisition-related intangibles",  "",         "",       "$261",    "$260",    "$251",    T_NORMAL,  true,  null  ],
  ["Total cost of sales",                              "+$1,135M", "+30.7%", "$4,837",  "$4,693",  "$3,702",  T_SUBTOTAL,false, "green"],
  ["Gross profit",                                     "+$1,680M", "+45.0%", "$5,416",  "$5,577",  "$3,736",  T_SUBTOTAL,false, "green"],
  ["Gross margin",                                     "+300bps",  "+3.0pp","53%",     "54%",     "50%",     T_SUBTOTAL,true,  "green"],
  ["Research and development",                         "+$669M",   "+38.7%", "$2,397",  "$2,330",  "$1,728",  T_NORMAL,  false, "red"  ],
  ["Marketing, general and administrative",            "+$367M",   "+41.4%", "$1,253",  "$1,198",  "$886",    T_NORMAL,  false, "red"  ],
  ["Amortization of acquisition-related intangibles",  "",         "",       "$290",    "$297",    "$316",    T_NORMAL,  true,  null  ],
  ["Total operating expenses",                         "+$1,010M", "+34.5%", "$3,940",  "$3,825",  "$2,930",  T_SUBTOTAL,false, "red"  ],
  ["Operating income",                                 "+$670M",   "+83.1%", "$1,476",  "$1,752",  "$806",    T_TOTAL,   false, "green"],
  ["Interest expense",                                 "",         "",       "($37)",   "($36)",   "($20)",   T_NORMAL,  false, null  ],
  ["Other income (expense), net",                      "",         "",       "$165",    "$358",    "$39",     T_NORMAL,  false, null  ],
  ["Income before taxes & equity income",              "",         "",       "$1,604",  "$2,074",  "$825",    T_SUBTOTAL,true,  null  ],
  ["Income tax provision",                             "",         "",       "$238",    "$455",    "$123",    T_NORMAL,  false, null  ],
  ["Equity income in investee",                        "",         "",       "$6",      "$1",      "$7",      T_NORMAL,  false, null  ],
  ["Income from continuing operations, net of tax",    "",         "",       "$1,372",  "$1,620",  "$709",    T_SUBTOTAL,false, null  ],
  ["Income (loss) from discontinued operations",       "",         "",       "$11",     "($109)",  "—",       T_NORMAL,  false, null  ],
  ["Net income",                                       "+$674M",   "+95.1%", "$1,383",  "$1,511",  "$709",    T_TOTAL,   false, "green"],
  ["", "", "", "", "", "", T_SPACER, false, null],
  ["EARNINGS (LOSS) PER SHARE",                        "",         "",       "",        "",        "",        T_SECTION, false, null  ],
  ["Basic — continuing operations",                    "",         "",       "$0.84",   "$1.00",   "$0.44",   T_NORMAL,  true,  null  ],
  ["Basic — discontinued operations",                  "",         "",       "$0.01",   "($0.07)", "—",       T_NORMAL,  true,  null  ],
  ["Basic earnings per share",                         "",         "",       "$0.85",   "$0.93",   "$0.44",   T_SUBTOTAL,false, null  ],
  ["Diluted — continuing operations",                  "",         "",       "$0.83",   "$0.99",   "$0.44",   T_NORMAL,  true,  null  ],
  ["Diluted — discontinued operations",                "",         "",       "$0.01",   "($0.07)", "—",       T_NORMAL,  true,  null  ],
  ["Diluted earnings per share",                       "+$0.40",   "+90.9%", "$0.84",   "$0.92",   "$0.44",   T_TOTAL,   false, "green"],
  ["", "", "", "", "", "", T_SPACER, false, null],
  ["SHARES USED IN PER SHARE CALCULATION (millions)",  "",         "",       "",        "",        "",        T_SECTION, false, null  ],
  ["Basic",                                            "",         "",       "1,631",   "1,630",   "1,620",   T_NORMAL,  true,  null  ],
  ["Diluted",                                          "",         "",       "1,650",   "1,649",   "1,626",   T_NORMAL,  true,  null  ],
];

// Header row
function hdrRow() {
  const labels = [
    ["Line Item",            AlignmentType.LEFT  ],
    ["Y/Y $ Change",         AlignmentType.CENTER],
    ["Y/Y %",                AlignmentType.CENTER],
    ["Q1 2026\nMar 28, 2026",AlignmentType.CENTER],
    ["Q4 2025\nDec 27, 2025",AlignmentType.CENTER],
    ["Q1 2025\nMar 29, 2025",AlignmentType.CENTER],
  ];
  return mkRow(labels.map(([h,a],i)=>
    mkCell(h, NAVY, COLS[i], {size:13,bold:true,color:WHITE,align:a,fullBorder:true,tight:true})
  ), true);
}

let altIdx = 0;
function dataRow([label, yyd, yyp, q1, q4, q1y, type, indent, yycolor]) {
  if (type === T_SPACER) {
    return mkRow([mkCell("", WHITE, CONTENT_W, {span:6, tight:true})]);
  }
  if (type === T_SECTION) {
    return mkRow([new TableCell({
      children:[new Paragraph({children:[txt(label,{size:13,bold:true,color:WHITE})],spacing:{before:0,after:0},indent:{left:60}})],
      width:{size:CONTENT_W,type:WidthType.DXA},
      shading:fill(BLUE), borders:thin(),
      margins:{top:70,bottom:70,left:120,right:100},
      columnSpan:6,
    })]);
  }
  const isTotal    = type === T_TOTAL;
  const isSubtotal = type === T_SUBTOTAL;
  const bg = isTotal ? L_BLUE : isSubtotal ? "EBF5FB" : altIdx % 2 === 0 ? WHITE : GRAY;
  altIdx++;
  const bold   = isTotal || isSubtotal;
  const italic = type === T_MARGIN;
  const indentDXA = indent ? 280 : 0;

  function yydCell(val, w) {
    if (!yycolor || val==="") return mkCell(val, bg, w, {size:13,bold,color:TEXT,align:AlignmentType.CENTER});
    const cellBg = yycolor==="green" ? GREEN_B : RED_B;
    return mkCell(val, cellBg, w, {size:13,bold,color:WHITE,align:AlignmentType.CENTER});
  }
  function yypCell(val, w) {
    if (!yycolor || val==="") return mkCell(val, bg, w, {size:13,bold,color:TEXT,align:AlignmentType.CENTER});
    const cellBg = yycolor==="green" ? GREEN_B : RED_B;
    return mkCell(val, cellBg, w, {size:13,bold,color:WHITE,align:AlignmentType.CENTER});
  }
  function q1Cell(val, w) {
    if (yycolor && val !== "" && val !== "—") {
      const cellBg = yycolor==="green" ? GREEN_B : RED_B;
      return mkCell(val, cellBg, w, {size:13,bold,color:WHITE,align:AlignmentType.CENTER});
    }
    return mkCell(val, bg, w, {size:13,bold,color:TEXT,align:AlignmentType.CENTER});
  }

  return mkRow([
    new TableCell({
      children:[new Paragraph({
        children:[txt(label,{size:13,bold,italic,color:TEXT})],
        spacing:{before:0,after:0}, indent:{left:indentDXA},
      })],
      width:{size:COLS[0],type:WidthType.DXA},
      shading:fill(bg), borders:noH(),
      margins:{top:90,bottom:90,left:120,right:100},
      verticalAlign:VerticalAlign.CENTER,
    }),
    yyd===""?mkCell("",bg,COLS[1],{size:13,align:AlignmentType.CENTER}):yydCell(yyd,COLS[1]),
    yyp===""?mkCell("",bg,COLS[2],{size:13,align:AlignmentType.CENTER}):yypCell(yyp,COLS[2]),
    q1Cell(q1, COLS[3]),
    mkCell(q4,  bg, COLS[4], {size:13,bold,align:AlignmentType.CENTER}),
    mkCell(q1y, bg, COLS[5], {size:13,bold,align:AlignmentType.CENTER}),
  ]);
}

const children = [];

// Cover
children.push(new Table({
  rows:[
    mkRow([mkCell([txt("ADVANCED MICRO DEVICES, INC. (AMD)",{size:28,bold:true,color:WHITE})],
           NAVY,CONTENT_W,{align:AlignmentType.CENTER,fullBorder:true})]),
    mkRow([mkCell([txt("SIMPLIFIED INCOME STATEMENT ANALYSIS",{size:20,bold:true,color:"A8C4D8"})],
           BLUE,CONTENT_W,{align:AlignmentType.CENTER,fullBorder:true})]),
    mkRow([mkCell([txt("Q1 2026  |  Quarter Ended March 28, 2026  |  Reported May 5, 2026",{size:16,color:"A8C4D8"})],
           ACCENT,CONTENT_W,{align:AlignmentType.CENTER,fullBorder:true})]),
  ],
  width:{size:CONTENT_W,type:WidthType.DXA},columnWidths:[CONTENT_W],
}));

children.push(para("",{before:120}));
children.push(para([txt("CONDENSED CONSOLIDATED STATEMENTS OF OPERATIONS",{size:17,bold:true,color:NAVY})],{align:AlignmentType.CENTER,before:60,after:20}));
children.push(para([txt("Millions of dollars except per share amounts  |  Unaudited  |  Y/Y = Q1 2026 vs Q1 2025",{size:14,italic:true,color:MUTED})],{align:AlignmentType.CENTER,before:0,after:100}));
children.push(mkTable([hdrRow(), ...ROWS.map(dataRow)]));
children.push(para([txt("Y/Y columns shown only for: Net Revenue, Cost of Sales, Total Cost of Sales, Gross Profit, Gross Margin, R&D, MG&A, Total OpEx, Operating Income, Net Income, Diluted EPS",{size:13,italic:true,color:MUTED})],{align:AlignmentType.CENTER,before:80,after:60}));
children.push(para([txt("Source: AMD Q1 2026 Earnings Release, May 5, 2026  |  GAAP figures only",{size:13,italic:true,color:MUTED})],{align:AlignmentType.CENTER,before:0}));

function analysisHdr(text, color) {
  return new Paragraph({children:[txt(text,{size:24,bold:true,color:WHITE})],shading:fill(color),spacing:{before:160,after:0},indent:{left:120}});
}
function analysisBullet(label, text) {
  return new Paragraph({children:[txt("• ",{size:22,color:TEXT}),txt(label,{size:22,bold:true,color:TEXT}),txt(text,{size:22,color:TEXT})],shading:fill(GRAY),spacing:{before:100,after:0,line:320},indent:{left:240,right:120}});
}
function analysisPara(text) {
  return new Paragraph({children:[txt(text,{size:22,color:TEXT})],shading:fill(GRAY),spacing:{before:100,after:0,line:320},indent:{left:120,right:120}});
}

children.push(analysisHdr("QUARTERLY ANALYSIS — Q1 2026 vs Q1 2025", NAVY));
children.push(analysisHdr("Strengths", GREEN_B));
children.push(analysisBullet("Net Revenue", " grew +37.9% Y/Y to $10.25B — strong and broad-based, flat Q/Q despite typical Q1 seasonal weakness, confirming underlying demand strength."));
children.push(analysisBullet("Cost of Sales", " grew +32.6% Y/Y — j�aningfully slower than revenue growth of +37.9%. Data Center GPU revenue, which carries higher margins, is becoming a larger share of the mix, pulling COGS down as a percentage of revenue."));
children.push(analysisBullet("Total Cost of Sales", " grew +30.7% Y/Y — the slowest growing cost line. Revenue grew 7 percentage points faster than total cost of sales, driving Gross Margin expansion from 50% to 53%."));
children.push(analysisBullet("Gross Profit", " grew +45.0% on +37.9% revenue — a 700bps spread that directly reflects cost discipline and product mix improvement."));
children.push(analysisBullet("Gross Margin", " expanded +300bps Y/Y from 50% to 53% — the clearest single indicator of portfolio improvement toward higher-ASP Data Center products."));
children.push(analysisBullet("Operating Income", " grew +83.1% Y/Y — more than double the rate of revenue growth."));
children.push(analysisBullet("Net Income", " grew +95.1% Y/Y — nearly doubling on a revenue base that grew 37.9%."));
children.push(analysisBullet("Diluted EPS", " grew +90.9% Y/Y from $0.44 to $0.84 — shareholders seeing the full benefit of operating leverage."));

children.push(analysisHdr("Concerns", RED_B));
children.push(analysisBullet("R&D", " grew +38.7% — essentially matching revenue growth. No operating leverage; needs to grow slower than revenue over time."));
children.push(analysisBullet("MG&A", " grew +41.4% — faster than revenue at +37.9%. Primary drag on AMD's path to 30% non-GAAP operating margin target."));
children.push(analysisBullet("Total Operating Expenses", " grew +34.5% — slightly below revenue growth in aggregate, but MG&A discipline remains the key watchpoint."));

children.push(analysisHdr("Overall Assessment", BLUE));
children.push(analysisPara("This is a strong quarter. Gross Margin is expanding, Operating Income is nearly doubling, and EPS is nearly doubling on revenue growth of 38%. MG&A growing above the revenue line is the one item to watch."));
children.push(analysisPara("If AMD can hold R&D flat as a percentage of revenue and bring MG&A growth below revenue growth, the path to 30%+ non-GAAP operating margins is credible."));
children.push(new Paragraph({children:[txt("")],shading:fill(GRAY),spacing:{before:80,after:0}}));

// Grade
children.push(new Paragraph({children:[txt("GRADE:  A-",{size:72,bold:true,color:TEXT})],shading:fill(WHITE),spacing:{before:200,after:200},alignment:AlignmentType.CENTER,border:{top:{style:BorderStyle.SINGLE,size:12,color:"1A1A1A"},bottom:{style:BorderStyle.SINGLE,size:12,color:"1A1A1A"}}}));
children.push(analysisHdr("Rationale", BLUE));
children.push(analysisPara("Gross Margin expansion +300bps Y/Y, Operating Income +83.1%, Net Income +95.1% on 37.9% revenue growth — textbook operating leverage. MG&A at +41.4% vs revenue at +37.9% is the one line moving in the wrong direction, keeping this from a straight A."));
children.push(new Paragraph({children:[txt("")],shading:fill(GRAY),spacing:{before:80,after:0}}));

const doc = new Document({
  styles:{default:{document:{ryn:{font:"Arial",size:18,color:TEXT}}}},
  sections:[{
    properties:{page:{size:{width:PAGE_W,height:PAGE_H},margin:{top:MARGIN,right:MARGIN,bottom:MARGIN,left:MARGIN}}},
    headers:{default:new Header({children:[new Paragraph({children:[txt("AMD Q1 2026  |  Simplified Income Statement Analysis  |  Quarter Ended March 28, 2026",{size:16,color:MUTED})],border:{bottom:{style:BorderStyle.SINGLE,size:4,color:"CCCCCC"}}})]})},
    footers:{default:new Footer({children:[new Paragraph({children:[txt("Advanced Micro Devices, Inc. (AMD)  |  Simplified GAAP Analysis  |  For Investment Research Purposes Only    Page ",{size:14,color:MUTED}),new TextRun({children:[PageNumber.CURRENT],font:"Arial",size:14,color:MUTED}),txt(" of ",{size:14,color:MUTED}),new TextRun({children:[PageNumber.TOTAL_PAGES],font:"Arial",size:14,color:MUTED})],alignment:AlignmentType.CENTER,border:{top:{style:BorderStyle.SINGLE,size:4,color:"CCCCCC"}}})]})},
    children,
  }]
});

Packer.toBuffer(doc).then(buf=>{
  fs.writeFileSync('/home/claude/AMD_Q1_2026_Simple.docx',buf);
  console.log('Done');
}).catch(e=>{console.error(e);process.exit(1);});
