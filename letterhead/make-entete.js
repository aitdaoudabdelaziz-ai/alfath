const fs=require('fs');const d=require('docx');
const {Document,Packer,Paragraph,TextRun,ImageRun,Table,TableRow,TableCell,WidthType,AlignmentType,BorderStyle,Header,Footer,VerticalAlign}=d;
const logo=fs.readFileSync('/home/user/alfath/logo.png');
const G='1E6B48',F='Arial';
const none={style:BorderStyle.NONE,size:0,color:'FFFFFF'};const nb={top:none,bottom:none,left:none,right:none};
const ar=(t,o={})=>new Paragraph({bidirectional:true,alignment:AlignmentType.CENTER,spacing:{after:o.after??20},children:[new TextRun({text:t,font:F,rightToLeft:true,size:o.size||18,bold:!!o.bold,color:o.color||'333333'})]});
const fr=(t,o={})=>new Paragraph({alignment:AlignmentType.CENTER,spacing:{after:o.after??20},children:[new TextRun({text:t,font:F,size:o.size||17,bold:!!o.bold,color:o.color||'333333'})]});
const W=[4200,1640,4200];
const cell=(ch,w)=>new TableCell({width:{size:w,type:WidthType.DXA},borders:nb,verticalAlign:VerticalAlign.CENTER,children:ch});
const head=new Table({width:{size:10040,type:WidthType.DXA},columnWidths:W,borders:{...nb,insideHorizontal:none,insideVertical:none},rows:[new TableRow({children:[
 cell([fr('ROYAUME DU MAROC',{size:15,color:'666666'}),fr('Association Al Fath pour le',{bold:true,size:20,color:G,after:0}),fr('Développement et la Coopération',{bold:true,size:20,color:G}),fr('Comité de l’eau potable',{size:16}),fr('Douar Agadir Isrghine – C. Akka Ighane',{size:15}),fr('Province de Tata',{size:15})],W[0]),
 cell([new Paragraph({alignment:AlignmentType.CENTER,children:[new ImageRun({type:'png',data:logo,transformation:{width:84,height:84}})]})],W[1]),
 cell([ar('المملكة المغربية',{size:16,color:'666666'}),ar('جمعية الفتح للتنمية والتعاون',{bold:true,size:28,color:G}),ar('لجنة الماء الصالح للشرب',{size:20}),ar('دوار أكادير اسرغين – جماعة أقايغان',{size:18}),ar('إقليم طاطا',{size:18})],W[2]),
]})]});
const rule=new Paragraph({spacing:{after:0},border:{bottom:{style:BorderStyle.DOUBLE,size:6,color:G,space:4}},children:[]});
const foot=new Footer({children:[new Paragraph({alignment:AlignmentType.CENTER,border:{top:{style:BorderStyle.SINGLE,size:6,color:G,space:4}},bidirectional:true,children:[new TextRun({text:'مقر الجمعية: دوار أكادير اسرغين، جماعة أقايغان، إقليم طاطا',font:F,size:16,rightToLeft:true,color:'444444'})]}),
 new Paragraph({alignment:AlignmentType.CENTER,children:[new TextRun({text:'Tél / WhatsApp : 06 78 53 51 96   •   Urgences : 06 58 69 42 18   •   association.elfath@gmail.com',font:F,size:16,color:'444444'})]})]});
const body=(t,o={})=>new Paragraph({bidirectional:true,alignment:o.al==='L'?AlignmentType.RIGHT:(o.al||AlignmentType.LEFT),spacing:{after:o.after??120,line:360},children:[new TextRun({text:t,font:F,size:24,rightToLeft:true,bold:!!o.bold})]});
const doc=new Document({styles:{default:{document:{run:{font:F,size:24}}}},sections:[{properties:{page:{size:{width:11906,height:16838},margin:{top:2950,bottom:1300,left:1000,right:1000,header:500,footer:400}}},
 headers:{default:new Header({children:[head,rule]})},footers:{default:foot},
 children:[body('أقايغان في: ..............................',{al:'L',after:360}),body('الرقم: ........ / ........',{after:240}),
 body('إلى السيد(ة): ..............................................',{al:AlignmentType.CENTER,bold:true,after:360}),
 body('الموضوع: ..............................................................',{bold:true,after:360}),
 body('سلام تام بوجود مولانا الإمام المؤيد بالله،'),body('وبعد، ...'),body(''),body(''),body(''),body(''),
 body('وتقبلوا فائق التقدير والاحترام.',{after:480}),body('عن الجمعية: الرئيس',{al:'L',bold:true})]}]});
Packer.toBuffer(doc).then(b=>fs.writeFileSync('/home/user/alfath/letterhead/entete-alfath.docx',b));
