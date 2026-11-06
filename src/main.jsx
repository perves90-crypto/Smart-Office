import React from 'react';
import * as XLSX from 'xlsx';
import {createRoot} from 'react-dom/client';
import {FileText,Table2,Presentation,File,FolderOpen,Clock3,Settings,ShieldCheck,Grid2X2,Home,Folder,Eye,Plus,Database,ClipboardCheck,Save,Printer,Undo2,Redo2,Copy,Scissors,Clipboard,Paintbrush,Type,Bold,Italic,Underline,AlignLeft,AlignCenter,AlignRight,WrapText,Merge,Filter,ArrowDownUp,Search,Check,MessageSquare,Lock,BarChart3,Image,Link,Calculator,Percent,Coins,CalendarDays,SortAsc,FileSpreadsheet,X,ChevronDown} from 'lucide-react';
import './style.css';

const ribbonData = {
 Home: [
  {group:'Clipboard',items:[['Undo',Undo2],['Redo',Redo2],['Paste',Clipboard],['Cut',Scissors],['Copy',Copy],['Format Painter',Paintbrush]]},
  {group:'Font',items:[['Font',Type],['Bold',Bold],['Italic',Italic],['Underline',Underline],['Font Color',Type]]},
  {group:'Alignment',items:[['Align Left',AlignLeft],['Center',AlignCenter],['Align Right',AlignRight],['Wrap Text',WrapText],['Merge Cells',Merge]]},
  {group:'Number',items:[['Number Format',Calculator],['Currency',Coins],['Percent',Percent],['Date',CalendarDays]]},
  {group:'Cells',items:[['Insert Rows',Plus],['Insert Columns',Plus],['Delete Rows',X],['Delete Columns',X],['Format',Grid2X2]]},
  {group:'Editing',items:[['AutoSum',Calculator],['Fill',ChevronDown],['Clear',X],['Find',Search],['Sort & Filter',Filter]]}
 ],
 File: [
  {group:'File',items:[['New Workbook',Plus],['Open',FolderOpen],['Save',Save],['Save As',FileSpreadsheet],['Print',Printer],['Export PDF',File]]},
  {group:'Information',items:[['Workbook Info',FileText],['Protect Workbook',Lock],['Version History',Clock3]]}
 ],
 View: [
  {group:'Workbook Views',items:[['Normal',Grid2X2],['Page Layout',FileText],['Page Break',FileSpreadsheet]]},
  {group:'Show',items:[['Formula Bar',Calculator],['Gridlines',Grid2X2],['Headings',Table2],['Ruler',AlignLeft]]},
  {group:'Zoom',items:[['Zoom In',Plus],['Zoom Out',Search],['100%',Eye]]},
  {group:'Window',items:[['Freeze Panes',Lock],['Split',Grid2X2],['New Window',Plus]]}
 ],
 Insert: [
  {group:'Tables',items:[['Table',Table2],['PivotTable',Grid2X2]]},
  {group:'Illustrations',items:[['Picture',Image],['Shapes',Grid2X2],['Icons',ShieldCheck]]},
  {group:'Charts',items:[['Recommended Charts',BarChart3],['Column Chart',BarChart3],['Line Chart',BarChart3],['Pie Chart',BarChart3]]},
  {group:'Links & Text',items:[['Hyperlink',Link],['Text Box',Type],['Header & Footer',FileText]]}
 ],
 Data: [
  {group:'Get & Transform',items:[['Get Data',Database],['From Text/CSV',FileText],['Refresh All',Redo2]]},
  {group:'Sort & Filter',items:[['Sort A to Z',SortAsc],['Sort Z to A',ArrowDownUp],['Filter',Filter],['Clear Filter',X]]},
  {group:'Data Tools',items:[['Text to Columns',ColumnsIcon],['Remove Duplicates',Copy],['Data Validation',Check]]},
  {group:'Forecast',items:[['What-If Analysis',Calculator],['Goal Seek',BarChart3]]}
 ],
 Review: [
  {group:'Proofing',items:[['Spelling',Check],['Workbook Statistics',FileSpreadsheet]]},
  {group:'Comments',items:[['New Comment',MessageSquare],['Show Comments',MessageSquare],['Notes',FileText]]},
  {group:'Protection',items:[['Protect Sheet',Lock],['Protect Workbook',ShieldCheck],['Allow Edit Ranges',Check]]}
 ]
};
function ColumnsIcon(){return <Table2/>}
const menus=[['Home',Home],['File',Folder],['View',Eye],['Insert',Plus],['Data',Database],['Review',ClipboardCheck]];
function cellToPosition(ref){
 const m=String(ref).toUpperCase().match(/^([A-Z]+)([1-9][0-9]*)$/);if(!m)return null;
 let col=0;for(const ch of m[1])col=col*26+ch.charCodeAt(0)-64;
 return {col,row:Number(m[2])};
}
function positionToCell(col,row){let s='';while(col>0){const n=(col-1)%26;s=String.fromCharCode(65+n)+s;col=Math.floor((col-1)/26)}return s+row}
function numericValue(value){
 if(typeof value==='number')return Number.isFinite(value)?value:null;
 if(value===null||value===undefined||String(value).trim()==='')return null;
 let s=String(value).trim().replace(/[৳$€£,\s]/g,'');
 if(s.endsWith('%')){const n=Number(s.slice(0,-1));return Number.isFinite(n)?n/100:null}
 if(/^\(.*\)$/.test(s))s='-'+s.slice(1,-1);
 const n=Number(s);return Number.isFinite(n)?n:null;
}
function calculateCell(ref,cells,stack=[]){
 const key=String(ref).toUpperCase();
 if(stack.includes(key))return '#CYCLE!';
 const raw=cells[key];if(raw===undefined||raw===null||String(raw).trim()==='')return 0;
 if(typeof raw==='string'&&raw.trim().startsWith('='))return calculateFormula(raw.trim().slice(1),cells,[...stack,key]);
 const n=numericValue(raw);return n===null?raw:n;
}
function calculateFormula(source,cells,stack=[]){
 const tokenize=(input)=>{
  const out=[];let i=0;
  while(i<input.length){
   const rest=input.slice(i);
   if(/^\s+/.test(rest)){i+=rest.match(/^\s+/)[0].length;continue}
   const m=rest.match(/^(?:"(?:[^"]|"")*"|(?:\$?[A-Z]{1,3}\$?[1-9][0-9]*)|(?:\d+(?:\.\d*)?|\.\d+)|(?:<=|>=|<>|[+\-*/^%&=<>(),:]))/i);
   if(!m)throw Error('Invalid token');
   out.push(m[0]);i+=m[0].length;
  }
  return out;
 };
 const tokens=tokenize(source);let index=0;
 const peek=()=>tokens[index];
 const take=()=>tokens[index++];
 const isCell=t=>typeof t==='string'&&/^\$?[A-Z]{1,3}\$?[1-9][0-9]*$/i.test(t);
 const cellValue=ref=>{
  const clean=ref.replace(/\$/g,'').toUpperCase();
  const value=calculateCell(clean,cells,stack);
  if(value==='#CYCLE!')throw Error('#CYCLE!');
  return value;
 };
 const rangeValues=(first,last)=>{
  const a=cellToPosition(first.replace(/\$/g,'')),b=cellToPosition(last.replace(/\$/g,''));
  if(!a||!b)throw Error('Bad range');
  const values=[];
  for(let r=Math.min(a.row,b.row);r<=Math.max(a.row,b.row);r++)
   for(let c=Math.min(a.col,b.col);c<=Math.max(a.col,b.col);c++)values.push(cellValue(positionToCell(c,r)));
  return values;
 };
 const flatten=items=>items.flat(Infinity);
 const numeric=v=>{if(typeof v==='number')return v;if(v===true)return 1;if(v===false)return 0;return numericValue(v)};
 const nums=items=>flatten(items).map(numeric).filter(v=>v!==null);
 const textValue=v=>v===null||v===undefined?'':String(v);
 const truth=v=>typeof v==='string'?v.length>0&&v.toUpperCase()!=='FALSE':!!v;
 const parseArgs=()=>{
  const args=[];
  if(peek()===')'){take();return args}
  while(true){args.push(parseCompare());if(peek()===','){take();continue}if(take()!==')')throw Error('Missing )');return args}
 };
 const call=(name,args)=>{
  const flat=flatten(args),n=nums(args),a=args[0],b=args[1];
  switch(name){
   case 'SUM':return n.reduce((x,y)=>x+y,0);
   case 'AVERAGE':case 'AVERAGEA':return n.length?n.reduce((x,y)=>x+y,0)/n.length:0;
   case 'COUNT':return n.length;
   case 'COUNTA':return flat.filter(v=>v!==''&&v!==null&&v!==undefined).length;
   case 'COUNTBLANK':return flat.filter(v=>v===''||v===null||v===undefined).length;
   case 'MAX':return n.length?Math.max(...n):0;
   case 'MIN':return n.length?Math.min(...n):0;
   case 'PRODUCT':return n.reduce((x,y)=>x*y,1);
   case 'ABS':return Math.abs(Number(a));
   case 'ROUND':return Number(Number(a).toFixed(Math.max(-10,Math.min(10,Number(b)||0))));
   case 'ROUNDUP':{const p=Math.pow(10,Number(b)||0);return Math.sign(Number(a))*Math.ceil(Math.abs(Number(a))*p)/p}
   case 'ROUNDDOWN':{const p=Math.pow(10,Number(b)||0);return Math.sign(Number(a))*Math.floor(Math.abs(Number(a))*p)/p}
   case 'INT':return Math.floor(Number(a));
   case 'MOD':return Number(a)%Number(b);
   case 'POWER':return Math.pow(Number(a),Number(b));
   case 'SQRT':return Math.sqrt(Number(a));
   case 'PI':return Math.PI;
   case 'RAND':return Math.random();
   case 'IF':return truth(a)?(args[1]??true):(args[2]??false);
   case 'AND':return flat.every(truth);
   case 'OR':return flat.some(truth);
   case 'NOT':return !truth(a);
   case 'IFERROR':return a;
   case 'CONCAT':case 'CONCATENATE':return flat.map(textValue).join('');
   case 'LEFT':return textValue(a).slice(0,Math.max(0,Number(b??1)));
   case 'RIGHT':return textValue(a).slice(-Math.max(0,Number(b??1)));
   case 'MID':return textValue(a).slice(Math.max(0,Number(b)-1),Math.max(0,Number(b)-1)+Math.max(0,Number(args[2])));
   case 'LEN':return textValue(a).length;
   case 'UPPER':return textValue(a).toUpperCase();
   case 'LOWER':return textValue(a).toLowerCase();
   case 'TRIM':return textValue(a).trim().replace(/\s+/g,' ');
   case 'PROPER':return textValue(a).toLowerCase().replace(/\b\w/g,c=>c.toUpperCase());
   case 'TODAY':return new Date().toISOString().slice(0,10);
   case 'NOW':return new Date().toLocaleString();
   case 'SUMIF':case 'COUNTIF':case 'AVERAGEIF':{
    const range=flatten([args[0]]),criteria=textValue(args[1]),sumRange=flatten([args[2]??args[0]]);
    const matches=range.map((v,i)=>({v,i})).filter(({v})=>criteriaMatch(v,criteria));
    if(name==='COUNTIF')return matches.length;
    const selected=matches.map(({i})=>numeric(sumRange[i])).filter(v=>v!==null);
    if(name==='SUMIF')return selected.reduce((x,y)=>x+y,0);
    return selected.length?selected.reduce((x,y)=>x+y,0)/selected.length:0;
   }
   case 'SUMPRODUCT':{
    const arrays=args.map(x=>flatten([x]).map(v=>numeric(v)??0));const len=Math.min(...arrays.map(x=>x.length));
    let total=0;for(let i=0;i<len;i++)total+=arrays.reduce((p,arr)=>p*arr[i],1);return total;
   }
   case 'COUNTIF':return 0;
   default:throw Error('Unknown function '+name);
  }
 };
 const criteriaMatch=(value,criteria)=>{
  const s=textValue(criteria),m=s.match(/^(>=|<=|<>|>|<|=)(.*)$/);
  if(!m)return textValue(value).toLowerCase()===s.toLowerCase();
  const left=numeric(value),right=numeric(m[2]);
  const x=left!==null?left:textValue(value),y=right!==null?right:m[2];
  switch(m[1]){case '>':return x>y;case '<':return x<y;case '>=':return x>=y;case '<=':return x<=y;case '<>':return x!==y;default:return x===y}
 };
 const primary=()=>{
  const t=take();if(t===undefined)throw Error('Missing value');
  if(t==='('){const v=parseCompare();if(take()!==')')throw Error('Missing )');return v}
  if(t[0]==='"')return t.slice(1,-1).replace(/""/g,'"');
  if(/^\d/.test(t)||/^\.\d/.test(t))return Number(t);
  if(isCell(t)){
   if(peek()===':'){take();const second=take();if(!isCell(second))throw Error('Bad range');return rangeValues(t,second)}
   return cellValue(t);
  }
  if(/^[A-Z_][A-Z0-9_.]*$/i.test(t)&&peek()==='('){take();return call(t.toUpperCase(),parseArgs())}
  throw Error('Unexpected value');
 };
 const unary=()=>{if(peek()==='+'){take();return +unary()}if(peek()==='-'){take();return -unary()}let v=primary();while(peek()==='%'){take();v=Number(v)/100}return v};
 const power=()=>{let v=unary();if(peek()==='^'){take();v=Math.pow(Number(v),Number(power()))}return v};
 const mul=()=>{let v=power();while(peek()==='*'||peek()==='/'){const op=take(),r=power();v=op==='*'?Number(v)*Number(r):Number(v)/Number(r)}return v};
 const add=()=>{let v=mul();while(peek()==='+'||peek()==='-'){const op=take(),r=mul();v=op==='+'?Number(v)+Number(r):Number(v)-Number(r)}return v};
 const concat=()=>{let v=add();while(peek()==='&'){take();v=textValue(v)+textValue(add())}return v};
 const parseCompare=()=>{let v=concat();while(['=','<>','>','<','>=','<='].includes(peek())){const op=take(),r=concat();switch(op){case '=':v=v===r;break;case '<>':v=v!==r;break;case '>':v=v>r;break;case '<':v=v<r;break;case '>=':v=v>=r;break;case '<=':v=v<=r;break}}return v};
 try{const result=parseCompare();if(index!==tokens.length)throw Error('Unexpected token');if(typeof result==='number'&&!Number.isFinite(result))return '#DIV/0!';return result}catch(error){return error.message==='#CYCLE!'?'#CYCLE!':'#VALUE!'}
}
function loadSavedWorkbook(){try{const raw=localStorage.getItem('smartoffice-workbook');if(!raw)return null;const data=JSON.parse(raw);return data&&typeof data==='object'?data:null}catch{return null}}
function ExcelEditor({onBack}){
 const [initialWorkbook]=React.useState(loadSavedWorkbook);
 const [isRestoring,setIsRestoring]=React.useState(true);
 const [isImporting,setIsImporting]=React.useState(false);
 React.useEffect(()=>{let live=true;try{const req=indexedDB.open('SmartOfficeWorkbookDB',1);req.onupgradeneeded=()=>{if(!req.result.objectStoreNames.contains('workbooks'))req.result.createObjectStore('workbooks')};req.onsuccess=()=>{const db=req.result,tx=db.transaction('workbooks','readonly'),get=tx.objectStore('workbooks').get('current');get.onsuccess=()=>{if(live&&get.result){const d=get.result;setSheetStore(d.sheets||{});setCells(d.cells||{});setCellStyles(d.cellStyles||{});setSheet(d.sheet||'Sheet1');setWorkbookName(d.workbookName||'Book1');setChart(d.chart||null);setPivot(d.pivot||null)}if(live)setIsRestoring(false);db.close()};get.onerror=()=>{if(live)setIsRestoring(false);db.close()}};req.onerror=()=>{if(live)setIsRestoring(false)}}catch{setIsRestoring(false)}return()=>{live=false}},[]);
 const [active,setActive]=React.useState('Home');
 const [menuOpen,setMenuOpen]=React.useState(false);
 const [selected,setSelected]=React.useState('A1');
 const [selectionBand,setSelectionBand]=React.useState(null);
 const [formula,setFormula]=React.useState('');
 const [editingCell,setEditingCell]=React.useState(null);
 const [cells,setCells]=React.useState(()=>initialWorkbook?.cells||{});
 const [cellStyles,setCellStyles]=React.useState(()=>initialWorkbook?.cellStyles||{});
 const [sheetStore,setSheetStore]=React.useState(()=>initialWorkbook?.sheets||{[initialWorkbook?.sheet||'Sheet1']:{cells:initialWorkbook?.cells||{},cellStyles:initialWorkbook?.cellStyles||{}}});
 const [cellHistory,setCellHistory]=React.useState(()=>[initialWorkbook?.cells||{}]);
 const [historyIndex,setHistoryIndex]=React.useState(0);
 const [clipboard,setClipboard]=React.useState(null);
 const [sheet,setSheet]=React.useState(()=>initialWorkbook?.sheet||'Sheet1');
 const [workbookName,setWorkbookName]=React.useState(()=>initialWorkbook?.workbookName||'Book1');
 const [notice,setNotice]=React.useState('');
 const [chart,setChart]=React.useState(()=>initialWorkbook?.chart||null);
 const [pivotWizard,setPivotWizard]=React.useState(null);
 const [rangeStart,setRangeStart]=React.useState(null);
 const [rangeEnd,setRangeEnd]=React.useState(null);
 const [isSelectingRange,setIsSelectingRange]=React.useState(false);
 const [visibleRows,setVisibleRows]=React.useState(100);
 const [visibleCols,setVisibleCols]=React.useState(10);
 const [sheetZoom,setSheetZoom]=React.useState(1);
 const touchPoints=React.useRef(new Map());
 const touchGesture=React.useRef(null);
 const cellTap=React.useRef({id:null,time:0});
 const [pivot,setPivot]=React.useState(()=>initialWorkbook?.pivot||null);
 const [showFormula,setShowFormula]=React.useState(true);
 const [showGrid,setShowGrid]=React.useState(true);
 const [filter,setFilter]=React.useState(false);
 const colName=n=>{let s='';while(n>0){let r=(n-1)%26;s=String.fromCharCode(65+r)+s;n=Math.floor((n-1)/26)}return s};
 const notify=t=>{setNotice(t);setTimeout(()=>setNotice(''),1800)};
 React.useEffect(()=>{if(isRestoring)return;const timer=setTimeout(()=>{const all={...sheetStore,[sheet]:{cells,cellStyles}};setSheetStore(all);const data={sheets:all,cells,cellStyles,sheet,workbookName,chart,pivot,savedAt:new Date().toISOString()};try{const request=indexedDB.open('SmartOfficeWorkbookDB',1);request.onupgradeneeded=()=>{if(!request.result.objectStoreNames.contains('workbooks'))request.result.createObjectStore('workbooks')};request.onsuccess=()=>{const db=request.result,tx=db.transaction('workbooks','readwrite');tx.objectStore('workbooks').put(data,'current');tx.oncomplete=()=>db.close();tx.onerror=()=>{db.close();notify('Workbook auto-save failed. Please export a backup')}};request.onerror=()=>notify('Device storage unavailable. Please export a backup')}catch{notify('Device storage unavailable. Please export a backup')}},500);return()=>clearTimeout(timer)},[cells,cellStyles,sheet,workbookName,chart,pivot,isRestoring]);
 const selectCell=(id)=>{setSelected(id);setFormula(cells[id]||'');setSelectionBand(null)};
 const selectWholeColumn=(col)=>{const id=positionToCell(col,1);setSelected(id);setFormula(cells[id]||'');setSelectionBand({axis:'col',index:col});notify(colName(col)+' কলাম সম্পূর্ণ সিলেক্ট হয়েছে');};
 const selectWholeRow=(row)=>{const id=positionToCell(1,row);setSelected(id);setFormula(cells[id]||'');setSelectionBand({axis:'row',index:row});notify(row+' নম্বর Row সম্পূর্ণ সিলেক্ট হয়েছে');};
 const updateCell=(id,value)=>{
  const next={...cells,[id]:value};
  setCells(next);
  setCellHistory(h=>[...h.slice(0,historyIndex+1),next].slice(-101));
  setHistoryIndex(i=>Math.min(i+1,100));
 };
 const undo=()=>{if(historyIndex<=0){notify('Nothing to undo');return}const i=historyIndex-1;setHistoryIndex(i);setCells(cellHistory[i]||{});setFormula((cellHistory[i]||{})[selected]||'');notify('Undo complete')};
 const redo=()=>{if(historyIndex>=cellHistory.length-1){notify('Nothing to redo');return}const i=historyIndex+1;setHistoryIndex(i);setCells(cellHistory[i]||{});setFormula((cellHistory[i]||{})[selected]||'');notify('Redo complete')};
 const copyCell=async(cut=false)=>{
  const item={kind:'cell',value:cells[selected]??'',style:cellStyles[selected]||{}};
  setClipboard(item);
  try{if(navigator.clipboard?.writeText)await navigator.clipboard.writeText(String(item.value))}catch{}
  if(cut)updateCell(selected,'');
  notify(cut?'Cell cut':'Cell copied');
 };
 const pasteCell=()=>{
  if(!clipboard){notify('Copy a cell first');return}
  if(clipboard.kind==='style'){
   setCellStyles(p=>({...p,[selected]:{...clipboard.style}}));
   notify('Format applied');return;
  }
  updateCell(selected,clipboard.value);
  if(Object.keys(clipboard.style||{}).length)setCellStyles(p=>({...p,[selected]:{...clipboard.style}}));
  setFormula(clipboard.value);
  notify('Cell pasted');
 };
 const shiftCells=(axis,at,remove=false)=>{
  const remap=(source)=>{const out={};for(const [key,value] of Object.entries(source)){const p=cellToPosition(key);if(!p)continue;const n=axis==='row'?p.row:p.col;if(remove&&n===at)continue;const moved=remove?(n>at?n-1:n):(n>=at?n+1:n);const dest=axis==='row'?positionToCell(p.col,moved):positionToCell(moved,p.row);out[dest]=value;}return out;};
  const next=remap(cells),styles=remap(cellStyles);setCells(next);setCellStyles(styles);setCellHistory(h=>[...h.slice(0,historyIndex+1),next].slice(-101));setHistoryIndex(i=>Math.min(i+1,100));setFormula(next[selected]||'');notify((remove?'Deleted ':'Inserted ')+(axis==='row'?'row':'column'));
 };
 const updateStyle=(patch)=>setCellStyles(p=>({...p,[selected]:{...(p[selected]||{}),...patch}}));
 const toggleStyle=(key)=>updateStyle({[key]:!cellStyles[selected]?.[key]});
 React.useEffect(()=>{
  const onKeyDown=e=>{
   if(!(e.ctrlKey||e.metaKey))return;
   const key=e.key.toLowerCase();
   if(key==='z'){e.preventDefault();undo()}
   else if(key==='y'){e.preventDefault();redo()}
   else if(key==='c'){e.preventDefault();copyCell(false)}
   else if(key==='x'){e.preventDefault();copyCell(true)}
   else if(key==='v'){e.preventDefault();pasteCell()}
  };
  window.addEventListener('keydown',onKeyDown);
  return ()=>window.removeEventListener('keydown',onKeyDown);
 });
 const importWorkbook=(file)=>{
  if(!file)return;
  if(file.size>50*1024*1024){notify('৫০ MB-এর বেশি ফাইল আপাতত সমর্থিত নয়');return}
  const ext=(file.name.split('.').pop()||'').toLowerCase();
  if(!['xlsx','xls','csv'].includes(ext)){notify('শুধু XLSX, XLS ও CSV ফাইল Open করা যাবে');return}
  setIsImporting(true);setNotice('ফাইল পড়া হচ্ছে…');
  const applyWorkbook=(imported,activeName)=>{
   const names=Object.keys(imported||{});
   if(!names.length)throw new Error('ফাইলে কোনো Worksheet পাওয়া যায়নি');
   const first=names.includes(activeName)?activeName:names[0];
   const next=imported[first]?.cells||{};
   setSheetStore(imported);setCells(next);setCellHistory([{},next]);setHistoryIndex(1);setCellStyles(imported[first]?.cellStyles||{});setSheet(first);setSelected('A1');setFormula(next.A1??'');setWorkbookName(file.name.replace(/\.(xlsx|xls|csv)$/i,''));setVisibleRows(100);setVisibleCols(10);notify('সফল! '+names.length+'টি Sheet Open হয়েছে');
  };
  const parseOnMainThread=(buffer)=>{
   try{
    setNotice('বিকল্প পদ্ধতিতে Excel ফাইল খোলা হচ্ছে…');
    const workbook=XLSX.read(buffer,{type:'array',cellFormula:true,dense:false,cellStyles:false,cellNF:false,cellHTML:false,bookFiles:false,bookVBA:false,WTF:false});
    const names=workbook.SheetNames||[],imported={};
    if(!names.length)throw new Error('কোনো Worksheet পাওয়া যায়নি');
    names.forEach(name=>{
     const ws=workbook.Sheets[name]||{},cells={};
     for(const address of Object.keys(ws)){
      if(!/^[A-Z]{1,3}[1-9][0-9]*$/.test(address))continue;
      const cell=ws[address];if(!cell)continue;
      if(cell.f)cells[address]='='+cell.f;
      else if(cell.v!==undefined&&cell.v!==null)cells[address]=typeof cell.v==='number'?cell.v:String(cell.v);
     }
     imported[name]={cells,cellStyles:{}};
    });
    const activeIndex=workbook.Workbook?.Views?.[0]?.activeTab;
    applyWorkbook(imported,Number.isInteger(activeIndex)?names[activeIndex]:names[0]);
   }catch(error){notify('ফাইল Open হয়নি: '+(error?.message||'Excel data পড়া যায়নি'))}
   finally{setIsImporting(false)}
  };
  const reader=new FileReader();
  reader.onload=event=>{
   const buffer=event.target.result;
   let worker=null,finished=false,timer=null;
   const fallback=()=>{if(finished)return;finished=true;if(timer)clearTimeout(timer);try{worker?.terminate()}catch{}parseOnMainThread(buffer)};
   try{
    worker=new Worker(new URL('./xlsx.worker.js',import.meta.url),{type:'module'});
    worker.onmessage=e=>{
     const msg=e.data||{};
     if(msg.type==='progress'){setNotice(msg.message||'Excel ফাইল প্রস্তুত হচ্ছে…');return}
     if(finished)return;
     if(msg.type==='error'){fallback();return}
     if(msg.type==='success'){
      finished=true;if(timer)clearTimeout(timer);worker.terminate();
      try{applyWorkbook(msg.sheets,msg.activeSheet)}catch{parseOnMainThread(buffer)}
      finally{setIsImporting(false)}
     }
    };
    worker.onerror=fallback;
    worker.onmessageerror=fallback;
    timer=setTimeout(fallback,25000);
    const copy=buffer.slice(0);
    worker.postMessage({buffer:copy},[copy]);
   }catch{fallback()}
  };
  reader.onerror=()=>{setIsImporting(false);notify('ফাইল পড়া যায়নি। আবার চেষ্টা করুন')};
  try{reader.readAsArrayBuffer(file)}catch(error){setIsImporting(false);notify('ফাইল পড়া যায়নি: '+(error?.message||''))}
 };
 const exportWorkbook=(format='xlsx')=>{
  try{const all={...sheetStore,[sheet]:{cells,cellStyles}};const wb=XLSX.utils.book_new();for(const [sheetName,data] of Object.entries(all)){let maxRow=0,maxCol=0;for(const key of Object.keys(data.cells||{})){const p=cellToPosition(key);if(p){maxRow=Math.max(maxRow,p.row);maxCol=Math.max(maxCol,p.col)}}const aoa=Array.from({length:Math.max(1,maxRow)},()=>Array(Math.max(1,maxCol)).fill(''));const ws=XLSX.utils.aoa_to_sheet(aoa);for(const [key,value] of Object.entries(data.cells||{})){if(typeof value==='string'&&value.startsWith('='))ws[key]={t:'n',f:value.slice(1)};else ws[key]={v:value,t:typeof value==='number'?'n':'s'};}ws['!ref']=XLSX.utils.encode_range({s:{r:0,c:0},e:{r:Math.max(0,maxRow-1),c:Math.max(0,maxCol-1)}});XLSX.utils.book_append_sheet(wb,ws,sheetName.slice(0,31)||'Sheet1');}const filename=(workbookName.trim()||'Book1')+(format==='csv'?'.csv':'.xlsx');if(format==='csv'){const ws=wb.Sheets[sheet];const csv=XLSX.utils.sheet_to_csv(ws);const blob=new Blob(['\\uFEFF',csv],{type:'text/csv;charset=utf-8'});const url=URL.createObjectURL(blob);const a=document.createElement('a');a.href=url;a.download=filename;a.click();URL.revokeObjectURL(url);}else XLSX.writeFile(wb,filename);notify('Workbook downloaded');}catch(error){notify('Export failed: '+error.message)}
 };
 const exportPivotSheet=()=>{
   if(!pivot)return;
   const base='PivotTable';let name=base,n=2;while(sheetStore[name]||name===sheet)name=base+' '+n++;
   const out={};const headers=[...pivot.headers,pivot.measureHeader||'Count'];
   headers.forEach((h,i)=>{out[positionToCell(i+1,1)]=h;});
   pivot.rows.forEach((row,r)=>{row.values.forEach((v,c)=>{out[positionToCell(c+1,r+2)]=v;});out[positionToCell(headers.length,r+2)]=row.result??row.aggregate??0;});
   const totalRow=pivot.rows.length+2;out[positionToCell(1,totalRow)]='Grand Total';out[positionToCell(headers.length,totalRow)]=pivot.grandTotal;
   const all={...sheetStore,[sheet]:{cells,cellStyles},[name]:{cells:out,cellStyles:{}}};
   setSheetStore(all);setCells(out);setCellStyles({});setCellHistory([out]);setHistoryIndex(0);setSheet(name);setSelected('A1');setFormula('');setPivot(null);notify('PivotTable নতুন Sheet-এ Export হয়েছে');
  };
  const addSheet=()=>{let n=Object.keys(sheetStore).length+1,name='Sheet'+n;while(sheetStore[name]){n++;name='Sheet'+n;}setSheetStore(p=>({...p,[sheet]:{cells,cellStyles},[name]:{cells:{},cellStyles:{}}}));setCells({});setCellStyles({});setCellHistory([{}]);setHistoryIndex(0);setSheet(name);setSelected('A1');setFormula('');notify('New sheet created');};
 const switchSheet=(name)=>{if(name===sheet)return;const all={...sheetStore,[sheet]:{cells,cellStyles}};const next=all[name]||{cells:{},cellStyles:{}};setSheetStore(all);setCells(next.cells);setCellStyles(next.cellStyles);setCellHistory([next.cells]);setHistoryIndex(0);setSheet(name);setSelected('A1');setFormula(next.cells.A1||'');};
 const renameSheet=()=>{const name=window.prompt('Enter new sheet name:',sheet);if(!name||name===sheet)return;if(name.length>31||/[\\\\/?*\\[\\]:]/.test(name)){notify('Sheet name is invalid (max 31 characters)');return}if(sheetStore[name]){notify('That sheet name already exists');return}const all={...sheetStore,[sheet]:{cells,cellStyles}};all[name]=all[sheet];delete all[sheet];setSheetStore(all);setSheet(name);notify('Sheet renamed');};
 const deleteSheet=()=>{const names=Object.keys({...sheetStore,[sheet]:{cells,cellStyles}});if(names.length<=1){notify('At least one sheet is required');return}if(!window.confirm('Delete sheet '+sheet+'?'))return;const all={...sheetStore,[sheet]:{cells,cellStyles}};delete all[sheet];const nextName=Object.keys(all)[0],next=all[nextName];setSheetStore(all);setSheet(nextName);setCells(next.cells);setCellStyles(next.cellStyles);setCellHistory([next.cells]);setHistoryIndex(0);setSelected('A1');setFormula(next.cells.A1||'');notify('Sheet deleted');};

 const readRange=(rangeText)=>{
  const parts=String(rangeText||'').toUpperCase().split(':');
  if(parts.length!==2)throw Error('Use a range like A1:B6');
  const start=cellToPosition(parts[0]),end=cellToPosition(parts[1]);
  if(!start||!end||end.row<start.row||end.col<start.col)throw Error('Invalid data range');
  const matrix=[];
  for(let r=start.row;r<=end.row;r++){const row=[];for(let c=start.col;c<=end.col;c++){const ref=positionToCell(c,r);row.push(cells[ref]===undefined?'':String(calculateCell(ref,cells)));}matrix.push(row);}
  if(matrix.length<2||matrix[0].length<2)throw Error('Select at least two columns and two rows');
  return matrix;
 };
 const insertChart=(type)=>{
  try{const range=window.prompt('Enter chart data range (example: A1:B6):','A1:B6');if(!range)return;const matrix=readRange(range);const labels=matrix.slice(1).map(row=>row[0]||'');const series=matrix[0].slice(1).map((name,i)=>({name:name||'Series '+(i+1),values:matrix.slice(1).map(row=>Number(row[i+1])||0)}));setChart({type,range,labels,series,title:type+' Chart'});notify(type+' chart inserted');}
  catch(error){notify(error.message)}
 };
 const createPivot=()=>{
  const filled=Object.entries(cells).map(([address,value])=>({address,value,pos:cellToPosition(address)})).filter(x=>x.pos&&x.value!==undefined&&x.value!==null&&String(x.value).trim()!=='');
  if(!filled.length){notify('PivotTable তৈরির জন্য আগে Sheet-এ ডেটা লিখুন');return}
  const minRow=Math.min(...filled.map(x=>x.pos.row)),maxRow=Math.max(...filled.map(x=>x.pos.row));
  const minCol=Math.min(...filled.map(x=>x.pos.col)),maxCol=Math.max(...filled.map(x=>x.pos.col));
  const start=positionToCell(minCol,minRow),end=positionToCell(maxCol,maxRow);
  setRangeStart(start);setRangeEnd(end);
  setPivotWizard({step:'range',rowFields:Array(10).fill(''),columnFields:Array(10).fill(''),range:''});
  notify('Sheet-এর সব লেখা থাকা সেল স্বয়ংক্রিয়ভাবে সিলেক্ট হয়েছে');
};
 const pivotNext=()=>{
  try{
   if(pivotWizard.step==='range'){
    if(!rangeStart||!rangeEnd)throw Error('প্রথমে ডেটার সেলগুলো টেনে সিলেক্ট করুন');
    const a=cellToPosition(rangeStart),b=cellToPosition(rangeEnd),top=Math.min(a.row,b.row),bottom=Math.max(a.row,b.row),left=Math.min(a.col,b.col),right=Math.max(a.col,b.col);
    const range=positionToCell(left,top)+':'+positionToCell(right,bottom),matrix=readRange(range);
    setPivotWizard({step:'fields',range,headers:matrix[0].map((h,i)=>String(h||'Column '+(i+1))),rowFields:Array(10).fill(''),columnFields:Array(10).fill(''),valueField:'',aggregation:'count'});setRangeStart(null);setRangeEnd(null);setIsSelectingRange(false);
   }else{
    const rowFields=pivotWizard.rowFields.filter(Boolean),columnFields=pivotWizard.columnFields.filter(Boolean);
    if(!rowFields.length&&!columnFields.length)throw Error('কমপক্ষে Rows অথবা Columns-এর একটি ঘরে Field দিন');
    const matrix=readRange(pivotWizard.range),headers=pivotWizard.headers;
    const allFields=[...rowFields,...columnFields],fieldIndexes=allFields.map(h=>headers.indexOf(h));
    const aggregation=pivotWizard.aggregation||'count',valueField=pivotWizard.valueField||'';
    const valueIndex=valueField?headers.indexOf(valueField):-1;
    if(aggregation!=='count'&&valueIndex<0)throw Error('যোগ/গড় করার জন্য Values থেকে একটি সংখ্যার Field নির্বাচন করুন');
    const grouped=new Map();let grandSum=0,grandCount=0;
    const cleanLabel=v=>String(v??'').replace(/\\s+/g,' ').trim()||'(blank)';
    for(const record of matrix.slice(1)){
     const values=fieldIndexes.map(i=>cleanLabel(record[i]));
     const key=JSON.stringify(values);
     if(!grouped.has(key))grouped.set(key,{values,aggregate:0,recordCount:0,valueCount:0});
     const item=grouped.get(key);item.recordCount++;
     if(aggregation==='count'){item.aggregate++;item.valueCount++;}
     else {const raw=record[valueIndex],num=numericValue(raw);if(num!==null){item.aggregate+=num;item.valueCount++;grandSum+=num;grandCount++;}}
    }
    const detailRows=[...grouped.values()].map(item=>({...item,result:aggregation==='average'?(item.valueCount?item.aggregate/item.valueCount:0):item.aggregate})).sort((a,b)=>JSON.stringify(a.values).localeCompare(JSON.stringify(b.values),'en',{numeric:true,sensitivity:'base'}));
    const rows=[];
    if(allFields.length>1){
     let currentParent=null,parentItems=[];
     const addParentSubtotal=()=>{
      if(currentParent===null)return;
      const parentSum=parentItems.reduce((n,x)=>n+x.aggregate,0),parentCount=parentItems.reduce((n,x)=>n+x.valueCount,0);
      const subtotal=aggregation==='count'?parentItems.reduce((n,x)=>n+x.recordCount,0):aggregation==='average'?(parentCount?parentSum/parentCount:0):parentSum;
      rows.push({values:[currentParent,...Array(allFields.length-1).fill('')],result:subtotal,rowType:'subtotal'});
     };
     for(const item of detailRows){
      const parent=item.values[0];
      if(currentParent!==null&&parent!==currentParent)addParentSubtotal();
      if(parent!==currentParent){currentParent=parent;parentItems=[];}
      parentItems.push(item);
      const shown=[...item.values];shown[0]='';
      rows.push({...item,values:shown,rowType:'detail'});
     }
     addParentSubtotal();
    }else rows.push(...detailRows.map(item=>({...item,rowType:'detail'})));
    const grandTotal=aggregation==='count'?matrix.length-1:aggregation==='average'?(grandCount?grandSum/grandCount:0):grandSum;
    const measureHeader=aggregation==='count'?'Count':(aggregation==='sum'?'Sum of ':'Average of ')+valueField;
    setPivot({range:pivotWizard.range,rowFields,columnFields,headers:allFields,rows,grandTotal,aggregation,valueField,measureHeader});setPivotWizard(null);notify('PivotTable রিপোর্ট তৈরি হয়েছে');
   }
  }catch(error){notify(error.message)}
 };
 const action=(name)=>{
  if(name==='Undo'){undo();return}
  if(name==='Redo'){redo();return}
  if(name==='Copy'){copyCell(false);return}
  if(name==='Cut'){copyCell(true);return}
  if(name==='Paste'){pasteCell();return}
  if(name==='Format Painter'){const style=cellStyles[selected]||{};setClipboard({kind:'style',value:'',style});notify('Format copied — select Paste to apply');return}
  if(name==='Save'||name==='Save As'){exportWorkbook('xlsx');return}
  if(name==='Open'){document.getElementById('excelFileInput')?.click();return}
  if(name==='Export PDF'){notify('PDF export will be added in a later update');return}
  if(name==='Export PDF'||name==='Print'){notify(name+' option selected');return}
  if(name==='Gridlines'){setShowGrid(v=>!v);return}
  if(name==='Formula Bar'){setShowFormula(v=>!v);return}
  if(name==='Filter'){setFilter(v=>!v);return}
  if(name==='New Workbook'){setCells({});setCellHistory([{}]);setHistoryIndex(0);setCellStyles({});setSheetStore({'Sheet1':{cells:{},cellStyles:{}}});setSheet('Sheet1');setSelected('A1');setFormula('');notify('New workbook created');return}
  if(name==='Column Chart'){insertChart('Column');return}
  if(name==='Line Chart'){insertChart('Line');return}
  if(name==='Pie Chart'||name==='Recommended Charts'){insertChart(name==='Pie Chart'?'Pie':'Column');return}
  if(name==='PivotTable'){createPivot();return}
  if(name==='AutoSum'){const p=cellToPosition(selected);if(p){let start=p.row-1;while(start>1&&cells[positionToCell(p.col,start-1)]!==undefined&&cells[positionToCell(p.col,start-1)]!=='')start--;const end=p.row-1;const ref=end>=start?positionToCell(p.col,start)+':'+positionToCell(p.col,end):positionToCell(p.col,1)+':'+positionToCell(p.col,1);const value='=SUM('+ref+')';setFormula(value);updateCell(selected,value);notify('SUM formula added')}return}
  if(name==='Insert Rows'||name==='Delete Rows'||name==='Insert Columns'||name==='Delete Columns'){
   const p=cellToPosition(selected);if(!p)return;
   const axis=name.includes('Rows')?'row':'col';const remove=name.startsWith('Delete');const at=axis==='row'?p.row:p.col;
   if(remove&&at===1){notify('The first row/column cannot be deleted');return;}
   shiftCells(axis,at,remove);return;
  }
  if(name==='Clear'){updateCell(selected,'');setFormula('');notify(selected+' সেল ক্লিয়ার হয়েছে');return}
  if(name==='Find'){const term=window.prompt('কোন লেখা/সংখ্যা খুঁজবেন?');if(term!==null&&term!==''){const found=Object.entries(cells).find(([,v])=>String(v).toLowerCase().includes(term.toLowerCase()));if(found){selectCell(found[0]);notify(found[0]+' পাওয়া গেছে')}else notify('কোনো মিল পাওয়া যায়নি')}return}
  if(name==='Bold'){toggleStyle('bold');return}
  if(name==='Italic'){toggleStyle('italic');return}
  if(name==='Underline'){toggleStyle('underline');return}
  if(name==='Align Left'){updateStyle({textAlign:'left'});return}
  if(name==='Center'){updateStyle({textAlign:'center'});return}
  if(name==='Align Right'){updateStyle({textAlign:'right'});return}
  if(name==='Font Color'){const color=window.prompt('Enter font color (example: #ff0000):',cellStyles[selected]?.color||'#202124');if(color)updateStyle({color});return}
  if(name==='Fill'){const color=window.prompt('Enter cell fill color (example: #fff2cc):',cellStyles[selected]?.backgroundColor||'#fff2cc');if(color)updateStyle({backgroundColor:color});return}
  if(name==='Currency'){updateStyle({numberFormat:'currency'});return}
  if(name==='Percent'){updateStyle({numberFormat:'percent'});return}
  if(name==='Date'){updateStyle({numberFormat:'date'});return}
  notify(name+' selected');
 };
 const usedMaxRow=Object.keys(cells).reduce((m,key)=>Math.max(m,cellToPosition(key)?.row||0),0);
 const usedMaxCol=Object.keys(cells).reduce((m,key)=>Math.max(m,cellToPosition(key)?.col||0),0);
 const rows=Array.from({length:visibleRows},(_,i)=>i+1),cols=Array.from({length:visibleCols},(_,i)=>i+1);
 return <div className="excelApp">
  <header className="excelTop"><button className="backBtn" onClick={onBack}>‹</button><div className="excelTitle"><b><Table2 size={19}/> Smart Excel</b><small>Workbook • {sheet}</small></div><button className="topIcon" onClick={()=>action('Save')}><Save size={19}/></button><button className="topIcon" onClick={()=>notify('More workbook options')}><Settings size={19}/></button></header>
  <div className="workbookBar"><span className="fileBadge"><FileSpreadsheet size={16}/></span><input aria-label="Workbook name" value={workbookName} onChange={e=>setWorkbookName(e.target.value)} /><span className="savedState">● Saved locally</span></div>
  <div className="ribbonStrip">{ribbonData[active].map(group=><div className="ribbonGroup" key={group.group}><div className="ribbonButtons">{group.items.map(([label,Icon])=><button key={label} className="ribbonAction" onClick={()=>action(label)} title={label}><Icon size={17}/><span>{label}</span></button>)}</div><small>{group.group}</small></div>)}</div>
  {showFormula&&<div className="formulaBar"><span className="cellRef">{selected}</span><span className="fx">fx</span><input value={formula} onChange={e=>{setFormula(e.target.value);updateCell(selected,e.target.value)}} placeholder="Enter value or formula"/></div>}
  <div className="sheetScroll" style={{touchAction:'none'}} onPointerDown={e=>{const box=e.currentTarget;touchPoints.current.set(e.pointerId,{x:e.clientX,y:e.clientY});try{box.setPointerCapture(e.pointerId)}catch{};if(touchPoints.current.size>=2){setIsSelectingRange(false);const pts=[...touchPoints.current.values()].slice(0,2),dx=pts[0].x-pts[1].x,dy=pts[0].y-pts[1].y;touchGesture.current={mode:'pinch',distance:Math.max(1,Math.hypot(dx,dy)),zoom:sheetZoom,midX:(pts[0].x+pts[1].x)/2,midY:(pts[0].y+pts[1].y)/2,scrollLeft:box.scrollLeft,scrollTop:box.scrollTop};return;}if(pivotWizard?.step==='range'){const td=e.target.closest('td[data-cell]');if(!td)return;e.preventDefault();const id=td.dataset.cell;setIsSelectingRange(true);setRangeStart(id);setRangeEnd(id);touchGesture.current={mode:'select'};}else{touchGesture.current={mode:'pan',x:e.clientX,y:e.clientY,scrollLeft:box.scrollLeft,scrollTop:box.scrollTop};}}} onPointerMove={e=>{if(!touchPoints.current.has(e.pointerId))return;touchPoints.current.set(e.pointerId,{x:e.clientX,y:e.clientY});const box=e.currentTarget;if(touchPoints.current.size>=2){const pts=[...touchPoints.current.values()].slice(0,2),g=touchGesture.current;if(!g||g.mode!=='pinch')return;const dx=pts[0].x-pts[1].x,dy=pts[0].y-pts[1].y,dist=Math.max(1,Math.hypot(dx,dy)),midX=(pts[0].x+pts[1].x)/2,midY=(pts[0].y+pts[1].y)/2;setSheetZoom(Math.max(.6,Math.min(2.5,g.zoom*dist/g.distance)));box.scrollLeft=g.scrollLeft-(midX-g.midX);box.scrollTop=g.scrollTop-(midY-g.midY);return;}const g=touchGesture.current;if(g?.mode==='select'){const rect=box.getBoundingClientRect(),edge=42,step=26;if(e.clientX>rect.right-edge)box.scrollLeft+=step;else if(e.clientX<rect.left+edge)box.scrollLeft-=step;if(e.clientY>rect.bottom-edge)box.scrollTop+=step;else if(e.clientY<rect.top+edge)box.scrollTop-=step;const px=Math.max(rect.left+2,Math.min(rect.right-2,e.clientX)),py=Math.max(rect.top+2,Math.min(rect.bottom-2,e.clientY));const el=document.elementFromPoint(px,py),td=el?.closest?.('td[data-cell]');if(td)setRangeEnd(td.dataset.cell);}else if(g?.mode==='pan'){box.scrollLeft=g.scrollLeft-(e.clientX-g.x);box.scrollTop=g.scrollTop-(e.clientY-g.y);}}} onPointerUp={e=>{touchPoints.current.delete(e.pointerId);if(touchPoints.current.size===0){touchGesture.current=null;setIsSelectingRange(false);}else if(touchPoints.current.size===1){const p=[...touchPoints.current.values()][0];touchGesture.current={mode:'pan',x:p.x,y:p.y,scrollLeft:e.currentTarget.scrollLeft,scrollTop:e.currentTarget.scrollTop};}}} onPointerCancel={e=>{touchPoints.current.delete(e.pointerId);touchGesture.current=null;setIsSelectingRange(false)}} onScroll={e=>{const el=e.currentTarget;if(el.scrollTop+el.clientHeight>=el.scrollHeight-180)setVisibleRows(v=>v+100);if(el.scrollLeft+el.clientWidth>=el.scrollWidth-180)setVisibleCols(v=>v+10)}}><table style={{zoom:sheetZoom}} className={showGrid?'sheetGrid':'sheetGrid noGrid'}><thead><tr><th className="corner" onClick={()=>{setSelectionBand({axis:'all'});setSelected('A1');setFormula(cells.A1||'');notify('সম্পূর্ণ Sheet সিলেক্ট হয়েছে')}}>◆</th>{cols.map(c=><th key={c} onClick={()=>selectWholeColumn(c)} className={selectionBand?.axis==='col'&&selectionBand.index===c?'axisSelected':''}>{colName(c)}</th>)}</tr></thead><tbody>{rows.map(r=><tr key={r}><th onClick={()=>selectWholeRow(r)} className={selectionBand?.axis==='row'&&selectionBand.index===r?'axisSelected':''}>{r}</th>{cols.map(c=>{const id=colName(c)+r;const style=cellStyles[id]||{};let display=cells[id]===undefined?'':String(calculateCell(id,cells));if(style.numberFormat==='currency'&&display!==''&&Number.isFinite(Number(display)))display='৳ '+Number(display).toFixed(2);if(style.numberFormat==='percent'&&display!==''&&Number.isFinite(Number(display)))display=(Number(display)*100).toFixed(2)+'%';return <td key={id} data-cell={id} onClick={()=>{if(pivotWizard?.step==='range')return;const now=Date.now(),last=cellTap.current;if(last.id===id&&now-last.time<450){selectCell(id);setEditingCell(id);cellTap.current={id:null,time:0};setTimeout(()=>{const el=document.querySelector('input[aria-label="'+id+'"]');el?.focus();el?.select();},0);}else{selectCell(id);setEditingCell(null);cellTap.current={id,time:now};}}} className={(selected===id?'selectedCell ':'')+((selectionBand?.axis==='col'&&selectionBand.index===c)||(selectionBand?.axis==='row'&&selectionBand.index===r)||selectionBand?.axis==='all'?'bandSelected ':'')+(pivotWizard?.step==='range'&&rangeStart&&rangeEnd&&(()=>{const p=cellToPosition(id),a=cellToPosition(rangeStart),b=cellToPosition(rangeEnd);return p&&a&&b&&p.row>=Math.min(a.row,b.row)&&p.row<=Math.max(a.row,b.row)&&p.col>=Math.min(a.col,b.col)&&p.col<=Math.max(a.col,b.col)})()?'rangeSelected':'')}><input style={{fontWeight:style.bold?'700':'400',fontStyle:style.italic?'italic':'normal',textDecoration:style.underline?'underline':'none',textAlign:style.textAlign||'left',color:style.color||'inherit',backgroundColor:style.backgroundColor||'transparent'}} readOnly={editingCell!==id} aria-label={id} value={editingCell===id?(cells[id]||''):display} onFocus={()=>{if(editingCell===id)selectCell(id)}} onBlur={()=>setEditingCell(null)} onChange={e=>{updateCell(id,e.target.value);if(selected===id)setFormula(e.target.value)}}/></td>})}</tr>)}</tbody></table></div>

  {pivotWizard&&<section className="pivotWizard"><div className="excelPanelHead"><b>PivotTable তৈরি করুন</b><button onClick={()=>setPivotWizard(null)}>×</button></div>{pivotWizard.step==='range'?<><p>Sheet-এর সব লেখা থাকা সেল স্বয়ংক্রিয়ভাবে সিলেক্ট করা হয়েছে। প্রয়োজনে রেঞ্জটি দেখে Next চাপুন।</p><b>নির্বাচিত রেঞ্জ: {rangeStart&&rangeEnd?(()=>{const a=cellToPosition(rangeStart),b=cellToPosition(rangeEnd);return positionToCell(Math.min(a.col,b.col),Math.min(a.row,b.row))+':'+positionToCell(Math.max(a.col,b.col),Math.max(a.row,b.row))})():'এখনো সিলেক্ট হয়নি'}</b><button className="pivotNext" onClick={pivotNext}>Next</button></>:<><p>নিচের Field-গুলো Rows ও Columns-এর ১০টি করে ঘরে টেনে রাখুন। যতগুলো ঘরে Field দেবেন, রিপোর্টে ততগুলো স্তর ব্যবহার হবে।</p><div className="pivotFields">{pivotWizard.headers.map(h=><div draggable onDragStart={e=>e.dataTransfer.setData('text/plain',h)} onClick={()=>notify('Field-টি Rows বা Columns-এর ঘরে টেনে দিন')} className="pivotField" key={h}>{h}</div>)}</div><div className="pivotMeasureBox"><b>Values / হিসাবের Field</b><select value={pivotWizard.valueField} onChange={e=>setPivotWizard(p=>({...p,valueField:e.target.value}))}><option value="">Field নির্বাচন করুন</option>{pivotWizard.headers.map(h=><option key={h} value={h}>{h}</option>)}</select><b>হিসাবের ধরন</b><select value={pivotWizard.aggregation} onChange={e=>setPivotWizard(p=>({...p,aggregation:e.target.value}))}><option value="count">Count — কয়টি রেকর্ড</option><option value="sum">Sum — মোট যোগফল</option><option value="average">Average — গড়</option></select><small>Count হলে Values Field লাগবে না। Sum/Average হলে সংখ্যার কলাম নির্বাচন করুন।</small></div><div className="pivotDropZones pivotTenZones">{['rowFields','columnFields'].map((zone,zi)=><div className="pivotZone" key={zone}><b>{zi===0?'Rows':'Columns'}</b>{pivotWizard[zone].map((field,index)=><div className="pivotSlot" key={index} onDragOver={e=>e.preventDefault()} onDrop={e=>{e.preventDefault();const h=e.dataTransfer.getData('text/plain');setPivotWizard(p=>({...p,[zone]:p[zone].map((v,i)=>i===index?h:v)}))}}><small>{index+1}</small><span>{field||'Field এখানে দিন'}</span><button onClick={()=>setPivotWizard(p=>({...p,[zone]:p[zone].map((v,i)=>i===index?'':v)}))}>×</button></div>)}</div>)}</div><button className="pivotNext" onClick={pivotNext}>Next — PivotTable রিপোর্ট তৈরি করুন</button></>}</section>}
   {chart&&<section className="excelDataPanel"><div className="excelPanelHead"><b>{chart.title}</b><button onClick={()=>setChart(null)}>×</button></div><div className="excelChart"><svg viewBox="0 0 340 190" role="img" aria-label={chart.type+' chart'}><line x1="28" y1="155" x2="330" y2="155" stroke="currentColor" opacity=".45"/>{chart.type==='Line'?chart.series.map((series,si)=>{const max=Math.max(1,...series.values.map(Math.abs));const points=series.values.map((v,i)=>(35+i*(285/Math.max(1,series.values.length-1))+','+(150-v/max*115)).join(' ')).join(' ');return <g key={si}><polyline points={points} fill="none" stroke={['#2563eb','#16a34a','#dc7a20'][si%3]} strokeWidth="3"/>{series.values.map((v,i)=><circle key={i} cx={35+i*(285/Math.max(1,series.values.length-1))} cy={150-v/max*115} r="3.5" fill={['#2563eb','#16a34a','#dc7a20'][si%3]}/>)}</g>}):chart.series.flatMap((series,si)=>series.values.map((v,i)=>{const max=Math.max(1,...chart.series.flatMap(x=>x.values).map(Math.abs));const group=285/Math.max(1,chart.labels.length);const width=Math.max(5,Math.min(28,group/(chart.series.length+1)));const height=Math.abs(v)/max*112;return <rect key={si+'-'+i} x={35+i*group+si*width} y={150-height} width={width-3} height={height} rx="2" fill={['#2563eb','#16a34a','#dc7a20'][si%3]}/>;}))}{chart.labels.map((label,i)=><text key={i} x={35+i*(285/Math.max(1,chart.labels.length))} y="174" fontSize="9" fill="currentColor">{String(label).slice(0,8)}</text>)}</svg></div><small>Source: {chart.range}</small></section>}
  {pivot&&<section className="excelDataPanel"><div className="excelPanelHead"><b>PivotTable রিপোর্ট</b><div><button onClick={()=>exportPivotSheet()}>Export → নতুন Sheet</button><button onClick={()=>setPivot(null)}>×</button></div></div><div className="pivotTable"><table><thead><tr>{pivot.headers.map((h,i)=><th key={i}>{h}</th>)}<th>{pivot.measureHeader||'Count'}</th></tr></thead><tbody>{pivot.rows.map((row,i)=><tr key={i} style={row.rowType==='subtotal'?{fontWeight:700,background:'#e8f0fe'}:{}}>{row.values.map((v,j)=><td key={j}>{v}</td>)}<th>{Number(row.result??row.aggregate??0).toLocaleString(undefined,{maximumFractionDigits:2})}</th></tr>)}</tbody><tfoot><tr>{pivot.headers.map((h,i)=><th key={i}>{i===0?'Grand Total':''}</th>)}<th>{Number(pivot.grandTotal).toLocaleString(undefined,{maximumFractionDigits:2})}</th></tr></tfoot></table></div><small>Rows: {pivot.rowFields.join(', ')||'None'} • Columns: {pivot.columnFields.join(', ')||'None'} • Source: {pivot.range} • প্রতিটি Field আলাদা কলামে দেখানো হয়েছে।</small></section>}
  <div className="sheetTabs"><button className="sheetNav" onClick={()=>notify('Previous sheet')}>‹</button><button className="sheetTab activeSheet" onClick={renameSheet} onContextMenu={e=>{e.preventDefault();deleteSheet()}} title="Tap to rename; long/right click to delete">{sheet}</button>{Object.keys({...sheetStore,[sheet]:{cells,cellStyles}}).filter(n=>n!==sheet).map(n=><button key={n} className="sheetTab" onClick={()=>switchSheet(n)}>{n}</button>)}<button className="addSheet" onClick={addSheet}><Plus size={17}/></button><button className="sheetNav" onClick={deleteSheet} title="Delete current sheet"><X size={15}/></button><span className="sheetStatus"><button onClick={()=>setSheetZoom(v=>Math.max(.6,Number((v-.1).toFixed(2))))}>−</button> {Math.round(sheetZoom*100)}% <button onClick={()=>setSheetZoom(v=>Math.min(2.5,Number((v+.1).toFixed(2))))}>+</button> Ready</span></div>
  <button className="ribbonMenuButton" onClick={()=>setMenuOpen(v=>!v)} aria-label="Open ribbon menu"><Grid2X2 size={24}/></button>
  {menuOpen&&<><div className="menuScrim" onClick={()=>setMenuOpen(false)}></div><div className="ribbonMenuPanel"><div className="menuPanelTitle">Workbook Menu <button onClick={()=>setMenuOpen(false)}><X size={18}/></button></div>{menus.map(([name,Icon])=><button key={name} className={'menuChoice '+(active===name?'chosen':'')} onClick={()=>{setActive(name);setMenuOpen(false)}}><Icon size={19}/><span>{name}</span>{active===name&&<Check size={16}/>}</button>)}</div></>}
  {isImporting&&<div className="excelToast">বড় Excel ফাইল প্রসেস হচ্ছে… অনুগ্রহ করে অপেক্ষা করুন</div>}
  <input id="excelFileInput" type="file" hidden accept=".xls,.xlsx,.csv" onChange={e=>{if(e.target.files?.[0])importWorkbook(e.target.files[0]);e.target.value=''}}/>
  {notice&&<div className="excelToast">{notice}</div>}
 </div>
}
function App(){
 const [tab,setTab]=React.useState('Home');
 const [excel,setExcel]=React.useState(false);
 const cards=[{name:'Word',desc:'Create and edit documents',icon:FileText,color:'#2463c7'},{name:'Excel',desc:'Spreadsheets and calculations',icon:Table2,color:'#16834b'},{name:'PowerPoint',desc:'Presentations and slides',icon:Presentation,color:'#d66a28'},{name:'PDF Reader',desc:'View PDF documents',icon:File,color:'#c83b45'}];
 if(excel)return <ExcelEditor onBack={()=>setExcel(false)}/>;
 return <div className="app"><header><div className="brand"><div className="brandIcon">S</div><div><b>Smart Office</b><small>Work smarter, anywhere</small></div></div><button className="round"><Settings size={20}/></button></header><main><section className="welcome"><div><span className="eyebrow">YOUR PRODUCTIVITY SPACE</span><h1>Everything you need<br/>to get work done.</h1><p>Create, edit and manage your office files in one place.</p></div><div className="welcomeMark">SO</div></section><div className="sectionTitle"><h2>Office tools</h2><span>4 tools</span></div><div className="grid">{cards.map(c=><button className="tool" key={c.name} onClick={()=>c.name==='Excel'?setExcel(true):setTab(c.name)}><div className="toolIcon" style={{background:c.color}}><c.icon size={25}/></div><b>{c.name}</b><small>{c.desc}</small><span className="arrow">›</span></button>)}</div><div className="sectionTitle recentTitle"><h2>Quick access</h2><button className="textBtn" onClick={()=>setTab('Files')}>See all</button></div><div className="quick"><div className="quickIcon"><FolderOpen/></div><div><b>Open a document</b><small>Browse files on your device</small></div><button className="openBtn" onClick={()=>document.getElementById('fileInput').click()}>Open</button><input id="fileInput" type="file" hidden accept=".doc,.docx,.xls,.xlsx,.ppt,.pptx,.pdf" onChange={e=>e.target.files?.[0]&&alert('Selected: '+e.target.files[0].name)}/></div><div className="notice"><ShieldCheck size={18}/><span>Your files stay on your device. Offline-ready design.</span></div></main><nav>{[['Home',FolderOpen],['Recent',Clock3],['Files',FileText]].map(([n,I])=><button className={tab===n?'active':''} onClick={()=>setTab(n)}><I size={20}/><span>{n}</span></button>)}</nav>{tab!=='Home'&&<div className="toast" onClick={()=>setTab('Home')}>{tab} section is coming soon. Tap to return Home.</div>}</div>
}
createRoot(document.getElementById('root')).render(<App/>);