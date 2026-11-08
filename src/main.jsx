import React from 'react';
import L from 'leaflet';
import 'leaflet/dist/leaflet.css';
import {registerPlugin,Capacitor} from '@capacitor/core';
const BikeTracking=registerPlugin('BikeTracking');
import * as XLSX from 'xlsx';
import {createRoot} from 'react-dom/client';
import {FileText,Table2,Presentation,File,FolderOpen,Clock3,Settings,ShieldCheck,Grid2X2,Home,Folder,Eye,Plus,Database,ClipboardCheck,Save,Printer,Undo2,Redo2,Copy,Scissors,Clipboard,Paintbrush,Type,Bold,Italic,Underline,AlignLeft,AlignCenter,AlignRight,WrapText,Merge,Filter,ArrowDownUp,Search,Check,MessageSquare,Lock,BarChart3,Image,Link,Calculator,Percent,Coins,CalendarDays,SortAsc,FileSpreadsheet,X,ChevronDown,Bike,MapPin,Gauge,Timer,Fuel,Route,Download,Play,Square,Navigation,TrendingUp,Trash2} from 'lucide-react';
import './style.css';

const ribbonData = {
 Home: [
  {group:'Clipboard',items:[['Undo',Undo2],['Redo',Redo2],['Paste',Clipboard],['Cut',Scissors],['Copy',Copy],['Format Painter',Paintbrush]]},
  {group:'Font',items:[['Font Name',Type],['Font Size',Type],['Increase Font Size',Plus],['Decrease Font Size',ChevronDown],['Bold',Bold],['Italic',Italic],['Underline',Underline],['Font Color',Type],['Fill Color',Paintbrush],['Borders',Grid2X2]]},
  {group:'Alignment',items:[['Align Left',AlignLeft],['Center',AlignCenter],['Align Right',AlignRight],['Top Align',AlignLeft],['Middle Align',AlignCenter],['Bottom Align',AlignRight],['Wrap Text',WrapText],['Merge Cells',Merge],['Unmerge Cells',X],['Increase Indent',ChevronDown],['Decrease Indent',ChevronDown]]},
  {group:'Number',items:[['General',Calculator],['Number',Calculator],['Currency',Coins],['Accounting',Coins],['Percentage',Percent],['Date',CalendarDays],['Time',Clock3],['Fraction',Calculator],['Scientific',Calculator],['Text',Type],['Increase Decimal',Plus],['Decrease Decimal',ChevronDown]]},
  {group:'Cells',items:[['Insert Rows',Plus],['Insert Columns',Plus],['Delete Rows',X],['Delete Columns',X],['Row Height',Grid2X2],['Column Width',Grid2X2],['AutoFit Row Height',ChevronDown],['AutoFit Column Width',ChevronDown],['Hide Rows',Eye],['Unhide Rows',Eye],['Hide Columns',Eye],['Unhide Columns',Eye]]},
  {group:'Editing',items:[['AutoSum',Calculator],['Fill Down',ChevronDown],['Fill Right',ChevronDown],['Clear Contents',X],['Clear Formats',X],['Clear All',X],['Find',Search],['Replace',Search],['Sort A to Z',SortAsc],['Sort Z to A',ArrowDownUp],['Filter',Filter],['Select All',Table2]]}
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
 Formulas: [
  {group:'Function Library',items:[['Insert Function',Calculator],['AutoSum',Calculator],['SUM',Calculator],['AVERAGE',Calculator],['COUNT',Calculator],['MAX',Calculator],['MIN',Calculator],['IF',Calculator]]},
  {group:'Formula Auditing',items:[['Show Precedents',Search],['Show Dependents',Search],['Evaluate Formula',Check],['Error Checking',ShieldCheck]]},
  {group:'Calculation',items:[['Calculate Now',Redo2]]}
 ],
 PageLayout: [
  {group:'Page Setup',items:[['Margins',FileText],['Orientation',FileText],['Paper Size',FileText],['Print Area',Printer]]},
  {group:'Scale to Fit',items:[['Fit to Width',Grid2X2],['Fit to Height',Grid2X2],['Scale',Percent]]},
  {group:'Sheet Options',items:[['Print Gridlines',Grid2X2],['Print Headings',Table2]]},
  {group:'Arrange',items:[['Page Order',ArrowDownUp],['Page Setup Reset',Undo2]]}
 ],
 Review: [
  {group:'Proofing',items:[['Spelling',Check],['Check Workbook',ClipboardCheck],['Workbook Statistics',FileSpreadsheet]]},
  {group:'Comments',items:[['New Comment',MessageSquare],['Show Comments',MessageSquare],['Notes',FileText],['Delete All Comments',X]]},
  {group:'Protection',items:[['Protect Sheet',Lock],['Protect Workbook',ShieldCheck],['Allow Edit Ranges',Check]]}
 ]
};
function ColumnsIcon(){return <Table2/>}
const menus=[['Home',Home],['File',Folder],['View',Eye],['Insert',Plus],['Data',Database],['Review',ClipboardCheck],['Formulas',Calculator],['Page Layout',FileText]];
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
 let tokens=tokenize(source);let index=0;
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
   case 'IFERROR':return (typeof a==='number'&&!Number.isFinite(a))||['#VALUE!','#DIV/0!','#REF!','#NAME?','#N/A','#CYCLE!'].includes(a)?(args[1]??''):a;
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
   }default:throw Error('Unknown function '+name);
  }
 };
 const criteriaMatch=(value,criteria)=>{
  const s=textValue(criteria),m=s.match(/^(>=|<=|<>|>|<|=)(.*)$/);
  if(!m)return textValue(value).toLowerCase()===s.toLowerCase();
  const left=numeric(value),right=numeric(m[2]);
  const x=left!==null?left:textValue(value),y=right!==null?right:m[2];
  switch(m[1]){case '>':return x>y;case '<':return x<y;case '>=':return x>=y;case '<=':return x<=y;case '<>':return x!==y;default:return x===y}
 };
 const readLazyArgs=()=>{
  const args=[];if(peek()===')'){take();return args}
  let start=index,depth=0;
  while(index<tokens.length){
   const token=tokens[index];
   if(token==='(')depth++;
   else if(token===')'){
    if(depth===0){args.push(tokens.slice(start,index));index++;return args}
    depth--;
   }else if(token===','&&depth===0){args.push(tokens.slice(start,index));index++;start=index;continue}
   index++;
  }
  throw Error('Missing )');
 };
 const evaluateTokens=part=>{
  const previousTokens=tokens,previousIndex=index;tokens=part;index=0;
  try{const value=parseCompare();if(index!==tokens.length)throw Error('Unexpected token');return value}
  finally{tokens=previousTokens;index=previousIndex}
 };
 const isFormulaError=value=>typeof value==='string'&&['#VALUE!','#DIV/0!','#REF!','#NAME?','#N/A','#CYCLE!'].includes(value);
 const primary=()=>{
  const t=take();if(t===undefined)throw Error('Missing value');
  if(t==='('){const v=parseCompare();if(take()!==')')throw Error('Missing )');return v}
  if(t[0]==='"')return t.slice(1,-1).replace(/""/g,'"');
  if(/^\d/.test(t)||/^\.\d/.test(t))return Number(t);
  if(isCell(t)){
   if(peek()===':'){take();const second=take();if(!isCell(second))throw Error('Bad range');return rangeValues(t,second)}
   return cellValue(t);
  }
  if(/^[A-Z_][A-Z0-9_.]*$/i.test(t)&&peek()==='('){
   take();const name=t.toUpperCase();
   if(name==='IF'||name==='IFERROR'){
    const args=readLazyArgs();
    if(name==='IF'){
     const condition=args.length?evaluateTokens(args[0]):false;
     const branch=truth(condition)?args[1]:args[2];
     return branch?.length?evaluateTokens(branch):(truth(condition)?true:false);
    }
    let value;
    try{value=args[0]?.length?evaluateTokens(args[0]):''}catch(error){value='#VALUE!'}
    if((typeof value==='number'&&!Number.isFinite(value))||isFormulaError(value))return args[1]?.length?evaluateTokens(args[1]):'';
    return value;
   }
   return call(name,parseArgs());
  }
  throw Error('Unexpected value');
 };
 const unary=()=>{if(peek()==='+'){take();return +unary()}if(peek()==='-'){take();return -unary()}let v=primary();while(peek()==='%'){take();v=Number(v)/100}return v};
 const power=()=>{let v=unary();if(peek()==='^'){take();v=Math.pow(Number(v),Number(power()))}return v};
 const mul=()=>{let v=power();while(peek()==='*'||peek()==='/'){const op=take(),r=power();v=op==='*'?Number(v)*Number(r):(Number(r)===0?'#DIV/0!':Number(v)/Number(r))}return v};
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
 React.useEffect(()=>{let live=true;try{const req=indexedDB.open('SmartOfficeWorkbookDB',1);req.onupgradeneeded=()=>{if(!req.result.objectStoreNames.contains('workbooks'))req.result.createObjectStore('workbooks')};req.onsuccess=()=>{const db=req.result,tx=db.transaction('workbooks','readonly'),get=tx.objectStore('workbooks').get('current');get.onsuccess=()=>{if(live){const disk=get.result;const local=initialWorkbook;const diskTime=Date.parse(disk?.savedAt||'')||0;const localTime=Date.parse(local?.savedAt||'')||0;const d=(local&&localTime>diskTime)?local:(disk||local);if(d){setSheetStore(d.sheets||{[d.sheet||'Sheet1']:{cells:d.cells||{},cellStyles:d.cellStyles||{}}});setCells(d.cells||{});setCellStyles(d.cellStyles||{});setSheet(d.sheet||'Sheet1');setWorkbookName(d.workbookName||'Book1');setChart(d.chart||null);setPivot(d.pivot||null);setRowHeights(d.rowHeights||{});setColumnWidths(d.columnWidths||{});setHiddenRows(d.hiddenRows||[]);setHiddenCols(d.hiddenCols||[]);setMergedRanges(d.mergedRanges||[]);setFilterConfig(d.filterConfig||null);setDataValidations(d.dataValidations||{});setHeaderFooter(d.headerFooter||{header:'',footer:''});setPageSetup(d.pageSetup||{margins:'normal',orientation:'portrait',paperSize:'A4',printArea:'',fitWidth:1,fitHeight:0,scale:100,printGridlines:false,printHeadings:false,pageOrder:'down'});setSheetProtected(!!d.sheetProtected);setSheetProtectionPassword(d.sheetProtectionPassword||'');setWorkbookProtected(!!d.workbookProtected);setWorkbookProtectionPassword(d.workbookProtectionPassword||'');setAllowedEditRanges(d.allowedEditRanges||[]);setFilter(!!d.filterConfig)}}if(live)setIsRestoring(false);db.close()};get.onerror=()=>{if(live)setIsRestoring(false);db.close()}};req.onerror=()=>{if(live)setIsRestoring(false)}}catch{setIsRestoring(false)}return()=>{live=false}},[]);
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
 const [styleHistory,setStyleHistory]=React.useState(()=>[initialWorkbook?.cellStyles||{}]);
 const [historyIndex,setHistoryIndex]=React.useState(0);
 const [clipboard,setClipboard]=React.useState(null);
 const [sheet,setSheet]=React.useState(()=>initialWorkbook?.sheet||'Sheet1');
 const [workbookName,setWorkbookName]=React.useState(()=>initialWorkbook?.workbookName||'Book1');
 const [headerFooter,setHeaderFooter]=React.useState(()=>initialWorkbook?.headerFooter||{header:'',footer:''});
 const [pageSetup,setPageSetup]=React.useState(()=>initialWorkbook?.pageSetup||{margins:'normal',orientation:'portrait',paperSize:'A4',printArea:'',fitWidth:1,fitHeight:0,scale:100,printGridlines:false,printHeadings:false,pageOrder:'down'});
 const [sheetProtected,setSheetProtected]=React.useState(()=>!!initialWorkbook?.sheetProtected);
 const [sheetProtectionPassword,setSheetProtectionPassword]=React.useState(()=>initialWorkbook?.sheetProtectionPassword||'');
 const [workbookProtected,setWorkbookProtected]=React.useState(()=>!!initialWorkbook?.workbookProtected);
 const [workbookProtectionPassword,setWorkbookProtectionPassword]=React.useState(()=>initialWorkbook?.workbookProtectionPassword||'');
 const [allowedEditRanges,setAllowedEditRanges]=React.useState(()=>initialWorkbook?.allowedEditRanges||[]);
 const [spellingEnabled,setSpellingEnabled]=React.useState(false);
 const [spellingLanguage,setSpellingLanguage]=React.useState(()=>initialWorkbook?.spellingLanguage||'en-US');
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
 const [showHeadings,setShowHeadings]=React.useState(true);
 const [showRuler,setShowRuler]=React.useState(false);
 const [freezePanes,setFreezePanes]=React.useState(false);
 const [viewMode,setViewMode]=React.useState('normal');
 const [pageBreaks,setPageBreaks]=React.useState(false);
 const [filter,setFilter]=React.useState(false);
 const [rowHeights,setRowHeights]=React.useState({});
 const [columnWidths,setColumnWidths]=React.useState({});
 const [hiddenRows,setHiddenRows]=React.useState([]);
 const [hiddenCols,setHiddenCols]=React.useState([]);
 const [mergedRanges,setMergedRanges]=React.useState(()=>initialWorkbook?.mergedRanges||[]);
 const [mergedHistory,setMergedHistory]=React.useState(()=>[initialWorkbook?.mergedRanges||[]]);
 const [filterConfig,setFilterConfig]=React.useState(null);
 const [dataValidations,setDataValidations]=React.useState(()=>initialWorkbook?.dataValidations||{});
 const colName=n=>{let s='';while(n>0){let r=(n-1)%26;s=String.fromCharCode(65+r)+s;n=Math.floor((n-1)/26)}return s};
 const notify=t=>{setNotice(t);setTimeout(()=>setNotice(''),1800)};
 const [confirmExit,setConfirmExit]=React.useState(false);
 const exitHandler=React.useRef(null);
 exitHandler.current=()=>{setConfirmExit(true)};
 const requestExit=()=>exitHandler.current?.();
 React.useEffect(()=>{
  let active=true;
  const pushBackState=()=>{try{window.history.pushState({smartExcelBack:true},'')}catch{}};
  pushBackState();
  const onPopState=()=>{if(!active)return;setConfirmExit(true);pushBackState()};
  const onKeyDown=e=>{if(e.key==='Escape'||e.key==='BrowserBack'||e.key==='Backspace'&& !['INPUT','TEXTAREA'].includes(document.activeElement?.tagName)){e.preventDefault();setConfirmExit(true)}};
  window.addEventListener('popstate',onPopState);
  window.addEventListener('keydown',onKeyDown);
  let nativeBackListener=null;
  import('@capacitor/app').then(({App})=>App.addListener('backButton',()=>{if(active)setConfirmExit(true)})).then(listener=>{nativeBackListener=listener}).catch(()=>{});
  return()=>{active=false;window.removeEventListener('popstate',onPopState);window.removeEventListener('keydown',onKeyDown);nativeBackListener?.remove?.()};
 },[]);

 React.useEffect(()=>{if(isRestoring)return;const timer=setTimeout(()=>{const all={...sheetStore,[sheet]:{cells,cellStyles}};setSheetStore(all);const data={sheets:all,cells,cellStyles,sheet,workbookName,chart,pivot,rowHeights,columnWidths,hiddenRows,hiddenCols,mergedRanges,filterConfig,dataValidations,headerFooter,pageSetup,sheetProtected,sheetProtectionPassword,workbookProtected,workbookProtectionPassword,allowedEditRanges,spellingLanguage,savedAt:new Date().toISOString()};try{localStorage.setItem('smartoffice-workbook',JSON.stringify(data))}catch{};try{const request=indexedDB.open('SmartOfficeWorkbookDB',1);request.onupgradeneeded=()=>{if(!request.result.objectStoreNames.contains('workbooks'))request.result.createObjectStore('workbooks')};request.onsuccess=()=>{const db=request.result,tx=db.transaction('workbooks','readwrite');tx.objectStore('workbooks').put(data,'current');tx.oncomplete=()=>db.close();tx.onerror=()=>{db.close();notify('Workbook auto-save failed. Please export a backup')}};request.onerror=()=>notify('Device storage unavailable. Please export a backup')}catch{notify('Device storage unavailable. Please export a backup')}},500);return()=>clearTimeout(timer)},[cells,cellStyles,sheet,workbookName,chart,pivot,rowHeights,columnWidths,hiddenRows,hiddenCols,mergedRanges,filterConfig,dataValidations,headerFooter,pageSetup,sheetProtected,sheetProtectionPassword,workbookProtected,workbookProtectionPassword,allowedEditRanges,spellingLanguage,isRestoring]);
 const selectCell=(id)=>{setSelected(id);setFormula(cells[id]??'');setSelectionBand(null)};
 const selectWholeColumn=(col)=>{const id=positionToCell(col,1);setSelected(id);setFormula(cells[id]??'');setSelectionBand({axis:'col',index:col});notify(colName(col)+' কলাম সম্পূর্ণ সিলেক্ট হয়েছে');};
 const selectWholeRow=(row)=>{const id=positionToCell(1,row);setSelected(id);setFormula(cells[id]??'');setSelectionBand({axis:'row',index:row});notify(row+' নম্বর Row সম্পূর্ণ সিলেক্ট হয়েছে');};
 const recordHistory=(nextCells,nextStyles=cellStyles,nextMerged=mergedRanges)=>{
  setCellHistory(h=>[...h.slice(0,historyIndex+1),nextCells].slice(-101));
  setStyleHistory(h=>[...h.slice(0,historyIndex+1),nextStyles].slice(-101));
  setMergedHistory(h=>[...h.slice(0,historyIndex+1),nextMerged].slice(-101));
  setHistoryIndex(i=>Math.min(i+1,100));
 };
 const isCellEditable=(id)=>{if(!sheetProtected)return true;const p=cellToPosition(id);return !!p&&allowedEditRanges.some(range=>{const m=String(range).toUpperCase().match(/^\$?([A-Z]+)\$?([1-9][0-9]*):\$?([A-Z]+)\$?([1-9][0-9]*)$/);if(!m)return false;const x=cellToPosition(m[1]+m[2]),y=cellToPosition(m[3]+m[4]);return x&&y&&p.row>=Math.min(x.row,y.row)&&p.row<=Math.max(x.row,y.row)&&p.col>=Math.min(x.col,y.col)&&p.col<=Math.max(x.col,y.col)});};
 const updateCell=(id,value)=>{
  if(!isCellEditable(id)){notify('Sheet protected — এই সেলে সম্পাদনার অনুমতি নেই');return}
  const rule=dataValidations[id];
  if(rule){
   const raw=String(value??'');
   if(raw.trim()!==''){
    let valid=true,message='';
    if(rule.type==='list'){valid=rule.values.includes(raw.trim());message='এই সেলে অনুমোদিত তালিকার একটি মান লিখুন: '+rule.values.join(', ')}
    else if(rule.type==='whole'||rule.type==='decimal'){const n=Number(raw);valid=Number.isFinite(n)&&(rule.type!=='whole'||Number.isInteger(n))&&n>=rule.min&&n<=rule.max;message='মান '+rule.min+' থেকে '+rule.max+(rule.type==='whole'?' এর মধ্যে পূর্ণসংখ্যা':' এর মধ্যে সংখ্যা')+' হতে হবে'}
    else if(rule.type==='date'){const time=Date.parse(raw),min=Date.parse(rule.min),max=Date.parse(rule.max);valid=Number.isFinite(time)&&time>=min&&time<=max;message='তারিখ '+rule.min+' থেকে '+rule.max+' এর মধ্যে হতে হবে'}
    else if(rule.type==='textLength'){const length=raw.length;valid=length>=rule.min&&length<=rule.max;message='Text length '+rule.min+' থেকে '+rule.max+' অক্ষরের মধ্যে হতে হবে'}
    if(!valid){notify(message);return}
   }
  }
  const next={...cells,[id]:value};
  setCells(next);recordHistory(next);
 };
 const undo=()=>{if(historyIndex<=0){notify('Nothing to undo');return}const i=historyIndex-1;setHistoryIndex(i);setCells(cellHistory[i]||{});setCellStyles(styleHistory[i]||{});setMergedRanges(mergedHistory[i]||[]);setFormula((cellHistory[i]||{})[selected]??'');notify('Undo complete')};
 const redo=()=>{if(historyIndex>=cellHistory.length-1){notify('Nothing to redo');return}const i=historyIndex+1;setHistoryIndex(i);setCells(cellHistory[i]||{});setCellStyles(styleHistory[i]||{});setFormula((cellHistory[i]||{})[selected]??'');notify('Redo complete')};
 const getSelectionBounds=()=>{
  const p=cellToPosition(selected)||{row:1,col:1};
  let minRow=p.row,maxRow=p.row,minCol=p.col,maxCol=p.col;
  const populated=Object.keys(cells).map(k=>cellToPosition(k)).filter(Boolean);
  if(selectionBand?.axis==='all'){
   minRow=1;minCol=1;
   maxRow=Math.max(visibleRows,...populated.map(q=>q.row),1);
   maxCol=Math.max(visibleCols,...populated.map(q=>q.col),1);
  }else if(selectionBand?.axis==='col'){
   minRow=1;maxRow=Math.max(visibleRows,...populated.map(q=>q.row),1);
   minCol=maxCol=Number(selectionBand.index)||p.col;
  }else if(selectionBand?.axis==='row'){
   minCol=1;maxCol=Math.max(visibleCols,...populated.map(q=>q.col),1);
   minRow=maxRow=Number(selectionBand.index)||p.row;
  }else if(selectionBand?.axis==='range'){
   const a=cellToPosition(selectionBand.start),b=cellToPosition(selectionBand.end);
   if(a&&b){minRow=Math.min(a.row,b.row);maxRow=Math.max(a.row,b.row);minCol=Math.min(a.col,b.col);maxCol=Math.max(a.col,b.col)}
  }
  return {minRow,maxRow,minCol,maxCol};
 };
 const copyCell=async(cut=false)=>{
  const b=getSelectionBounds(),matrix=[],stylesMatrix=[];
  for(let r=b.minRow;r<=b.maxRow;r++){
   const line=[],styleLine=[];
   for(let c=b.minCol;c<=b.maxCol;c++){const id=positionToCell(c,r);line.push(cells[id]??'');styleLine.push(cellStyles[id]||null)}
   matrix.push(line);stylesMatrix.push(styleLine);
  }
  setClipboard({kind:'range',matrix,stylesMatrix,bounds:b});
  const tsv=matrix.map(row=>row.map(v=>String(v).replace(/\t/g,' ')).join('\t')).join('\n');
  try{await navigator.clipboard?.writeText(tsv)}catch{}
  if(cut){
   const next={...cells},nextStyles={...cellStyles};
   for(let r=b.minRow;r<=b.maxRow;r++)for(let c=b.minCol;c<=b.maxCol;c++){const id=positionToCell(c,r);delete next[id];delete nextStyles[id]}
   setCells(next);setCellStyles(nextStyles);
   recordHistory(next,nextStyles);
  }
  notify(cut?'নির্বাচিত অংশ Cut হয়েছে':'নির্বাচিত অংশ Copy হয়েছে');
 };
 const pasteMatrix=(matrix,stylesMatrix=null)=>{
  if(!matrix?.length)return;
  const start=cellToPosition(selected)||{row:1,col:1},next={...cells},nextStyles={...cellStyles};
  let skipped=0,invalid=0;
  matrix.forEach((line,ri)=>line.forEach((value,ci)=>{
   const id=positionToCell(start.col+ci,start.row+ri);
   if(!isCellEditable(id)){skipped++;return}
   const rule=dataValidations[id],raw=String(value??'').trim();
   if(rule&&raw!==''){
    const n=Number(raw);
    const valid=rule.type==='list'?rule.values.includes(raw):((rule.type==='whole'||rule.type==='decimal')&&Number.isFinite(n)&&(rule.type!=='whole'||Number.isInteger(n))&&n>=rule.min&&n<=rule.max);
    if(!valid){invalid++;return}
   }
   next[id]=value;
   const copiedStyle=stylesMatrix?.[ri]?.[ci];
   if(copiedStyle)nextStyles[id]={...copiedStyle};else if(stylesMatrix)delete nextStyles[id];
  }));
  setCells(next);if(stylesMatrix)setCellStyles(nextStyles);
  recordHistory(next,nextStyles);
  const last=positionToCell(start.col+(matrix[0]?.length||1)-1,start.row+matrix.length-1);
  setSelectionBand({axis:'range',start:selected,end:last});setFormula(next[selected]??'');notify('Paste সম্পন্ন: '+(matrix.reduce((n,line)=>n+line.length,0)-skipped-invalid)+'টি; Protected '+skipped+', Validation rejected '+invalid);
 };
 const pasteCell=()=>{
  if(!clipboard){notify('আগে Copy করুন অথবা বাইরের লেখা Paste করুন');return}
  if(clipboard.kind==='style'){const nextStyles={...cellStyles,[selected]:{...clipboard.style}};setCellStyles(nextStyles);recordHistory(cells,nextStyles);notify('Format applied');return}
  if(clipboard.kind==='range'){pasteMatrix(clipboard.matrix,clipboard.stylesMatrix||null);return}
  pasteMatrix([[clipboard.value]]);
 };
 const shiftCells=(axis,at,remove=false,count=1)=>{
  const amount=Math.max(1,Math.floor(Number(count)||1));
  const remap=(source)=>{const out={};for(const [key,value] of Object.entries(source)){const p=cellToPosition(key);if(!p)continue;const n=axis==='row'?p.row:p.col;if(remove&&n>=at&&n<at+amount)continue;const moved=remove?(n>=at+amount?n-amount:n):(n>=at?n+amount:n);const dest=axis==='row'?positionToCell(p.col,moved):positionToCell(moved,p.row);out[dest]=value;}return out;};
  const next=remap(cells),styles=remap(cellStyles);setCells(next);setCellStyles(styles);recordHistory(next,styles);setFormula(next[selected]??'');notify((remove?'Deleted ':'Inserted ')+(axis==='row'?'row':'column'));
 };
 const updateStyle=(patch)=>{const next={...cellStyles,[selected]:{...(cellStyles[selected]||{}),...patch}};setCellStyles(next);recordHistory(cells,next)};
 const applyStyleToSelection=(patch)=>{
  const b=getSelectionBounds(),next={...cellStyles};
  for(let r=b.minRow;r<=b.maxRow;r++)for(let c=b.minCol;c<=b.maxCol;c++){const id=positionToCell(c,r);next[id]={...(next[id]||{}),...patch};}
  setCellStyles(next);recordHistory(cells,next);notify('নির্বাচিত সেলের ফরম্যাট পরিবর্তন হয়েছে');
 };
 const clearSelection=(mode)=>{
  const b=getSelectionBounds(),nextCells={...cells},nextStyles={...cellStyles};
  let skipped=0;for(let r=b.minRow;r<=b.maxRow;r++)for(let c=b.minCol;c<=b.maxCol;c++){const id=positionToCell(c,r);if(!isCellEditable(id)){skipped++;continue}if(mode!=='formats')delete nextCells[id];if(mode!=='contents')delete nextStyles[id];}
  setCells(nextCells);setCellStyles(nextStyles);recordHistory(nextCells,nextStyles);setFormula(nextCells[selected]??'');notify((mode==='contents'?'নির্বাচিত সেলের লেখা মুছে ফেলা হয়েছে':mode==='formats'?'নির্বাচিত সেলের ফরম্যাট মুছে ফেলা হয়েছে':'নির্বাচিত সেল সম্পূর্ণ পরিষ্কার হয়েছে')+(skipped?' ('+skipped+' protected cell skipped)':''));
 };
 const setFormat=(format)=>{
  const formats={General:null,Number:'number',Currency:'currency',Accounting:'accounting',Percentage:'percent',Date:'date',Time:'time',Fraction:'fraction',Scientific:'scientific',Text:'text'};
  const value=formats[format];applyStyleToSelection({numberFormat:value||'general'});
 };
 const fillSelection=(axis)=>{
  const b=getSelectionBounds(),origin=cellToPosition(selected)||{row:b.minRow,col:b.minCol},source=cells[selected]??'',next={...cells};
  let skipped=0;if(axis==='down'){for(let r=origin.row+1;r<=b.maxRow;r++){const id=positionToCell(origin.col,r);if(isCellEditable(id))next[id]=source;else skipped++;}}
  else {for(let c=origin.col+1;c<=b.maxCol;c++){const id=positionToCell(c,origin.row);if(isCellEditable(id))next[id]=source;else skipped++;}}
  setCells(next);recordHistory(next);
  notify((axis==='down'?'Fill Down সম্পন্ন':'Fill Right সম্পন্ন')+(skipped?' ('+skipped+' protected cell skipped)':''));
 };
 const sortSelection=(direction)=>{
  const b=getSelectionBounds(),p=cellToPosition(selected)||{row:b.minRow,col:b.minCol};
  const rows=[];for(let r=b.minRow;r<=b.maxRow;r++){const values=[],styles=[];for(let c=b.minCol;c<=b.maxCol;c++){const id=positionToCell(c,r);values.push(cells[id]??'');styles.push(cellStyles[id]||null);}rows.push({values,styles,original:r});}
  if(rows.length<2){notify('Sort করার জন্য অন্তত ২টি Row নির্বাচন করুন');return}
  const header=rows.shift();
  rows.sort((a,bv)=>String(a.values[p.col-b.minCol]??'').localeCompare(String(bv.values[p.col-b.minCol]??''),'en',{numeric:true,sensitivity:'base'})*(direction==='asc'?1:-1));
  const next={...cells},nextStyles={...cellStyles};
  [...rows].forEach((row,index)=>{const target=b.minRow+1+index;row.values.forEach((v,i)=>{const id=positionToCell(b.minCol+i,target);if(v===''||v===undefined)delete next[id];else next[id]=v;const st=row.styles[i];if(st)nextStyles[id]=st;else delete nextStyles[id];});});
  header.values.forEach((v,i)=>{const id=positionToCell(b.minCol+i,b.minRow);if(v===''||v===undefined)delete next[id];else next[id]=v;});
  setCells(next);setCellStyles(nextStyles);recordHistory(next,nextStyles);notify(direction==='asc'?'A to Z Sort সম্পন্ন':'Z to A Sort সম্পন্ন');
 };
 const replaceText=()=>{
  const find=window.prompt('কোন লেখা পরিবর্তন করবেন?');if(find===null||find==='')return;
  const replacement=window.prompt('নতুন লেখা লিখুন:');if(replacement===null)return;
  const next={...cells};let count=0;
  Object.entries(next).forEach(([id,value])=>{if(typeof value==='string'&&!value.startsWith('=')&&value.includes(find)){next[id]=value.split(find).join(replacement);count++;}});
  setCells(next);recordHistory(next);notify(count+'টি সেলে Replace সম্পন্ন');
 };
 const mergeSelection=(unmerge=false)=>{
  const b=getSelectionBounds();
  if(unmerge){
   const nextMerged=mergedRanges.filter(m=>!(m.startRow<=b.maxRow&&m.endRow>=b.minRow&&m.startCol<=b.maxCol&&m.endCol>=b.minCol));
   if(nextMerged.length===mergedRanges.length){notify('এই Selection-এ Merge পাওয়া যায়নি');return}
   setMergedRanges(nextMerged);recordHistory(cells,cellStyles,nextMerged);notify('Merge সরানো হয়েছে');return;
  }
  if(b.minRow===b.maxRow&&b.minCol===b.maxCol){notify('Merge করার জন্য একাধিক সেল নির্বাচন করুন');return}
  const overlaps=mergedRanges.some(m=>m.startRow<=b.maxRow&&m.endRow>=b.minRow&&m.startCol<=b.maxCol&&m.endCol>=b.minCol);
  if(overlaps){notify('নির্বাচিত অংশে আগে থেকেই Merge আছে');return}
  const nextMerged=[...mergedRanges,{startRow:b.minRow,endRow:b.maxRow,startCol:b.minCol,endCol:b.maxCol}];
  const anchor=positionToCell(b.minCol,b.minRow),next={...cells};
  for(let r=b.minRow;r<=b.maxRow;r++)for(let c=b.minCol;c<=b.maxCol;c++){const id=positionToCell(c,r);if(id!==anchor)delete next[id];}
  setCells(next);setMergedRanges(nextMerged);recordHistory(next,cellStyles,nextMerged);notify('নির্বাচিত সেল Merge হয়েছে');
 };
 const applyFilter=()=>{
  const p=cellToPosition(selected)||{row:1,col:1};
  const maxRow=Math.max(1,...Object.keys(cells).map(k=>cellToPosition(k)?.row||1));
  const mode=window.prompt('Filter ধরন লিখুন: contains / equals / starts / ends / blank / nonblank / greater / less / clear','contains');
  if(mode===null)return;
  const kind=mode.trim().toLowerCase();
  if(kind==='clear'){setFilterConfig(null);setFilter(false);notify('Filter পরিষ্কার হয়েছে');return}
  if(!['contains','equals','starts','ends','blank','nonblank','greater','less'].includes(kind)){notify('সঠিক Filter ধরন লিখুন');return}
  const term=['blank','nonblank'].includes(kind)?'':window.prompt('Filter মান লিখুন:',kind==='greater'||kind==='less'?'0':'');
  if(term===null)return;
  const target=Number(term),numeric=['greater','less'].includes(kind);
  if(numeric&&!Number.isFinite(target)){notify('Greater/Less Filter-এর জন্য সংখ্যা দিন');return}
  const keep=[];
  for(let r=p.row+1;r<=maxRow;r++){
   const raw=cells[positionToCell(p.col,r)],value=String(raw??''),lower=value.toLowerCase(),needle=String(term).toLowerCase();
   const match=kind==='blank'?(raw===undefined||raw===null||value.trim()===''):kind==='nonblank'?(raw!==undefined&&raw!==null&&value.trim()!==''):kind==='equals'?lower===needle:kind==='starts'?lower.startsWith(needle):kind==='ends'?lower.endsWith(needle):kind==='greater'?Number(raw)>target:kind==='less'?Number(raw)<target:lower.includes(needle);
   if(match)keep.push(r);
  }
  setFilterConfig({rows:keep,range:{start:p.row+1,end:maxRow}});setFilter(true);notify('Filter প্রয়োগ হয়েছে: '+keep.length+'টি Row পাওয়া গেছে');
 };
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
   setSheetStore(imported);setCells(next);setCellHistory([{},next]);setStyleHistory([{},imported[first]?.cellStyles||{}]);setMergedHistory([[],[]]);setHistoryIndex(1);setCellStyles(imported[first]?.cellStyles||{});setSheet(first);setSelected('A1');setFormula(next.A1??'');setWorkbookName(file.name.replace(/\.(xlsx|xls|csv)$/i,''));setVisibleRows(100);setVisibleCols(10);notify('সফল! '+names.length+'টি Sheet Open হয়েছে');
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
    const sheetMeta=workbook.Workbook?.Sheets||[];
    const firstVisible=names.findIndex((_,index)=>!sheetMeta[index]?.Hidden);
    const chosenIndex=Number.isInteger(activeIndex)&&names[activeIndex]&&!sheetMeta[activeIndex]?.Hidden?activeIndex:(firstVisible>=0?firstVisible:0);
    applyWorkbook(imported,names[chosenIndex]);
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
 const exportWorkbook=(format='xlsx',nameOverride=workbookName)=>{
  try{const all={...sheetStore,[sheet]:{cells,cellStyles}};const wb=XLSX.utils.book_new();for(const [sheetName,data] of Object.entries(all)){let maxRow=0,maxCol=0;for(const key of Object.keys(data.cells||{})){const p=cellToPosition(key);if(p){maxRow=Math.max(maxRow,p.row);maxCol=Math.max(maxCol,p.col)}}const aoa=Array.from({length:Math.max(1,maxRow)},()=>Array(Math.max(1,maxCol)).fill(''));const ws=XLSX.utils.aoa_to_sheet(aoa);for(const [key,value] of Object.entries(data.cells||{})){if(typeof value==='string'&&value.startsWith('='))ws[key]={t:'n',f:value.slice(1)};else ws[key]={v:value,t:typeof value==='number'?'n':'s'};}ws['!ref']=XLSX.utils.encode_range({s:{r:0,c:0},e:{r:Math.max(0,maxRow-1),c:Math.max(0,maxCol-1)}});XLSX.utils.book_append_sheet(wb,ws,sheetName.slice(0,31)||'Sheet1');}const filename=(String(nameOverride||'Book1').trim()||'Book1')+(format==='csv'?'.csv':'.xlsx');if(format==='csv'){const ws=wb.Sheets[sheet];const csv=XLSX.utils.sheet_to_csv(ws);const blob=new Blob(['\\uFEFF',csv],{type:'text/csv;charset=utf-8'});const url=URL.createObjectURL(blob);const a=document.createElement('a');a.href=url;a.download=filename;a.click();URL.revokeObjectURL(url);}else XLSX.writeFile(wb,filename);notify('Workbook downloaded');}catch(error){notify('Export failed: '+error.message)}
 };
 const exportPivotSheet=()=>{
   if(workbookProtected){notify('Workbook protected — PivotTable নতুন Sheet-এ Export করা যাবে না');return}
   if(!pivot)return;
   const base='PivotTable';let name=base,n=2;while(sheetStore[name]||name===sheet)name=base+' '+n++;
   const out={};const headers=[...pivot.headers,pivot.measureHeader||'Count'];
   headers.forEach((h,i)=>{out[positionToCell(i+1,1)]=h;});
   pivot.rows.forEach((row,r)=>{row.values.forEach((v,c)=>{out[positionToCell(c+1,r+2)]=v;});out[positionToCell(headers.length,r+2)]=row.result??row.aggregate??0;});
   const totalRow=pivot.rows.length+2;out[positionToCell(1,totalRow)]='Grand Total';out[positionToCell(headers.length,totalRow)]=pivot.grandTotal;
   const all={...sheetStore,[sheet]:{cells,cellStyles},[name]:{cells:out,cellStyles:{}}};
   setSheetStore(all);setCells(out);setCellStyles({});setCellHistory([out]);setStyleHistory([{}]);setMergedHistory([[]]);setHistoryIndex(0);setSheet(name);setSelected('A1');setFormula('');setPivot(null);notify('PivotTable নতুন Sheet-এ Export হয়েছে');
  };
  const addSheet=()=>{if(workbookProtected){notify('Workbook protected — নতুন Sheet যোগ করা যাবে না');return}let n=Object.keys(sheetStore).length+1,name='Sheet'+n;while(sheetStore[name]){n++;name='Sheet'+n;}setSheetStore(p=>({...p,[sheet]:{cells,cellStyles},[name]:{cells:{},cellStyles:{}}}));setCells({});setCellStyles({});setCellHistory([{}]);setStyleHistory([{}]);setMergedHistory([[]]);setHistoryIndex(0);setSheet(name);setSelected('A1');setFormula('');notify('New sheet created');};
 const switchSheet=(name)=>{if(name===sheet)return;const all={...sheetStore,[sheet]:{cells,cellStyles}};const next=all[name]||{cells:{},cellStyles:{}};setSheetStore(all);setCells(next.cells);setCellStyles(next.cellStyles);setCellHistory([next.cells]);setStyleHistory([next.cellStyles||{}]);setMergedHistory([[]]);setHistoryIndex(0);setSheet(name);setSelected('A1');setFormula(next.cells.A1||'');};
 const renameSheet=()=>{if(workbookProtected){notify('Workbook protected — Sheet-এর নাম পরিবর্তন করা যাবে না');return}const name=window.prompt('Enter new sheet name:',sheet);if(!name||name===sheet)return;if(name.length>31||/[\\\\/?*\\[\\]:]/.test(name)){notify('Sheet name is invalid (max 31 characters)');return}if(sheetStore[name]){notify('That sheet name already exists');return}const all={...sheetStore,[sheet]:{cells,cellStyles}};all[name]=all[sheet];delete all[sheet];setSheetStore(all);setSheet(name);notify('Sheet renamed');};
 const deleteSheet=()=>{if(workbookProtected){notify('Workbook protected — Sheet মুছে ফেলা যাবে না');return}const names=Object.keys({...sheetStore,[sheet]:{cells,cellStyles}});if(names.length<=1){notify('At least one sheet is required');return}if(!window.confirm('Delete sheet '+sheet+'?'))return;const all={...sheetStore,[sheet]:{cells,cellStyles}};delete all[sheet];const nextName=Object.keys(all)[0],next=all[nextName];setSheetStore(all);setSheet(nextName);setCells(next.cells);setCellStyles(next.cellStyles);setCellHistory([next.cells]);setStyleHistory([next.cellStyles||{}]);setHistoryIndex(0);setSelected('A1');setFormula(next.cells.A1||'');notify('Sheet deleted');};

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
    if(fieldIndexes.some(i=>i<0))throw Error('নির্বাচিত Field ডেটার Header-এর সঙ্গে মিলছে না');
    if(new Set(allFields).size!==allFields.length)throw Error('একই Field একাধিকবার নির্বাচন করা যাবে না');
    const aggregation=pivotWizard.aggregation||'count',valueField=pivotWizard.valueField||'';
    const valueIndex=valueField?headers.indexOf(valueField):-1;
    if(aggregation!=='count'&&valueIndex<0)throw Error('যোগ/গড় করার জন্য Values থেকে একটি সংখ্যার Field নির্বাচন করুন');
    const grouped=new Map();
    const cleanLabel=v=>String(v??'').replace(/\\s+/g,' ').trim()||'(blank)';
    const groupKey=values=>JSON.stringify(values.map(v=>v.toLocaleLowerCase()));
    for(const record of matrix.slice(1)){
     if(record.every(v=>String(v??'').trim()===''))continue;
     const values=fieldIndexes.map(i=>cleanLabel(record[i]));
     const key=groupKey(values);
     if(!grouped.has(key))grouped.set(key,{values,aggregate:0,recordCount:0,valueCount:0,parentKey:values[0].toLocaleLowerCase()});
     const item=grouped.get(key);item.recordCount++;
     if(aggregation==='count'){item.aggregate++;item.valueCount++;}
     else {const raw=record[valueIndex],num=numericValue(raw);if(num!==null){item.aggregate+=num;item.valueCount++;}}
    }
    if(!grouped.size)throw Error('নির্বাচিত Range-এ কোনো ডেটা রেকর্ড পাওয়া যায়নি');
    const detailRows=[...grouped.values()].map(item=>({...item,result:aggregation==='average'?(item.valueCount?item.aggregate/item.valueCount:0):item.aggregate})).sort((a,b)=>JSON.stringify(a.values).localeCompare(JSON.stringify(b.values),'en',{numeric:true,sensitivity:'base'}));
    const rows=[];
    if(rowFields.length>1){
     let currentParent=null,currentParentKey=null,parentItems=[];
     const addParentSubtotal=()=>{
      if(currentParentKey===null)return;
      const parentSum=parentItems.reduce((n,x)=>n+x.aggregate,0),parentCount=parentItems.reduce((n,x)=>n+x.valueCount,0);
      const subtotal=aggregation==='count'?parentItems.reduce((n,x)=>n+x.recordCount,0):aggregation==='average'?(parentCount?parentSum/parentCount:0):parentSum;
      rows.push({values:[currentParent,...Array(allFields.length-1).fill('')],result:subtotal,rowType:'subtotal'});
     };
     for(const item of detailRows){
      const parent=item.values[0],parentKey=item.parentKey;
      if(currentParentKey!==null&&parentKey!==currentParentKey)addParentSubtotal();
      if(parentKey!==currentParentKey){currentParent=parent;currentParentKey=parentKey;parentItems=[];}
      parentItems.push(item);
      const shown=[...item.values];shown[0]='';
      rows.push({...item,values:shown,rowType:'detail'});
     }
     addParentSubtotal();
    }else rows.push(...detailRows.map(item=>({...item,rowType:'detail'})));
    // Grand Total is calculated from detail groups only; subtotal display rows are never added again.
    const groupedCount=detailRows.reduce((n,x)=>n+x.recordCount,0);
    const groupedSum=detailRows.reduce((n,x)=>n+x.aggregate,0);
    const groupedValueCount=detailRows.reduce((n,x)=>n+x.valueCount,0);
    const grandTotal=aggregation==='count'?groupedCount:aggregation==='average'?(groupedValueCount?groupedSum/groupedValueCount:0):groupedSum;
    const measureHeader=aggregation==='count'?'Count':(aggregation==='sum'?'Sum of ':'Average of ')+valueField;
    setPivot({range:pivotWizard.range,rowFields,columnFields,headers:allFields,rows,grandTotal,aggregation,valueField,measureHeader});setPivotWizard(null);notify('PivotTable রিপোর্ট তৈরি হয়েছে');
   }
  }catch(error){notify(error.message)}
 };
 const printWorkbook=()=>{
  const esc=value=>String(value??'').replace(/[&<>"']/g,ch=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[ch]));
  let bounds=getSelectionBounds();
  if(pageSetup.printArea){
   const parts=pageSetup.printArea.split(':'),a=cellToPosition(parts[0]),b=cellToPosition(parts[1]);
   if(a&&b)bounds={minCol:Math.min(a.col,b.col),maxCol:Math.max(a.col,b.col),minRow:Math.min(a.row,b.row),maxRow:Math.max(a.row,b.row)};
  }else{
   const used=Object.keys(cells).map(cellToPosition).filter(Boolean);
   if(used.length)bounds={minCol:Math.min(...used.map(p=>p.col)),maxCol:Math.max(...used.map(p=>p.col)),minRow:Math.min(...used.map(p=>p.row)),maxRow:Math.max(...used.map(p=>p.row))};
  }
  const width=Math.max(1,bounds.maxCol-bounds.minCol+1),height=Math.max(1,bounds.maxRow-bounds.minRow+1);
  const cssMargin=pageSetup.margins==='narrow'?'8mm':pageSetup.margins==='wide'?'25mm':'15mm';
  const size=({A4:'A4',A3:'A3',A5:'A5',LETTER:'letter',LEGAL:'legal'})[pageSetup.paperSize]||'A4';
  const tableRows=[];
  for(let r=bounds.minRow;r<=bounds.maxRow;r++){
   let row='<tr>';
   if(pageSetup.printHeadings)row+='<th class="heading">'+r+'</th>';
   for(let c=bounds.minCol;c<=bounds.maxCol;c++){
    const addr=positionToCell(c,r),raw=cells[addr],value=typeof raw==='string'&&raw.startsWith('=')?calculateCell(addr,cells):raw;
    const st=cellStyles[addr]||{},styles=[];
    if(st.bold)styles.push('font-weight:bold');if(st.italic)styles.push('font-style:italic');if(st.underline)styles.push('text-decoration:underline');
    if(st.color)styles.push('color:'+esc(st.color));if(st.backgroundColor)styles.push('background-color:'+esc(st.backgroundColor));
    if(st.textAlign)styles.push('text-align:'+st.textAlign);if(st.verticalAlign)styles.push('vertical-align:'+st.verticalAlign);
    if(st.border)styles.push('border:'+esc(st.border));if(st.whiteSpace==='normal')styles.push('white-space:normal;overflow-wrap:anywhere');
    row+='<td style="'+styles.join(';')+'">'+esc(value===undefined?'':value)+'</td>';
   }
   row+='</tr>';tableRows.push(row);
  }
  const headings=pageSetup.printHeadings?'<tr><th class="heading"></th>'+Array.from({length:width},(_,i)=>'<th class="heading">'+esc(colName(bounds.minCol+i))+'</th>').join('')+'</tr>':'';
  const grid=pageSetup.printGridlines?'1px solid #cbd5e1':'1px solid transparent';
  const popup=window.open('','_blank');
  if(!popup){notify('Print window blocked — popup permission চালু করুন');return}
  const html='<!doctype html><html><head><meta charset="utf-8"><title>'+esc(workbookName)+' - '+esc(sheet)+'</title><style>@page{size:'+size+' '+pageSetup.orientation+';margin:'+cssMargin+'}*{box-sizing:border-box}body{font:11px Arial,sans-serif;color:#111;margin:0}.printHead{text-align:center;font-weight:bold;margin-bottom:12px}.printFoot{text-align:center;margin-top:12px}table{border-collapse:collapse;width:'+((pageSetup.scale||100)/100*100)+'%;table-layout:auto}td,th{border:'+grid+';padding:5px;min-width:55px;overflow-wrap:anywhere}th.heading{background:#f1f5f9;text-align:center;font-weight:bold}.sheetName{font-size:10px;color:#555;margin-bottom:6px}</style></head><body><div class="printHead">'+esc(headerFooter.header||'')+'</div><div class="sheetName">'+esc(workbookName)+' — '+esc(sheet)+'</div><table>'+headings+tableRows.join('')+'</table><div class="printFoot">'+esc(headerFooter.footer||'')+'</div><script>window.onload=()=>setTimeout(()=>window.print(),250)<\\/script></body></html>';
  popup.document.open();popup.document.write(html);popup.document.close();
 };
 const action=(name)=>{
  if(name==='Undo'){undo();return}
  if(name==='Redo'){redo();return}
  if(name==='Copy'){copyCell(false);return}
  if(name==='Cut'){copyCell(true);return}
  if(name==='Paste'){pasteCell();return}
  if(name==='Format Painter'){const style=cellStyles[selected]||{};setClipboard({kind:'style',value:'',style});notify('Format copied — select Paste to apply');return}
  if(name==='Save'){exportWorkbook('xlsx');return}
  if(name==='Save As'){const proposed=window.prompt('Workbook file name লিখুন:',workbookName);if(proposed===null)return;const clean=proposed.trim().replace(/[\\/:*?"<>|]/g,'_');if(!clean){notify('ফাইলের নাম দিন');return}const finalName=clean.replace(/\\.xlsx$/i,'');setWorkbookName(finalName);exportWorkbook('xlsx',finalName);return}
  if(name==='Version History'){const entries=cellHistory.map((snapshot,index)=>({index,used:Object.values(snapshot||{}).filter(v=>v!==''&&v!==null&&v!==undefined).length}));const listing=entries.map(x=>(x.index===historyIndex?'▶ ':'  ')+(x.index+1)+'. Used cells: '+x.used).join('\\n');const choice=window.prompt('Version History — বর্তমান সংস্করণে ▶ চিহ্ন।\\n'+listing+'\\n\\nRestore করতে version number লিখুন:',String(historyIndex+1));if(choice===null)return;const n=Number(choice);if(!Number.isInteger(n)||n<1||n>cellHistory.length){notify('সঠিক version number দিন');return}const idx=n-1;setCells(cellHistory[idx]||{});setCellStyles(styleHistory[idx]||{});setMergedRanges(mergedHistory[idx]||[]);setHistoryIndex(idx);setSelected('A1');setFormula((cellHistory[idx]||{})['A1']??'');notify('Version '+n+' restore হয়েছে');return}
  if(name==='Open'){document.getElementById('excelFileInput')?.click();return}
  if(name==='Gridlines'){setShowGrid(v=>!v);return}
  if(name==='Headings'){setShowHeadings(v=>!v);return}
  if(name==='Ruler'){setShowRuler(v=>!v);return}
  if(name==='Freeze Panes'){setFreezePanes(v=>!v);notify(freezePanes?'Freeze Panes বন্ধ হয়েছে':'প্রথম Row ও Column freeze হয়েছে');return}
  if(name==='Normal'){setViewMode('normal');setPageBreaks(false);return}
  if(name==='Page Layout'){setViewMode('page');return}
  if(name==='Page Break'){setPageBreaks(v=>!v);notify(pageBreaks?'Page Break preview বন্ধ':'Page Break preview চালু');return}
  if(name==='Split'){const row=Number(window.prompt('কোন Row-এর পরে split করবেন?',String(cellToPosition(selected)?.row||1)));if(!Number.isInteger(row)||row<1){notify('সঠিক Row number দিন');return}setFreezePanes(true);setSelected('A'+row);notify('Row '+row+'-এর পরের অংশে Freeze Pane প্রয়োগ হয়েছে');return}
  if(name==='New Window'){const w=window.open(window.location.href,'_blank');if(!w)notify('নতুন Window খুলতে Browser popup permission দিন');return}
  if(name==='Formula Bar'){setShowFormula(v=>!v);return}
  if(name==='New Workbook'){if(Object.values(cells).some(v=>v!==''&&v!==null&&v!==undefined)&&!window.confirm('বর্তমান Workbook বন্ধ করে নতুন Workbook তৈরি করবেন? বর্তমান কাজ auto-save থাকবে।'))return;const empty={};setCells(empty);setCellHistory([empty]);setStyleHistory([empty]);setMergedHistory([[]]);setHistoryIndex(0);setCellStyles(empty);setSheetStore({'Sheet1':{cells:empty,cellStyles:empty}});setSheet('Sheet1');setWorkbookName('Book1');setSelected('A1');setFormula('');setSelectionBand(null);setChart(null);setPivot(null);setRowHeights({});setColumnWidths({});setHiddenRows([]);setHiddenCols([]);setMergedRanges([]);setFilterConfig(null);setFilter(false);setDataValidations({});setHeaderFooter({header:'',footer:''});setPageSetup({margins:'normal',orientation:'portrait',paperSize:'A4',printArea:'',fitWidth:1,fitHeight:0,scale:100,printGridlines:false,printHeadings:false,pageOrder:'down'});setSheetProtected(false);setSheetProtectionPassword('');setWorkbookProtected(false);setWorkbookProtectionPassword('');setAllowedEditRanges([]);notify('নতুন Blank Workbook তৈরি হয়েছে');return}
  if(name==='Bold'){applyStyleToSelection({bold:!cellStyles[selected]?.bold});return}
  if(name==='Italic'){applyStyleToSelection({italic:!cellStyles[selected]?.italic});return}
  if(name==='Underline'){applyStyleToSelection({underline:!cellStyles[selected]?.underline});return}
  if(name==='Font Name'){const font=window.prompt('Font name লিখুন:',cellStyles[selected]?.fontFamily||'Arial');if(font)applyStyleToSelection({fontFamily:font});return}
  if(name==='Font Size'){const size=Number(window.prompt('Font size (6–72):',cellStyles[selected]?.fontSize||11));if(size>=6&&size<=72)applyStyleToSelection({fontSize:size});else notify('Font size 6 থেকে 72 হতে হবে');return}
  if(name==='Increase Font Size'||name==='Decrease Font Size'){const old=Number(cellStyles[selected]?.fontSize||11),size=Math.max(6,Math.min(72,old+(name==='Increase Font Size'?1:-1)));applyStyleToSelection({fontSize:size});return}
  if(name==='Font Color'||name==='Fill Color'){const key=name==='Font Color'?'color':'backgroundColor',color=window.prompt('রং লিখুন (যেমন #ff0000):',cellStyles[selected]?.[key]||(key==='color'?'#202124':'#fff2cc'));if(color)applyStyleToSelection({[key]:color});return}
  if(name==='Borders'){applyStyleToSelection({border:'1px solid #64748b'});return}
  if(name==='Align Left'||name==='Center'||name==='Align Right'){applyStyleToSelection({textAlign:name==='Align Left'?'left':name==='Center'?'center':'right'});return}
  if(name==='Top Align'||name==='Middle Align'||name==='Bottom Align'){applyStyleToSelection({verticalAlign:name==='Top Align'?'top':name==='Middle Align'?'middle':'bottom'});return}
  if(name==='Wrap Text'){applyStyleToSelection({whiteSpace:cellStyles[selected]?.whiteSpace==='normal'?'nowrap':'normal'});return}
  if(name==='Merge Cells'){mergeSelection(false);return}
  if(name==='Unmerge Cells'){mergeSelection(true);return}
  if(name==='Increase Indent'||name==='Decrease Indent'){const old=Number(cellStyles[selected]?.indent||0),indent=Math.max(0,Math.min(10,old+(name==='Increase Indent'?1:-1)));applyStyleToSelection({indent});return}
  if(['General','Number','Currency','Accounting','Percentage','Date','Time','Fraction','Scientific','Text'].includes(name)){setFormat(name);return}
  if(name==='Increase Decimal'||name==='Decrease Decimal'){const old=Number(cellStyles[selected]?.decimalPlaces??2),decimalPlaces=Math.max(0,Math.min(10,old+(name==='Increase Decimal'?1:-1)));applyStyleToSelection({decimalPlaces});return}
  if(name==='Insert Rows'||name==='Delete Rows'||name==='Insert Columns'||name==='Delete Columns'){
   const p=cellToPosition(selected);if(!p)return;const axis=name.includes('Rows')?'row':'col',remove=name.startsWith('Delete'),bounds=getSelectionBounds();const at=axis==='row'?(selectionBand?.axis==='row'?Number(selectionBand.index):bounds.minRow):(selectionBand?.axis==='col'?Number(selectionBand.index):bounds.minCol);const count=selectionBand?.axis==='range'?(axis==='row'?bounds.maxRow-bounds.minRow+1:bounds.maxCol-bounds.minCol+1):1;shiftCells(axis,at,remove,count);if(remove){setSelected(axis==='row'?positionToCell(p.col,Math.max(1,at)):positionToCell(Math.max(1,at),p.row));setSelectionBand(null)}return;
  }
  if(name==='Row Height'||name==='Column Width'){const p=cellToPosition(selected);if(!p)return;const value=Number(window.prompt(name==='Row Height'?'Row height (px):':'Column width (px):',name==='Row Height'?(rowHeights[p.row]||32):(columnWidths[p.col]||110)));if(!Number.isFinite(value)||value<10||value>500){notify('মান ১০ থেকে ৫০০-এর মধ্যে দিন');return}if(name==='Row Height')setRowHeights(v=>({...v,[p.row]:value}));else setColumnWidths(v=>({...v,[p.col]:value}));return}
  if(name==='AutoFit Row Height'){setRowHeights(v=>({...v,[cellToPosition(selected)?.row||1]:44}));notify('Row height AutoFit করা হয়েছে');return}
  if(name==='AutoFit Column Width'){setColumnWidths(v=>({...v,[cellToPosition(selected)?.col||1]:Math.max(90,String(cells[selected]??'').length*9+24)}));notify('Column width AutoFit করা হয়েছে');return}
  if(name==='Hide Rows'){const b=getSelectionBounds(),p=cellToPosition(selected);const start=selectionBand?.axis==='row'?Number(selectionBand.index):b.minRow,end=selectionBand?.axis==='row'?Number(selectionBand.index):b.maxRow;setHiddenRows(v=>[...new Set([...v,...Array.from({length:end-start+1},(_,i)=>start+i)])]);return}
  if(name==='Unhide Rows'){setHiddenRows([]);notify('সব লুকানো Row দেখানো হয়েছে');return}
  if(name==='Hide Columns'){const b=getSelectionBounds(),start=selectionBand?.axis==='col'?Number(selectionBand.index):b.minCol,end=selectionBand?.axis==='col'?Number(selectionBand.index):b.maxCol;setHiddenCols(v=>[...new Set([...v,...Array.from({length:end-start+1},(_,i)=>start+i)])]);return}
  if(name==='Unhide Columns'){setHiddenCols([]);notify('সব লুকানো Column দেখানো হয়েছে');return}
  if(name==='Insert Function'||['SUM','AVERAGE','COUNT','MAX','MIN','IF'].includes(name)){
   const fn=name==='Insert Function'?(window.prompt('Function name লিখুন (SUM, AVERAGE, COUNT, MAX, MIN, IF):','SUM')||'').trim().toUpperCase():name;
   if(!['SUM','AVERAGE','COUNT','MAX','MIN','IF'].includes(fn)){notify('এই Function এখনো সমর্থিত নয়');return}
   const p=cellToPosition(selected),defaultRange=p&&p.row>1?positionToCell(p.col,1)+':'+positionToCell(p.col,p.row-1):'A1:A1';
   const range=window.prompt(fn+'-এর Range লিখুন (যেমন A1:A10):',defaultRange);if(range===null)return;
   const clean=range.trim().replace(/\s+/g,'');if(!/^[A-Z]{1,3}[1-9][0-9]*(:[A-Z]{1,3}[1-9][0-9]*)?$/i.test(clean)){notify('সঠিক Cell বা Range লিখুন');return}
   let formula='='+fn+'('+clean.toUpperCase()+')';
   if(fn==='IF'){const condition=window.prompt('IF condition লিখুন (যেমন A1>10):','A1>10');if(condition===null)return;const yes=window.prompt('True হলে কী দেখাবে?','Yes');if(yes===null)return;const no=window.prompt('False হলে কী দেখাবে?','No');if(no===null)return;formula='=IF('+condition+',"'+yes.replace(/"/g,'""')+'","'+no.replace(/"/g,'""')+'")'}
   setFormula(formula);updateCell(selected,formula);notify(fn+' formula যুক্ত হয়েছে');return;
  }
  if(name==='Show Precedents'||name==='Show Dependents'){
   const refs=new Set(),escape=selected.replace(/[$]/g,'').toUpperCase();
   if(name==='Show Precedents'){const raw=String(cells[selected]??'');const re=/\$?[A-Z]{1,3}\$?[1-9][0-9]*/gi;for(const match of raw.matchAll(re))refs.add(match[0].replace(/\$/g,'').toUpperCase())}
   else {const re=new RegExp('\\$?'+escape.replace(/([A-Z]+)([0-9]+)/,'$1\\$?$2')+'\\b','i');for(const [addr,value] of Object.entries(cells))if(typeof value==='string'&&value.startsWith('=')&&re.test(value))refs.add(addr)}
   const list=[...refs];window.alert(name+' — '+selected+'\\n'+(list.length?list.join(', '):'কোনো সম্পর্কিত Cell পাওয়া যায়নি'));return;
  }
  if(name==='Evaluate Formula'){
   try{const value=calculateCell(selected,cells);window.alert(selected+' = '+String(value))}catch(error){window.alert('Formula Error: '+error.message)}return;
  }
  if(name==='Error Checking'){
   const errors=[];for(const [addr,value] of Object.entries(cells))if(typeof value==='string'&&value.trim().startsWith('=')){try{const result=calculateCell(addr,cells);if(typeof result==='string'&&/^#(REF|DIV\/0|VALUE|NAME|NUM|N\/A|CYCLE|ERROR)/.test(result))errors.push(addr+': '+result)}catch(error){errors.push(addr+': '+error.message)}}
   window.alert(errors.length?'Formula Errors ('+errors.length+')\\n'+errors.slice(0,50).join('\\n'):'কোনো Formula Error পাওয়া যায়নি');return;
  }
  if(name==='Calculate Now'){let count=0,errors=0;for(const [addr,value] of Object.entries(cells))if(typeof value==='string'&&value.trim().startsWith('=')){count++;try{calculateCell(addr,cells)}catch{errors++}}setCells(v=>({...v}));notify('Calculation সম্পন্ন: '+count+'টি Formula, '+errors+'টি Error');return}
  if(name==='AutoSum'){const p=cellToPosition(selected);if(p){let start=p.row-1;while(start>1&&cells[positionToCell(p.col,start-1)]!==undefined&&cells[positionToCell(p.col,start-1)]!=='')start--;const end=p.row-1;const ref=end>=start?positionToCell(p.col,start)+':'+positionToCell(p.col,end):positionToCell(p.col,1)+':'+positionToCell(p.col,1);const value='=SUM('+ref+')';setFormula(value);updateCell(selected,value);notify('SUM formula added')}return}
  if(name==='Fill Down'||name==='Fill Right'){fillSelection(name==='Fill Down'?'down':'right');return}
  if(name==='Clear Contents'){clearSelection('contents');return}
  if(name==='Clear Formats'){clearSelection('formats');return}
  if(name==='Clear All'){clearSelection('all');return}
  if(name==='Find'){const term=window.prompt('কোন লেখা/সংখ্যা খুঁজবেন?');if(term!==null&&term!==''){const found=Object.entries(cells).find(([,v])=>String(v).toLowerCase().includes(term.toLowerCase()));if(found){selectCell(found[0]);notify(found[0]+' পাওয়া গেছে')}else notify('কোনো মিল পাওয়া যায়নি')}return}
  if(name==='Replace'){replaceText();return}
  if(name==='Sort A to Z'||name==='Sort Z to A'){sortSelection(name==='Sort A to Z'?'asc':'desc');return}
  if(name==='Filter'){applyFilter();return}
  if(name==='Select All'){setSelectionBand({axis:'all'});setSelected('A1');setFormula(cells.A1??'');notify('সম্পূর্ণ Sheet সিলেক্ট হয়েছে');return}
  if(name==='Zoom In'){setSheetZoom(v=>Math.min(2,+(v+0.1).toFixed(2)));return}
  if(name==='Zoom Out'){setSheetZoom(v=>Math.max(0.5,+(v-0.1).toFixed(2)));return}
  if(name==='100%'){setSheetZoom(1);return}
  if(name==='Print'||name==='Export PDF'){printWorkbook();return}
  if(name==='Margins'){
   const choice=window.prompt('Page Margins নির্বাচন করুন: normal / narrow / wide','normal');if(choice===null)return;
   if(!['normal','narrow','wide'].includes(choice.toLowerCase())){notify('normal, narrow অথবা wide লিখুন');return}
   setPageSetup(v=>({...v,margins:choice.toLowerCase()}));notify('Page margins: '+choice);return;
  }
  if(name==='Orientation'){
   const choice=window.prompt('Orientation লিখুন: portrait / landscape',pageSetup.orientation);if(choice===null)return;
   if(!['portrait','landscape'].includes(choice.toLowerCase())){notify('portrait অথবা landscape লিখুন');return}
   setPageSetup(v=>({...v,orientation:choice.toLowerCase()}));notify('Orientation: '+choice);return;
  }
  if(name==='Paper Size'){
   const choice=window.prompt('Paper size লিখুন: A4 / A3 / A5 / Letter / Legal',pageSetup.paperSize);if(choice===null)return;
   if(!['A4','A3','A5','LETTER','LEGAL'].includes(choice.toUpperCase())){notify('A4, A3, A5, Letter অথবা Legal লিখুন');return}
   setPageSetup(v=>({...v,paperSize:choice.toUpperCase()}));notify('Paper size: '+choice.toUpperCase());return;
  }
  if(name==='Print Area'){
   const b=getSelectionBounds(),current=pageSetup.printArea||positionToCell(b.minCol,b.minRow)+':'+positionToCell(b.maxCol,b.maxRow);
   const choice=window.prompt('Print Area লিখুন (যেমন A1:F25), খালি রাখলে Clear হবে:',current);if(choice===null)return;
   const clean=choice.trim().toUpperCase().replace(/\\s/g,'');
   if(clean&&!/^[A-Z]{1,3}[1-9][0-9]*:[A-Z]{1,3}[1-9][0-9]*$/.test(clean)){notify('সঠিক Range লিখুন, যেমন A1:F25');return}
   setPageSetup(v=>({...v,printArea:clean}));notify(clean?'Print Area সেট হয়েছে: '+clean:'Print Area Clear হয়েছে');return;
  }
  if(name==='Fit to Width'||name==='Fit to Height'){
   const key=name==='Fit to Width'?'fitWidth':'fitHeight',value=Number(window.prompt('কত Page-এ Fit করবেন? (0 = automatic)',String(pageSetup[key])));if(!Number.isInteger(value)||value<0||value>20){notify('০ থেকে ২০-এর মধ্যে সংখ্যা দিন');return}
   setPageSetup(v=>({...v,[key]:value}));notify(name+' সেট হয়েছে');return;
  }
  if(name==='Scale'){
   const value=Number(window.prompt('Print Scale (%) লিখুন: 10–400',String(pageSetup.scale)));if(!Number.isInteger(value)||value<10||value>400){notify('Scale ১০ থেকে ৪০০-এর মধ্যে দিন');return}
   setPageSetup(v=>({...v,scale:value}));notify('Print scale '+value+'%');return;
  }
  if(name==='Print Gridlines'){setPageSetup(v=>({...v,printGridlines:!v.printGridlines}));notify('Print Gridlines '+(!pageSetup.printGridlines?'On':'Off'));return}
  if(name==='Print Headings'){setPageSetup(v=>({...v,printHeadings:!v.printHeadings}));notify('Print Headings '+(!pageSetup.printHeadings?'On':'Off'));return}
  if(name==='Page Order'){const value=window.prompt('Page order লিখুন: down / over',pageSetup.pageOrder);if(value===null)return;if(!['down','over'].includes(value.toLowerCase())){notify('down অথবা over লিখুন');return}setPageSetup(v=>({...v,pageOrder:value.toLowerCase()}));return}
  if(name==='Page Setup Reset'){setPageSetup({margins:'normal',orientation:'portrait',paperSize:'A4',printArea:'',fitWidth:1,fitHeight:0,scale:100,printGridlines:false,printHeadings:false,pageOrder:'down'});notify('Page Setup default-এ Reset হয়েছে');return}
  if(name==='Workbook Info'){const used=Object.values(cells).filter(v=>v!==''&&v!==null&&v!==undefined).length;window.alert('Workbook: '+workbookName+'\\nSheet: '+sheet+'\\nUsed cells: '+used+'\\nZoom: '+Math.round(sheetZoom*100)+'%');return}
  if(name==='Workbook Statistics'){const vals=Object.values(cells).filter(v=>v!==''&&v!==null&&v!==undefined);const nums=vals.map(Number).filter(Number.isFinite);window.alert('Cells: '+vals.length+'\\nNumbers: '+nums.length+'\\nText: '+(vals.length-nums.length)+'\\nSheets: '+(Object.keys(sheetStore).length||1));return}
  if(name==='Sort A to Z'||name==='Sort Z to A'){sortSelection(name==='Sort A to Z'?'asc':'desc');return}
  if(name==='Text to Columns'){
   const b=getSelectionBounds();
   const mode=window.prompt('Delimiter নির্বাচন করুন: comma (,), tab, semicolon (;), space অথবা custom লিখুন','comma');
   if(mode===null)return;
   let sep=mode.toLowerCase()==='comma'?',':mode.toLowerCase()==='tab'?'\\t':mode.toLowerCase()==='semicolon'?';':mode.toLowerCase()==='space'?' ':mode;
   if(mode.toLowerCase()==='custom'){sep=window.prompt('Custom delimiter লিখুন (এক বা একাধিক অক্ষর)','|');if(sep===null)return}
   if(!sep){notify('Delimiter সঠিক নয়');return}
   const parse=(input)=>{
    const out=[];let value='',quoted=false;
    for(let i=0;i<input.length;i++){const ch=input[i];if(ch==='"'){if(quoted&&input[i+1]==='"'){value+='"';i++}else quoted=!quoted}else if(!quoted&&input.startsWith(sep,i)){out.push(value.trim());value='';i+=sep.length-1}else value+=ch}
    out.push(value.trim());return out;
   };
   const next={...cells},styles={...cellStyles};let splitCount=0,overwrite=false;
   const rows=[];
   for(let r=b.minRow;r<=b.maxRow;r++)for(let col=b.minCol;col<=b.maxCol;col++){
    const addr=positionToCell(col,r),raw=cells[addr];if(raw===undefined||String(raw).indexOf(sep)<0)continue;
    const parts=parse(String(raw));if(parts.length<2)continue;
    rows.push({r,col,addr,parts});
    for(let j=1;j<parts.length;j++)if(cells[positionToCell(col+j,r)]!==undefined&&String(cells[positionToCell(col+j,r)])!=='')overwrite=true;
   }
   if(!rows.length){notify('নির্বাচিত Cell-এ এই Delimiter পাওয়া যায়নি');return}
   if(overwrite&&!window.confirm('ডান পাশের কিছু Cell-এ আগে থেকেই Data আছে। সেগুলো Overwrite করবেন?'))return;
   rows.forEach(({r,col,addr,parts})=>{
    const sourceStyle=cellStyles[addr];
    parts.forEach((part,j)=>{const target=positionToCell(col+j,r);next[target]=part;if(sourceStyle&&j>0)styles[target]={...sourceStyle};});
    splitCount+=parts.length-1;
   });
   recordHistory(next,styles);setCells(next);setCellStyles(styles);
   notify('Text to Columns সম্পন্ন: '+rows.length+'টি Cell থেকে '+splitCount+'টি অতিরিক্ত Column তৈরি হয়েছে');return;
  }
  if(name==='Remove Duplicates'){
   const b=getSelectionBounds(),seen=new Set(),remove=[];for(let r=b.minRow;r<=b.maxRow;r++){const key=Array.from({length:b.maxCol-b.minCol+1},(_,i)=>String(cells[positionToCell(b.minCol+i,r)]??'').trim().toLowerCase()).join('¦');if(key.replace(/¦/g,'')==='')continue;if(seen.has(key))remove.push(r);else seen.add(key)}
   if(!remove.length){notify('কোনো Duplicate Row পাওয়া যায়নি');return}if(!window.confirm(remove.length+'টি duplicate row মুছে ফেলবেন?'))return;
   const removed=new Set(remove),next={};for(const [addr,val] of Object.entries(cells)){const p=cellToPosition(addr);if(!p)continue;const shift=remove.filter(r=>r<p.row).length;if(!removed.has(p.row))next[positionToCell(p.col,p.row-shift)]=val}
   const styles={};for(const [addr,val] of Object.entries(cellStyles)){const p=cellToPosition(addr);if(!p)continue;const shift=remove.filter(r=>r<p.row).length;if(!removed.has(p.row))styles[positionToCell(p.col,p.row-shift)]=val}
   recordHistory(next,styles);setCells(next);setCellStyles(styles);notify(remove.length+'টি duplicate row বাদ');return;
  }
  if(name==='Table'){
   const b=getSelectionBounds();if(b.maxRow<=b.minRow){notify('Table বানাতে Header ও অন্তত একটি Data Row নির্বাচন করুন');return}
   const next={...cellStyles};for(let r=b.minRow;r<=b.maxRow;r++)for(let c=b.minCol;c<=b.maxCol;c++){const id=positionToCell(c,r);next[id]={...(next[id]||{}),border:'1px solid #94a3b8',backgroundColor:r===b.minRow?'#dbeafe':(r%2===0?'#f8fafc':'#ffffff'),bold:r===b.minRow};}
   setCellStyles(next);recordHistory(cells,next);setFilterConfig({range:positionToCell(b.minCol,b.minRow)+':'+positionToCell(b.maxCol,b.maxRow),headerRow:b.minRow});setFilter(true);notify('নির্বাচিত Range-এ Table format ও Filter প্রয়োগ হয়েছে');return;
  }
  if(name==='Icons'){
   const choices=['★','☆','✓','✗','●','◆','▲','▼','➜','☎','✉','⌂','⚠','☑','♥','৳'];
   const symbol=window.prompt('একটি Icon/Symbol বেছে নিন:\n'+choices.join('  '),'★');if(symbol===null)return;
   if(!symbol.trim()){notify('Icon নির্বাচন করুন');return}updateCell(selected,symbol.trim());applyStyleToSelection({fontSize:22,textAlign:'center',color:'#2563eb'});return;
  }
  if(name==='Header & Footer'){
   const header=window.prompt('Print Header লিখুন (খালি রাখলে থাকবে না):',headerFooter.header||'');if(header===null)return;
   const footer=window.prompt('Print Footer লিখুন (খালি রাখলে থাকবে না):',headerFooter.footer||'');if(footer===null)return;
   setHeaderFooter({header,footer});notify('Header ও Footer সংরক্ষণ হয়েছে');return;
  }
  if(name==='Hyperlink'){
   const url=window.prompt('ওয়েবসাইটের সম্পূর্ণ URL লিখুন (https:// দিয়ে):','https://');if(url===null)return;
   if(!/^https?:\/\//i.test(url)){notify('সঠিক URL দিন (https:// বা http://)');return}
   const label=window.prompt('সেলে যে লেখা দেখাবে:',url);if(label===null)return;
   updateCell(selected,label);applyStyleToSelection({hyperlink:url,color:'#0563c1',underline:true});notify('Hyperlink যুক্ত হয়েছে');return;
  }
  if(name==='Picture'){
   const picker=document.createElement('input');picker.type='file';picker.accept='image/*';
   picker.onchange=()=>{const file=picker.files?.[0];if(!file)return;if(file.size>4*1024*1024){notify('ছবির সাইজ ৪ MB-এর কম হতে হবে');return}
    const reader=new FileReader();reader.onload=()=>{applyStyleToSelection({imageData:String(reader.result),imageName:file.name});notify('নির্বাচিত সেলে ছবি যুক্ত হয়েছে')};reader.onerror=()=>notify('ছবি পড়া যায়নি');reader.readAsDataURL(file);
   };picker.click();return;
  }
  if(name==='Shapes'){
   const shape=window.prompt('Shape লিখুন: rectangle / oval / triangle','rectangle');if(shape===null)return;
   if(!['rectangle','oval','triangle'].includes(shape.toLowerCase())){notify('rectangle, oval অথবা triangle লিখুন');return}
   applyStyleToSelection({shape:shape.toLowerCase(),backgroundColor:'#dbeafe',border:'2px solid #2563eb'});notify('নির্বাচিত সেলে '+shape+' shape যুক্ত হয়েছে');return;
  }
  if(name==='Text Box'){
   const value=window.prompt('Text Box-এর লেখা লিখুন:');if(value===null)return;
   updateCell(selected,value);applyStyleToSelection({textBox:true,whiteSpace:'normal',border:'1px solid #64748b',backgroundColor:'#fff7ed'});notify('নির্বাচিত সেলে Text Box তৈরি হয়েছে');return;
  }
  if(name==='Get Data'){document.getElementById('excelFileInput')?.click();return}
  if(name==='From Text/CSV'){
   const picker=document.createElement('input');picker.type='file';picker.accept='.csv,text/csv';
   picker.onchange=()=>{const file=picker.files?.[0];if(!file)return;if(!file.name.toLowerCase().endsWith('.csv')){notify('শুধু CSV ফাইল নির্বাচন করুন');return}importWorkbook(file)};
   picker.click();return;
  }
  if(name==='Refresh All'){
   try{const next={...cells};let count=0;for(const [ref,value] of Object.entries(cells)){if(typeof value==='string'&&value.trim().startsWith('=')){calculateCell(ref,cells);count++}}
    setCells(next);recordHistory(next);setSheetStore(v=>({...v,[sheet]:{cells:next,cellStyles}}));notify(count+'টি Formula পুনরায় হিসাব করা হয়েছে');return;
   }catch(error){notify('Refresh ব্যর্থ: '+error.message);return}
  }
  if(name==='What-If Analysis'){
   const formulaCell=(window.prompt('যে Formula সেলের ফলাফল পরীক্ষা করবেন (যেমন C2):','C2')||'').trim().toUpperCase();if(!formulaCell)return;
   if(!cells[formulaCell]||typeof cells[formulaCell]!=='string'||!cells[formulaCell].startsWith('=')){notify('Formula থাকা সেল লিখুন');return}
   const inputCell=(window.prompt('যে Input সেলের মান পরিবর্তন করে পরীক্ষা করবেন (যেমন A2):','A2')||'').trim().toUpperCase();if(!cellToPosition(inputCell)){notify('সঠিক Input Cell লিখুন');return}
   const raw=window.prompt('পরীক্ষার মানগুলো কমা দিয়ে লিখুন (যেমন 10,20,30):','10,20,30');if(raw===null)return;
   const values=raw.split(',').map(v=>Number(v.trim()));if(!values.length||values.some(v=>!Number.isFinite(v))){notify('সব Scenario মান সংখ্যা হতে হবে');return}
   const out=cellToPosition(selected),next={...cells};next[positionToCell(out.col,out.row)]='Input: '+inputCell;next[positionToCell(out.col+1,out.row)]='Result: '+formulaCell;
   values.forEach((value,i)=>{const row=out.row+i+1;next[positionToCell(out.col,row)]=value;try{const result=calculateCell(formulaCell,{...cells,[inputCell]:value});next[positionToCell(out.col+1,row)]=typeof result==='number'?result:String(result)}catch{next[positionToCell(out.col+1,row)]='#ERROR!'}});
   setCells(next);recordHistory(next);notify('What-If Scenario Table তৈরি হয়েছে');return;
  }
  if(name==='Goal Seek'){
   const formulaCell=(window.prompt('যে Formula সেলের ফলাফল পরিবর্তন করবেন (যেমন C2):','C2')||'').trim().toUpperCase();if(!formulaCell)return;
   if(typeof cells[formulaCell]!=='string'||!cells[formulaCell].trim().startsWith('=')){notify('Formula থাকা সেল লিখুন');return}
   const inputCell=(window.prompt('কোন Input Cell পরিবর্তন করবেন (যেমন A2):','A2')||'').trim().toUpperCase();if(!cellToPosition(inputCell)){notify('সঠিক Input Cell লিখুন');return}
   const target=Number(window.prompt('Formula-র কাঙ্ক্ষিত ফলাফল কত?','100'));if(!Number.isFinite(target)){notify('সঠিক Target Number দিন');return}
   let x0=Number(cells[inputCell]??0);if(!Number.isFinite(x0))x0=0;let x1=x0===0?1:x0*1.05+0.01,solved=false;
   const evaluate=x=>{const y=calculateCell(formulaCell,{...cells,[inputCell]:x});const n=Number(y);if(!Number.isFinite(n))throw Error('Formula ফলাফল সংখ্যা নয়');return n-target};
   try{let y0=evaluate(x0),y1=evaluate(x1);for(let i=0;i<100;i++){if(Math.abs(y1)<1e-8){solved=true;break}const den=y1-y0;if(Math.abs(den)<1e-14)break;let x2=x1-y1*(x1-x0)/den;if(!Number.isFinite(x2))break;x2=Math.max(-1e12,Math.min(1e12,x2));x0=x1;y0=y1;x1=x2;y1=evaluate(x1)}if(!solved&&Math.abs(y1)<1e-6)solved=true;
    if(!solved){notify('Goal Seek সমাধান পায়নি। অন্য Initial Value বা Target চেষ্টা করুন');return}
    updateCell(inputCell,x1);notify('Goal Seek সফল: '+inputCell+' = '+Number(x1.toFixed(8)));return;
   }catch(error){notify('Goal Seek ব্যর্থ: '+error.message);return}
  }
  if(name==='Remove Duplicates'){
   const b=getSelectionBounds();
   if(b.maxRow<=b.minRow){notify('Duplicate খুঁজতে একাধিক Row নির্বাচন করুন');return}
   const headers=Array.from({length:b.maxCol-b.minCol+1},(_,i)=>positionToCell(b.minCol+i,b.minRow));
   const headerChoice=window.confirm('প্রথম Row কি Header? OK = Header, Cancel = সব Row-কে Data ধরুন');
   const start=headerChoice?b.minRow+1:b.minRow;
   if(start>b.maxRow){notify('Data Row পাওয়া যায়নি');return}
   const seen=new Set(),keep=[];
   for(let r=start;r<=b.maxRow;r++){
    const values=headers.map((_,i)=>cells[positionToCell(b.minCol+i,r)]??'');
    if(values.every(v=>String(v).trim()==='')){keep.push(r);continue}
    const key=JSON.stringify(values.map(v=>String(v).trim().toLowerCase()));
    if(!seen.has(key)){seen.add(key);keep.push(r)}
   }
   const removed=(b.maxRow-start+1)-(keep.length-(headerChoice?0:0));
   if(removed<=0){notify('কোনো Duplicate Row পাওয়া যায়নি');return}
   const next={...cells},nextStyles={...cellStyles};
   const allRows=new Map();
   for(let r=b.minRow;r<=b.maxRow;r++){
    const rowCells={},rowStyles={};
    Object.keys(cells).forEach(id=>{const p=cellToPosition(id);if(p?.row===r)rowCells[p.col]=cells[id]});
    Object.keys(cellStyles).forEach(id=>{const p=cellToPosition(id);if(p?.row===r)rowStyles[p.col]=cellStyles[id]});
    allRows.set(r,{rowCells,rowStyles});
   }
   for(let r=b.minRow;r<=b.maxRow;r++){
    Object.keys(next).forEach(id=>{if(cellToPosition(id)?.row===r)delete next[id]});
    Object.keys(nextStyles).forEach(id=>{if(cellToPosition(id)?.row===r)delete nextStyles[id]});
   }
   const outputRows=headerChoice?[b.minRow,...keep]:keep;
   outputRows.forEach((source,index)=>{
    const target=b.minRow+index,record=allRows.get(source);
    Object.entries(record?.rowCells||{}).forEach(([col,value])=>next[positionToCell(Number(col),target)]=value);
    Object.entries(record?.rowStyles||{}).forEach(([col,value])=>nextStyles[positionToCell(Number(col),target)]=value);
   });
   setCells(next);setCellStyles(nextStyles);recordHistory(next,nextStyles);
   notify(removed+'টি Duplicate Row সরানো হয়েছে');return;
  }
  if(name==='Goal Seek'){
   const targetRef=String(selected||'').toUpperCase(),formula=cells[targetRef];
   if(typeof formula!=='string'||!formula.trim().startsWith('=')){notify('প্রথমে একটি Formula Cell নির্বাচন করুন');return}
   const targetInput=window.prompt('Formula cell '+targetRef+'-এর কাঙ্ক্ষিত ফলাফল লিখুন','100');if(targetInput===null)return;
   const target=Number(targetInput);if(!Number.isFinite(target)){notify('সঠিক Target Number দিন');return}
   const variableInput=window.prompt('যে Input Cell পরিবর্তন করে ফলাফল মিলাবেন তার ঠিকানা লিখুন','A1');if(variableInput===null)return;
   const variableRef=variableInput.trim().toUpperCase().replace(/\$/g,'');
   if(!/^[A-Z]{1,3}[1-9][0-9]*$/.test(variableRef)||variableRef===targetRef){notify('সঠিক আলাদা Input Cell দিন');return}
   const original=cells[variableRef]??0,initial=Number(original);
   if(!Number.isFinite(initial)){notify('Input Cell-এ Numeric Value থাকতে হবে');return}
   const evaluate=x=>{const draft={...cells,[variableRef]:x};const value=calculateCell(targetRef,draft);const n=Number(value);if(!Number.isFinite(n))throw Error('Formula result numeric নয়');return n};
   let x=initial,fx;
   try{fx=evaluate(x)-target}catch(error){notify('Formula হিসাব করা যায়নি: '+error.message);return}
   if(Math.abs(fx)<1e-8){notify('বর্তমান Input-এই Target মিলে গেছে');return}
   let x2=x+(Math.abs(x)*0.01||0.01),f2;
   try{f2=evaluate(x2)-target}catch(error){notify('Formula হিসাব করা যায়নি: '+error.message);return}
   let solved=false;
   for(let i=0;i<100;i++){
    if(Math.abs(f2)<1e-7){solved=true;break}
    const denominator=f2-fx;
    if(Math.abs(denominator)<1e-14)break;
    const nextX=x2-f2*(x2-x)/denominator;
    if(!Number.isFinite(nextX)||Math.abs(nextX)>1e12)break;
    x=x2;fx=f2;x2=nextX;
    try{f2=evaluate(x2)-target}catch(error){break}
   }
   if(!solved&&Math.abs(f2)<1e-5)solved=true;
   if(!solved){notify('Goal Seek সমাধান পায়নি। অন্য Initial Value দিয়ে চেষ্টা করুন।');return}
   const next={...cells,[variableRef]:Number(x2.toPrecision(12))};
   setCells(next);recordHistory(next,cellStyles);setSelected(variableRef);
   notify('Goal Seek সফল: '+variableRef+' = '+next[variableRef]+'; '+targetRef+' ≈ '+target);return;
  }
  if(name==='Data Validation'){
   const mode=window.prompt('Validation type লিখুন: list / whole / decimal / clear','list');if(mode===null)return;
   const b=getSelectionBounds(),next={...dataValidations};
   if(mode.toLowerCase()==='clear'){for(let r=b.minRow;r<=b.maxRow;r++)for(let c=b.minCol;c<=b.maxCol;c++)delete next[positionToCell(c,r)];setDataValidations(next);notify('নির্বাচিত সেলের Validation সরানো হয়েছে');return}
   let rule;
   if(mode.toLowerCase()==='list'){const input=window.prompt('অনুমোদিত মানগুলো কমা দিয়ে লিখুন (যেমন: Yes,No,Pending)');if(input===null)return;const values=input.split(',').map(x=>x.trim()).filter(Boolean);if(!values.length){notify('তালিকা খালি রাখা যাবে না');return}rule={type:'list',values}}
   else if(mode.toLowerCase()==='whole'||mode.toLowerCase()==='decimal'||mode.toLowerCase()==='textlength'){const min=Number(window.prompt(mode.toLowerCase()==='textlength'?'সর্বনিম্ন অক্ষর সংখ্যা':'সর্বনিম্ন মান','0')),max=Number(window.prompt(mode.toLowerCase()==='textlength'?'সর্বোচ্চ অক্ষর সংখ্যা':'সর্বোচ্চ মান','100'));if(!Number.isInteger(min)||!Number.isInteger(max)||min<0||max<min){notify('সঠিক সর্বনিম্ন ও সর্বোচ্চ মান দিন');return}rule={type:mode.toLowerCase()==='textlength'?'textLength':mode.toLowerCase(),min,max}}
   else if(mode.toLowerCase()==='date'){const min=window.prompt('শুরুর তারিখ (YYYY-MM-DD)','2026-01-01'),max=window.prompt('শেষ তারিখ (YYYY-MM-DD)','2026-12-31');if(min===null||max===null||!/^\\d{4}-\\d{2}-\\d{2}$/.test(min)||!/^\\d{4}-\\d{2}-\\d{2}$/.test(max)||min>max){notify('সঠিক তারিখের সীমা দিন');return}rule={type:'date',min,max}}
   else{notify('list, whole, decimal, textLength, date অথবা clear লিখুন');return}
   for(let r=b.minRow;r<=b.maxRow;r++)for(let c=b.minCol;c<=b.maxCol;c++)next[positionToCell(c,r)]=rule;
   setDataValidations(next);notify('নির্বাচিত সেলে Data Validation চালু হয়েছে');return;
  }
  if(name==='Clear Filter'){setFilter(false);setFilterConfig(null);notify('Filter সরানো হয়েছে');return}
  if(name==='Normal'||name==='Page Layout'||name==='Page Break'){notify(name+' view নির্বাচিত');return}
  if(name==='Spelling'){if(!spellingEnabled){const lang=window.prompt('Spelling language লিখুন: en-US অথবা bn-BD',spellingLanguage);if(lang===null)return;const normalized=lang.trim().toLowerCase()==='bn-bd'?'bn-BD':'en-US';setSpellingLanguage(normalized);setSpellingEnabled(true);notify('Browser spelling check চালু ('+normalized+') — dictionary support ব্রাউজারের ওপর নির্ভরশীল')}else{setSpellingEnabled(false);notify('Spelling check বন্ধ হয়েছে')}return}
  if(name==='Check Workbook'){const used=Object.entries(cells).filter(([,v])=>v!==''&&v!==null&&v!==undefined);const formulas=used.filter(([,v])=>typeof v==='string'&&v.trim().startsWith('='));const errors=used.map(([addr,v])=>{let value=v;if(typeof v==='string'&&v.trim().startsWith('=')){try{value=calculateCell(addr,cells)}catch(error){value='#ERROR! '+(error?.message||'Formula evaluation failed')}}return {addr,value}}).filter(x=>typeof x.value==='string'&&['#REF','#DIV/0','#VALUE','#NAME','#NUM','#N/A','#CYCLE','#ERROR'].some(code=>x.value.startsWith(code)));const blanks=Object.keys(cellStyles).filter(addr=>!cells[addr]&&(cellStyles[addr]?.comment||cellStyles[addr]?.note));const commentCount=Object.values(cellStyles).filter(style=>style?.comment).length,noteCount=Object.values(cellStyles).filter(style=>style?.note).length;const lines=['Workbook Check Summary','Used cells: '+used.length,'Formula cells: '+formulas.length,'Formula errors: '+errors.length,'Comments: '+commentCount,'Notes: '+noteCount,'Comments/Notes on blank cells: '+blanks.length];if(errors.length)lines.push('\\nFormula errors:',...errors.slice(0,30).map(x=>x.addr+': '+x.value));if(errors.length>30)lines.push('...and '+(errors.length-30)+' more');window.alert(lines.join('\\n'));return}
  if(name==='New Comment'||name==='Notes'){const key=name==='Notes'?'note':'comment',label=name==='Notes'?'Note':'Comment',old=cellStyles[selected]?.[key]||'';const value=window.prompt(selected+' সেলের '+label+' লিখুন:',old);if(value!==null){const next={...cellStyles},entry={...(next[selected]||{})};if(value.trim())entry[key]=value.trim();else delete entry[key];if(Object.keys(entry).length)next[selected]=entry;else delete next[selected];setCellStyles(next);recordHistory(cells,next);notify(value.trim()?label+' সংরক্ষণ হয়েছে':label+' মুছে ফেলা হয়েছে')}return}
  if(name==='Show Comments'){const entries=Object.entries(cellStyles).flatMap(([addr,style])=>['comment','note'].filter(k=>style?.[k]).map(k=>addr+' ['+(k==='note'?'Note':'Comment')+']: '+style[k]));window.alert(entries.length?entries.join('\\n'):'কোনো Comment বা Note নেই');return}
  if(name==='Delete All Comments'){const next={...cellStyles};let removed=0;for(const [addr,style] of Object.entries(next)){if(!style?.comment&&!style?.note)continue;const updated={...style};if(updated.comment){delete updated.comment;removed++}if(updated.note){delete updated.note;removed++}if(Object.keys(updated).length)next[addr]=updated;else delete next[addr]}if(!removed){notify('মুছে ফেলার মতো Comment বা Note নেই');return}if(!window.confirm(removed+'টি Comment/Note মুছে ফেলবেন?'))return;setCellStyles(next);recordHistory(cells,next);notify(removed+'টি Comment/Note মুছে ফেলা হয়েছে');return}
  if(name==='Protect Sheet'){if(sheetProtected){const pass=window.prompt('Sheet protection password দিন:');if(pass===null)return;if(pass!==sheetProtectionPassword){notify('ভুল password');return}setSheetProtected(false);setSheetProtectionPassword('');notify('Sheet protection বন্ধ হয়েছে')}else{const pass=window.prompt('Sheet protect করতে password দিন:');if(pass===null||!pass)return;const confirm=window.prompt('Password আবার লিখুন:');if(confirm!==pass){notify('Password মিলেনি');return}setSheetProtectionPassword(pass);setSheetProtected(true);notify('Sheet protected — শুধু Allow Edit Ranges সম্পাদনাযোগ্য')}return}
  if(name==='Protect Workbook'){if(workbookProtected){const pass=window.prompt('Workbook protection password দিন:');if(pass===null)return;if(pass!==workbookProtectionPassword){notify('ভুল password');return}setWorkbookProtected(false);setWorkbookProtectionPassword('');notify('Workbook protection বন্ধ হয়েছে')}else{const pass=window.prompt('Workbook protect করতে password দিন:');if(pass===null||!pass)return;const confirm=window.prompt('Password আবার লিখুন:');if(confirm!==pass){notify('Password মিলেনি');return}setWorkbookProtectionPassword(pass);setWorkbookProtected(true);notify('Workbook structure protected হয়েছে')}return}
  if(name==='PivotTable'){createPivot();return}
  if(name==='Recommended Charts'){insertChart('Column');return}
  if(name==='Column Chart'){insertChart('Column');return}
  if(name==='Line Chart'){insertChart('Line');return}
  if(name==='Pie Chart'){insertChart('Pie');return}
  if(name==='Allow Edit Ranges'){const current=allowedEditRanges.join(', ');const value=window.prompt('Editable range লিখুন (যেমন A1:B5)। একাধিক range কমা দিয়ে লিখুন। ফাঁকা দিলে সব range মুছবে।',current);if(value===null)return;const ranges=value.split(',').map(x=>x.trim().toUpperCase()).filter(Boolean);if(ranges.some(x=>!/^[A-Z]{1,3}[1-9][0-9]*:[A-Z]{1,3}[1-9][0-9]*$/.test(x))){notify('Range format সঠিক নয় — যেমন A1:B5');return}setAllowedEditRanges(ranges);notify(ranges.length?'Editable range সংরক্ষণ হয়েছে':'Editable range মুছে ফেলা হয়েছে');return}
  if(name==='Freeze Panes'){const p=cellToPosition(selected);if(p){const rows=p.row-1,cols=p.col-1;document.documentElement.style.setProperty('--freeze-top',rows+'px');document.documentElement.style.setProperty('--freeze-left',cols+'px');notify('Freeze boundary: '+selected)}return}
  notify(name+' selected');
 };
 const usedMaxRow=Object.keys(cells).reduce((m,key)=>Math.max(m,cellToPosition(key)?.row||0),0);
 const usedMaxCol=Object.keys(cells).reduce((m,key)=>Math.max(m,cellToPosition(key)?.col||0),0);
 const rows=Array.from({length:visibleRows},(_,i)=>i+1).filter(r=>!hiddenRows.includes(r)&&(!filterConfig||r<filterConfig.range.start||r>filterConfig.range.end||filterConfig.rows.includes(r))),cols=Array.from({length:visibleCols},(_,i)=>i+1).filter(c=>!hiddenCols.includes(c));
 return <div className="excelApp">
  <header className="excelTop"><button className="backBtn" onClick={requestExit} aria-label="Back">‹</button><div className="excelTitle"><b><Table2 size={19}/> Smart Excel</b><small>Workbook • {sheet}</small></div><button className="topIcon" onClick={()=>action('Save')}><Save size={19}/></button><button className="topIcon" onClick={()=>notify('More workbook options')}><Settings size={19}/></button></header>
  <div className="workbookBar"><span className="fileBadge"><FileSpreadsheet size={16}/></span><input aria-label="Workbook name" value={workbookName} onChange={e=>setWorkbookName(e.target.value)} /><span className="savedState">● Saved locally</span></div>
  <div className="ribbonStrip">{ribbonData[active].map(group=><div className="ribbonGroup" key={group.group}><div className="ribbonButtons">{group.items.map(([label,Icon])=><button key={label} className="ribbonAction" onClick={()=>action(label)} title={label}><Icon size={17}/><span>{label}</span></button>)}</div><small>{group.group}</small></div>)}</div>
  {showFormula&&<div className="formulaBar"><span className="cellRef">{selected}</span><span className="fx">fx</span><input value={formula} onChange={e=>{setFormula(e.target.value);updateCell(selected,e.target.value)}} placeholder="Enter value or formula"/></div>}
  {headerFooter.header&&<div className="printHeader">{headerFooter.header}</div>}
  {showRuler&&<div className="sheetRuler" style={{display:'flex',height:22,overflow:'hidden',background:'#f1f5f9',borderBottom:'1px solid #cbd5e1',fontSize:11}}><span style={{width:42,flex:'0 0 42px'}}>↔</span>{cols.map(c=><span key={c} style={{width:columnWidths[c]||110,flex:'0 0 '+(columnWidths[c]||110)+'px',textAlign:'center',borderLeft:'1px solid #cbd5e1'}}>{colName(c)}</span>)}</div>}<div className="sheetScroll" style={{touchAction:'none'}} onPaste={e=>{const raw=e.clipboardData?.getData('text/plain');if(raw!==undefined&&raw!==''){e.preventDefault();pasteMatrix(raw.replace(/\r/g,'').replace(/\n$/,'').split('\n').map(line=>line.split('\t')))}}} onPointerDown={e=>{const box=e.currentTarget;touchPoints.current.set(e.pointerId,{x:e.clientX,y:e.clientY});try{box.setPointerCapture(e.pointerId)}catch{};if(touchPoints.current.size>=2){setIsSelectingRange(false);const pts=[...touchPoints.current.values()].slice(0,2),dx=pts[0].x-pts[1].x,dy=pts[0].y-pts[1].y;touchGesture.current={mode:'pinch',distance:Math.max(1,Math.hypot(dx,dy)),zoom:sheetZoom,midX:(pts[0].x+pts[1].x)/2,midY:(pts[0].y+pts[1].y)/2,scrollLeft:box.scrollLeft,scrollTop:box.scrollTop};return;}const td=e.target.closest('td[data-cell]');if(td){e.preventDefault();const id=td.dataset.cell;if(pivotWizard?.step==='range'){setIsSelectingRange(true);setRangeStart(id);setRangeEnd(id);touchGesture.current={mode:'select'};}else{selectCell(id);setIsSelectingRange(true);setRangeStart(id);setRangeEnd(id);setSelectionBand({axis:'range',start:id,end:id});touchGesture.current={mode:'select'};}}else{touchGesture.current={mode:'pan',x:e.clientX,y:e.clientY,scrollLeft:box.scrollLeft,scrollTop:box.scrollTop};}}} onPointerMove={e=>{if(!touchPoints.current.has(e.pointerId))return;touchPoints.current.set(e.pointerId,{x:e.clientX,y:e.clientY});const box=e.currentTarget;if(touchPoints.current.size>=2){const pts=[...touchPoints.current.values()].slice(0,2),g=touchGesture.current;if(!g||g.mode!=='pinch')return;const dx=pts[0].x-pts[1].x,dy=pts[0].y-pts[1].y,dist=Math.max(1,Math.hypot(dx,dy)),midX=(pts[0].x+pts[1].x)/2,midY=(pts[0].y+pts[1].y)/2;setSheetZoom(Math.max(.6,Math.min(2.5,g.zoom*dist/g.distance)));box.scrollLeft=g.scrollLeft-(midX-g.midX);box.scrollTop=g.scrollTop-(midY-g.midY);return;}const g=touchGesture.current;if(g?.mode==='select'){const rect=box.getBoundingClientRect(),edge=42,step=26;if(e.clientX>rect.right-edge)box.scrollLeft+=step;else if(e.clientX<rect.left+edge)box.scrollLeft-=step;if(e.clientY>rect.bottom-edge)box.scrollTop+=step;else if(e.clientY<rect.top+edge)box.scrollTop-=step;const px=Math.max(rect.left+2,Math.min(rect.right-2,e.clientX)),py=Math.max(rect.top+2,Math.min(rect.bottom-2,e.clientY));const el=document.elementFromPoint(px,py),td=el?.closest?.('td[data-cell]');if(td){setRangeEnd(td.dataset.cell);if(pivotWizard?.step!=='range')setSelectionBand({axis:'range',start:rangeStart||selected,end:td.dataset.cell});}}else if(g?.mode==='pan'){box.scrollLeft=g.scrollLeft-(e.clientX-g.x);box.scrollTop=g.scrollTop-(e.clientY-g.y);}}} onPointerUp={e=>{touchPoints.current.delete(e.pointerId);if(touchPoints.current.size===0){touchGesture.current=null;setIsSelectingRange(false);}else if(touchPoints.current.size===1){const p=[...touchPoints.current.values()][0];touchGesture.current={mode:'pan',x:p.x,y:p.y,scrollLeft:e.currentTarget.scrollLeft,scrollTop:e.currentTarget.scrollTop};}}} onPointerCancel={e=>{touchPoints.current.delete(e.pointerId);touchGesture.current=null;setIsSelectingRange(false)}} onScroll={e=>{const el=e.currentTarget;if(el.scrollTop+el.clientHeight>=el.scrollHeight-180)setVisibleRows(v=>v+100);if(el.scrollLeft+el.clientWidth>=el.scrollWidth-180)setVisibleCols(v=>v+10)}}><table style={{zoom:sheetZoom,outline:viewMode==='page'?'1px solid #94a3b8':undefined}} className={(showGrid?'sheetGrid':'sheetGrid noGrid')+(pageBreaks?' pageBreakPreview':'')}><thead style={freezePanes?{position:"sticky",top:0,zIndex:5}:undefined}><tr>{showHeadings&&<th className="corner" style={freezePanes?{position:'sticky',left:0,zIndex:7}:undefined} title="Select All — সম্পূর্ণ Sheet নির্বাচন" aria-label="Select All" onClick={()=>{setSelectionBand({axis:'all'});setSelected('A1');setFormula(cells.A1||'');notify('সম্পূর্ণ Sheet সিলেক্ট হয়েছে')}}>☷</th>}{showHeadings&&cols.map(c=><th key={c} style={{minWidth:columnWidths[c]||110,width:columnWidths[c]||110,position:freezePanes?'sticky':undefined,top:freezePanes?0:undefined,zIndex:freezePanes?6:undefined}} onClick={()=>selectWholeColumn(c)} className={selectionBand?.axis==='col'&&selectionBand.index===c?'axisSelected':''}>{colName(c)}</th>)}</tr></thead><tbody>{rows.map(r=><tr key={r} style={{height:rowHeights[r]||32,borderBottom:pageBreaks&&r%20===0?'2px dashed #2563eb':undefined}}>{showHeadings&&<th style={freezePanes?{position:'sticky',left:0,zIndex:4}:undefined} onClick={()=>selectWholeRow(r)} className={selectionBand?.axis==='row'&&selectionBand.index===r?'axisSelected':''}>{r}</th>}{cols.map(c=>{const id=colName(c)+r;const style=cellStyles[id]||{};const merged=mergedRanges.find(m=>r>=m.startRow&&r<=m.endRow&&c>=m.startCol&&c<=m.endCol);if(merged&&(r!==merged.startRow||c!==merged.startCol))return null;let display=cells[id]===undefined?'':String(calculateCell(id,cells));if(display!==''&&Number.isFinite(Number(display))){const num=Number(display),dp=Number.isInteger(style.decimalPlaces)?style.decimalPlaces:2;if(style.numberFormat==='currency'||style.numberFormat==='accounting')display='৳ '+num.toLocaleString('en-US',{minimumFractionDigits:dp,maximumFractionDigits:dp});else if(style.numberFormat==='percent')display=(num*100).toFixed(dp)+'%';else if(style.numberFormat==='number')display=num.toLocaleString('en-US',{minimumFractionDigits:dp,maximumFractionDigits:dp});else if(style.numberFormat==='scientific')display=num.toExponential(dp);else if(style.numberFormat==='fraction')display=num.toString();}return <td key={id} title={[style.comment&&('Comment: '+style.comment),style.note&&('Note: '+style.note)].filter(Boolean).join(' | ')} colSpan={merged?merged.endCol-merged.startCol+1:1} rowSpan={merged?merged.endRow-merged.startRow+1:1} data-cell={id} style={{border:style.border||undefined,boxShadow:(style.comment||style.note)?'inset -3px 3px 0 #e0a020':undefined,minWidth:columnWidths[c]||110,width:columnWidths[c]||110,borderRadius:style.shape==='oval'?'50%':undefined,clipPath:style.shape==='triangle'?'polygon(50% 0%, 0% 100%, 100% 100%)':undefined,backgroundColor:style.shape?style.backgroundColor:undefined}} onClick={()=>{if(pivotWizard?.step==='range')return;const now=Date.now(),last=cellTap.current;if(last.id===id&&now-last.time<450){selectCell(id);setEditingCell(id);cellTap.current={id:null,time:0};setTimeout(()=>{const el=document.querySelector('input[aria-label="'+id+'"]');el?.focus();el?.select();},0);}else{selectCell(id);setEditingCell(null);cellTap.current={id,time:now};}}} className={(selected===id?'selectedCell ':'')+((selectionBand?.axis==='col'&&selectionBand.index===c)||(selectionBand?.axis==='row'&&selectionBand.index===r)||selectionBand?.axis==='all'?'bandSelected ':'')+(((pivotWizard?.step==='range'&&rangeStart&&rangeEnd)||selectionBand?.axis==='range')&&(()=>{const a=cellToPosition(selectionBand?.axis==='range'?selectionBand.start:rangeStart),b=cellToPosition(selectionBand?.axis==='range'?selectionBand.end:rangeEnd),p=cellToPosition(id);return p&&a&&b&&p.row>=Math.min(a.row,b.row)&&p.row<=Math.max(a.row,b.row)&&p.col>=Math.min(a.col,b.col)&&p.col<=Math.max(a.col,b.col)})()?'rangeSelected':'')}>{style.imageData?<img src={style.imageData} alt={style.imageName||"Inserted picture"} style={{display:"block",maxWidth:180,maxHeight:110,objectFit:"contain",margin:"auto"}}/>:style.hyperlink?<a href={style.hyperlink} target="_blank" rel="noopener noreferrer" onClick={e=>e.stopPropagation()} style={{color:style.color||"#0563c1",textDecoration:"underline",display:"block",padding:6}}>{display||style.hyperlink}</a>:<input style={{fontWeight:style.bold?'700':'400',fontStyle:style.italic?'italic':'normal',textDecoration:style.underline?'underline':'none',textAlign:style.textAlign||'left',verticalAlign:style.verticalAlign||'middle',fontFamily:style.fontFamily||'inherit',fontSize:style.fontSize?style.fontSize+'px':undefined,paddingLeft:style.indent?(6+style.indent*8)+'px':undefined,whiteSpace:style.whiteSpace||'nowrap',color:style.color||'inherit',backgroundColor:style.backgroundColor||'transparent',border:style.border||undefined}} readOnly={editingCell!==id} aria-label={id} value={editingCell===id?(cells[id]??''):display} spellCheck={spellingEnabled} lang={spellingLanguage} onFocus={()=>{if(editingCell===id)selectCell(id)}} onBlur={()=>setEditingCell(null)} onChange={e=>{updateCell(id,e.target.value);if(selected===id)setFormula(e.target.value)}}/>}</td>})}</tr>)}</tbody></table></div>{headerFooter.footer&&<div className="printFooter">{headerFooter.footer}</div>}

  {pivotWizard&&<section className="pivotWizard"><div className="excelPanelHead"><b>PivotTable তৈরি করুন</b><button onClick={()=>setPivotWizard(null)}>×</button></div>{pivotWizard.step==='range'?<><p>Sheet-এর সব লেখা থাকা সেল স্বয়ংক্রিয়ভাবে সিলেক্ট করা হয়েছে। প্রয়োজনে রেঞ্জটি দেখে Next চাপুন।</p><b>নির্বাচিত রেঞ্জ: {rangeStart&&rangeEnd?(()=>{const a=cellToPosition(rangeStart),b=cellToPosition(rangeEnd);return positionToCell(Math.min(a.col,b.col),Math.min(a.row,b.row))+':'+positionToCell(Math.max(a.col,b.col),Math.max(a.row,b.row))})():'এখনো সিলেক্ট হয়নি'}</b><button className="pivotNext" onClick={pivotNext}>Next</button></>:<><p>নিচের Field-গুলো Rows ও Columns-এর ১০টি করে ঘরে টেনে রাখুন। যতগুলো ঘরে Field দেবেন, রিপোর্টে ততগুলো স্তর ব্যবহার হবে।</p><div className="pivotFields">{pivotWizard.headers.map(h=><div draggable onDragStart={e=>e.dataTransfer.setData('text/plain',h)} onClick={()=>notify('Field-টি Rows বা Columns-এর ঘরে টেনে দিন')} className="pivotField" key={h}>{h}</div>)}</div><div className="pivotMeasureBox"><b>Values / হিসাবের Field</b><select value={pivotWizard.valueField} onChange={e=>setPivotWizard(p=>({...p,valueField:e.target.value}))}><option value="">Field নির্বাচন করুন</option>{pivotWizard.headers.map(h=><option key={h} value={h}>{h}</option>)}</select><b>হিসাবের ধরন</b><select value={pivotWizard.aggregation} onChange={e=>setPivotWizard(p=>({...p,aggregation:e.target.value}))}><option value="count">Count — কয়টি রেকর্ড</option><option value="sum">Sum — মোট যোগফল</option><option value="average">Average — গড়</option></select><small>Count হলে Values Field লাগবে না। Sum/Average হলে সংখ্যার কলাম নির্বাচন করুন।</small></div><div className="pivotDropZones pivotTenZones">{['rowFields','columnFields'].map((zone,zi)=><div className="pivotZone" key={zone}><b>{zi===0?'Rows':'Columns'}</b>{pivotWizard[zone].map((field,index)=><div className="pivotSlot" key={index} onDragOver={e=>e.preventDefault()} onDrop={e=>{e.preventDefault();const h=e.dataTransfer.getData('text/plain');setPivotWizard(p=>({...p,[zone]:p[zone].map((v,i)=>i===index?h:v)}))}}><small>{index+1}</small><span>{field||'Field এখানে দিন'}</span><button onClick={()=>setPivotWizard(p=>({...p,[zone]:p[zone].map((v,i)=>i===index?'':v)}))}>×</button></div>)}</div>)}</div><button className="pivotNext" onClick={pivotNext}>Next — PivotTable রিপোর্ট তৈরি করুন</button></>}</section>}
   {chart&&<section className="excelDataPanel"><div className="excelPanelHead"><b>{chart.title}</b><button onClick={()=>setChart(null)}>×</button></div><div className="excelChart"><svg viewBox="0 0 340 190" role="img" aria-label={chart.type+' chart'}><line x1="28" y1="155" x2="330" y2="155" stroke="currentColor" opacity=".45"/>{chart.type==='Pie'?(()=>{const values=(chart.series[0]?.values||[]).map(v=>Math.max(0,Number(v)||0));const total=values.reduce((a,b)=>a+b,0)||1;const colors=['#2563eb','#16a34a','#dc7a20','#9333ea','#0891b2','#ca8a04'];let offset=0;return values.map((v,i)=>{const length=v/total*251.2;const circle=<circle key={i} cx="125" cy="92" r="40" fill="none" stroke={colors[i%colors.length]} strokeWidth="80" strokeDasharray={length+' '+(251.2-length)} strokeDashoffset={-offset} transform="rotate(-90 125 92)"/>;offset+=length;return circle})})():chart.type==='Line'?chart.series.map((series,si)=>{const max=Math.max(1,...series.values.map(Math.abs));const points=series.values.map((v,i)=>(35+i*(285/Math.max(1,series.values.length-1))+','+(150-v/max*115)).join(' ')).join(' ');return <g key={si}><polyline points={points} fill="none" stroke={['#2563eb','#16a34a','#dc7a20'][si%3]} strokeWidth="3"/>{series.values.map((v,i)=><circle key={i} cx={35+i*(285/Math.max(1,series.values.length-1))} cy={150-v/max*115} r="3.5" fill={['#2563eb','#16a34a','#dc7a20'][si%3]}/>)}</g>}):chart.series.flatMap((series,si)=>series.values.map((v,i)=>{const max=Math.max(1,...chart.series.flatMap(x=>x.values).map(Math.abs));const group=285/Math.max(1,chart.labels.length);const width=Math.max(5,Math.min(28,group/(chart.series.length+1)));const height=Math.abs(v)/max*112;return <rect key={si+'-'+i} x={35+i*group+si*width} y={150-height} width={width-3} height={height} rx="2" fill={['#2563eb','#16a34a','#dc7a20'][si%3]}/>;}))}{chart.labels.map((label,i)=><text key={i} x={35+i*(285/Math.max(1,chart.labels.length))} y="174" fontSize="9" fill="currentColor">{String(label).slice(0,8)}</text>)}</svg></div><small>Source: {chart.range}</small></section>}
  {pivot&&<section className="excelDataPanel"><div className="excelPanelHead"><b>PivotTable রিপোর্ট</b><div><button onClick={()=>exportPivotSheet()}>Export → নতুন Sheet</button><button onClick={()=>setPivot(null)}>×</button></div></div><div className="pivotTable"><table><thead><tr>{pivot.headers.map((h,i)=><th key={i}>{h}</th>)}<th>{pivot.measureHeader||'Count'}</th></tr></thead><tbody>{pivot.rows.map((row,i)=><tr key={i} style={row.rowType==='subtotal'?{fontWeight:700,background:'#e8f0fe'}:{}}>{row.values.map((v,j)=><td key={j}>{v}</td>)}<th>{Number(row.result??row.aggregate??0).toLocaleString(undefined,{maximumFractionDigits:2})}</th></tr>)}</tbody><tfoot><tr>{pivot.headers.map((h,i)=><th key={i}>{i===0?'Grand Total':''}</th>)}<th>{Number(pivot.grandTotal).toLocaleString(undefined,{maximumFractionDigits:2})}</th></tr></tfoot></table></div><small>Rows: {pivot.rowFields.join(', ')||'None'} • Columns: {pivot.columnFields.join(', ')||'None'} • Source: {pivot.range} • প্রতিটি Field আলাদা কলামে দেখানো হয়েছে।</small></section>}
  <div className="sheetTabs"><button className="sheetNav" onClick={()=>notify('Previous sheet')}>‹</button><button className="sheetTab activeSheet" onClick={renameSheet} onContextMenu={e=>{e.preventDefault();deleteSheet()}} title="Tap to rename; long/right click to delete">{sheet}</button>{Object.keys({...sheetStore,[sheet]:{cells,cellStyles}}).filter(n=>n!==sheet).map(n=><button key={n} className="sheetTab" onClick={()=>switchSheet(n)}>{n}</button>)}<button className="addSheet" onClick={addSheet}><Plus size={17}/></button><button className="sheetNav" onClick={deleteSheet} title="Delete current sheet"><X size={15}/></button><span className="sheetStatus"><button onClick={()=>setSheetZoom(v=>Math.max(.6,Number((v-.1).toFixed(2))))}>−</button> {Math.round(sheetZoom*100)}% <button onClick={()=>setSheetZoom(v=>Math.min(2.5,Number((v+.1).toFixed(2))))}>+</button> Ready</span></div>
  <button className="ribbonMenuButton" onClick={()=>setMenuOpen(v=>!v)} aria-label="Open ribbon menu"><Grid2X2 size={24}/></button>
  {menuOpen&&<><div className="menuScrim" onClick={()=>setMenuOpen(false)}></div><div className="ribbonMenuPanel"><div className="menuPanelTitle">Workbook Menu <button onClick={()=>setMenuOpen(false)}><X size={18}/></button></div>{menus.map(([name,Icon])=><button key={name} className={'menuChoice '+(active===name?'chosen':'')} onClick={()=>{setActive(name);setMenuOpen(false)}}><Icon size={19}/><span>{name}</span>{active===name&&<Check size={16}/>}</button>)}</div></>}
  {isImporting&&<div className="excelToast">বড় Excel ফাইল প্রসেস হচ্ছে… অনুগ্রহ করে অপেক্ষা করুন</div>}
  <input id="excelFileInput" type="file" hidden accept=".xls,.xlsx,.csv" onChange={e=>{if(e.target.files?.[0])importWorkbook(e.target.files[0]);e.target.value=''}}/>
  {notice&&<div className="excelToast">{notice}</div>}\n  {confirmExit&&<div role="dialog" aria-modal="true" aria-labelledby="exitTitle" style={{position:'fixed',inset:0,zIndex:10000,background:'rgba(15,23,42,.62)',display:'flex',alignItems:'center',justifyContent:'center',padding:22}}>
   <div style={{width:'100%',maxWidth:360,background:'#fff',color:'#172033',borderRadius:18,padding:22,boxShadow:'0 18px 60px rgba(0,0,0,.3)'}}>
    <h3 id="exitTitle" style={{margin:'0 0 10px',fontSize:19}}>Excel বন্ধ করবেন?</h3>
    <p style={{margin:'0 0 18px',lineHeight:1.55,fontSize:14}}>Excel থেকে বের হওয়ার আগে ফাইলটি Save করতে চান? Save করলে একটি Excel ফাইল ডাউনলোড হবে।</p>
    <button onClick={()=>{exportWorkbook('xlsx');setConfirmExit(false);setTimeout(onBack,500)}} style={{width:'100%',padding:13,border:0,borderRadius:10,background:'#16834b',color:'#fff',fontWeight:700,marginBottom:9}}>Save & Exit</button>
    <button onClick={()=>{setConfirmExit(false);onBack()}} style={{width:'100%',padding:12,border:'1px solid #d1d5db',borderRadius:10,background:'#fff',color:'#b42318',fontWeight:600,marginBottom:9}}>Exit Without Saving</button>
    <button onClick={()=>setConfirmExit(false)} style={{width:'100%',padding:12,border:0,borderRadius:10,background:'#eef2f7',color:'#172033',fontWeight:600}}>Cancel — Excel-এ থাকুন</button>
   </div>
  </div>}
 </div>
}

function WordEditor({file,onBack}){
 const key='smartoffice-word-'+file.id;
 const [title,setTitle]=React.useState(file.name.replace(/\.(docx?|txt)$/i,''));
 const [body,setBody]=React.useState(()=>{try{return localStorage.getItem(key)||''}catch{return ''}});
 const [saved,setSaved]=React.useState(true);
 React.useEffect(()=>{const t=setTimeout(()=>{try{localStorage.setItem(key,body);localStorage.setItem(key+'-title',title);setSaved(true)}catch{}},400);return()=>clearTimeout(t)},[body,title]);
 const save=()=>{localStorage.setItem(key,body);localStorage.setItem(key+'-title',title);setSaved(true)};
 const exportFile=()=>{save();const u=URL.createObjectURL(new Blob([body],{type:'text/plain;charset=utf-8'}));const a=document.createElement('a');a.href=u;a.download=(title||'Document')+'.txt';a.click();URL.revokeObjectURL(u)};
 return <div style={{minHeight:'100vh',background:'#eef2f7',color:'#172033'}}><header style={{background:'#fff',padding:12,display:'flex',gap:10,alignItems:'center'}}><button onClick={()=>{save();onBack()}}>‹ Back</button><b style={{flex:1}}>Word Editor</b><button onClick={exportFile}>Export</button><button onClick={save}>Save</button></header><main style={{maxWidth:850,margin:'18px auto',padding:16}}><input value={title} onChange={e=>{setTitle(e.target.value);setSaved(false)}} style={{width:'100%',padding:12,fontSize:20,fontWeight:700,marginBottom:12}}/><textarea value={body} onChange={e=>{setBody(e.target.value);setSaved(false)}} placeholder="এখানে ডকুমেন্ট লিখুন..." style={{boxSizing:'border-box',width:'100%',minHeight:'70vh',padding:28,border:0,background:'#fff',fontSize:16,lineHeight:1.8,resize:'vertical'}}/><small>{saved?'Saved on this device':'Saving…'}</small></main></div>
}
function PowerPointEditor({file,onBack}){
 const key='smartoffice-ppt-'+file.id;
 const [title,setTitle]=React.useState(file.name.replace(/\.(pptx?)$/i,''));
 const [slides,setSlides]=React.useState(()=>{try{return JSON.parse(localStorage.getItem(key)||'null')||[{heading:'প্রেজেন্টেশনের শিরোনাম',body:'এখানে বক্তব্য লিখুন'}]}catch{return[{heading:'প্রেজেন্টেশনের শিরোনাম',body:''}]}});
 const [active,setActive]=React.useState(0);
 const [saved,setSaved]=React.useState(true);
 React.useEffect(()=>{const t=setTimeout(()=>{try{localStorage.setItem(key,JSON.stringify(slides));localStorage.setItem(key+'-title',title);setSaved(true)}catch{}},400);return()=>clearTimeout(t)},[slides,title]);
 const save=()=>{localStorage.setItem(key,JSON.stringify(slides));localStorage.setItem(key+'-title',title);setSaved(true)};
 const update=(field,value)=>setSlides(old=>old.map((x,i)=>i===active?{...x,[field]:value}:x));
 const exportFile=()=>{save();const data=slides.map((x,i)=>'SLIDE '+(i+1)+'\n'+x.heading+'\n'+x.body).join('\n\n');const u=URL.createObjectURL(new Blob([data],{type:'text/plain;charset=utf-8'}));const a=document.createElement('a');a.href=u;a.download=(title||'Presentation')+'.txt';a.click();URL.revokeObjectURL(u)};
 return <div style={{minHeight:'100vh',background:'#eef2f7',color:'#172033'}}><header style={{background:'#fff',padding:12,display:'flex',gap:10,alignItems:'center'}}><button onClick={()=>{save();onBack()}}>‹ Back</button><b style={{flex:1}}>PowerPoint Editor</b><button onClick={exportFile}>Export</button><button onClick={save}>Save</button></header><main style={{maxWidth:900,margin:'16px auto',padding:16}}><input value={title} onChange={e=>{setTitle(e.target.value);setSaved(false)}} style={{width:'100%',padding:10,fontSize:19,fontWeight:700,marginBottom:12}}/><div style={{display:'flex',gap:8,overflowX:'auto',paddingBottom:12}}>{slides.map((x,i)=><button key={i} onClick={()=>setActive(i)} style={{padding:12,borderRadius:8,border:i===active?'2px solid #d66a28':'1px solid #cbd5e1',background:'#fff'}}>Slide {i+1}</button>)}<button onClick={()=>{setSlides(v=>[...v,{heading:'নতুন স্লাইড',body:''}]);setActive(slides.length);setSaved(false)}}>＋ Add slide</button><button disabled={slides.length<2} onClick={()=>{setSlides(v=>v.filter((_,i)=>i!==active));setActive(Math.max(0,active-1));setSaved(false)}}>Delete slide</button></div><section style={{background:'#fff',aspectRatio:'16/9',maxWidth:760,margin:'12px auto',padding:'7% 8%',boxShadow:'0 3px 16px #0f172a18',display:'flex',flexDirection:'column',justifyContent:'center',gap:18}}><textarea value={slides[active]?.heading||''} onChange={e=>{update('heading',e.target.value);setSaved(false)}} placeholder="Slide title" style={{width:'100%',fontSize:28,fontWeight:700,textAlign:'center',border:0,resize:'none',outline:0}}/><textarea value={slides[active]?.body||''} onChange={e=>{update('body',e.target.value);setSaved(false)}} placeholder="Slide content" style={{width:'100%',minHeight:140,fontSize:19,textAlign:'center',border:0,resize:'vertical',outline:0,lineHeight:1.6}}/></section><small>{saved?'Saved on this device':'Saving…'}</small></main></div>
}


function CalculatorEditor({onBack}){
 const [expr,setExpr]=React.useState('');
 const [result,setResult]=React.useState('0');
 const [memory,setMemory]=React.useState(0);
 const [history,setHistory]=React.useState(()=>{try{return JSON.parse(localStorage.getItem('smartoffice-calculator-history')||'[]')}catch{return[]}});
 const [showHistory,setShowHistory]=React.useState(false);
 const [justEvaluated,setJustEvaluated]=React.useState(false);
 const calculate=input=>{
  const s=String(input).replace(/,/g,'').replace(/×/g,'*').replace(/÷/g,'/').replace(/−/g,'-').replace(/%/g,'/100');
  if(!s.trim())return 0;
  if(!/^[0-9+*/().\s-]+$/.test(s))throw Error('Invalid expression');
  const tokens=s.match(/(?:\d+(?:\.\d*)?|\.\d+)|[()+*/-]/g)||[];let i=0;
  const primary=()=>{const t=tokens[i++];if(t==='('){const v=add();if(tokens[i++]!==')')throw Error('Bracket');return v}if(t==='-')return -primary();if(t==='+')return primary();if(t===undefined||(!/^\d/.test(t)&&!/^\.\d/.test(t)))throw Error('Number');return Number(t)};
  const mul=()=>{let v=primary();while(tokens[i]==='*'||tokens[i]==='/'){const op=tokens[i++],r=primary();if(op==='/'&&r===0)throw Error('Division by zero');v=op==='*'?v*r:v/r}return v};
  const add=()=>{let v=mul();while(tokens[i]==='+'||tokens[i]==='-'){const op=tokens[i++],r=mul();v=op==='+'?v+r:v-r}return v};
  const v=add();if(i!==tokens.length||!Number.isFinite(v))throw Error('Invalid');return v;
 };
 const fmt=n=>Number(n).toLocaleString('en-US',{maximumFractionDigits:10});
 const equals=()=>{try{const v=calculate(expr);const shown=fmt(v);setResult(shown);setHistory(old=>{const next=[{expression:expr,result:shown},...old].slice(0,100);localStorage.setItem('smartoffice-calculator-history',JSON.stringify(next));return next});setExpr(String(v));setJustEvaluated(true)}catch(e){setResult(e.message==='Division by zero'?'Cannot divide by zero':'Error');setJustEvaluated(false)}};
 const press=k=>{if(k==='AC'){setExpr('');setResult('0');setJustEvaluated(false);return}if(k==='DEL'){setExpr(v=>v.slice(0,-1));setJustEvaluated(false);return}if(k==='='){equals();return}if(k==='M+'){try{setMemory(v=>v+calculate(expr||result))}catch{}return}if(k==='M-'){try{setMemory(v=>v-calculate(expr||result))}catch{}return}if(k==='MS'){try{setMemory(calculate(expr||result))}catch{}return}if(k==='MR'){setExpr(v=>v+String(memory));setJustEvaluated(false);return}if(k==='MC'){setMemory(0);return}if(k==='+/-'){setExpr(v=>v.startsWith('-')?v.slice(1):'-('+v+')');setJustEvaluated(false);return}if(k==='%'){setExpr(v=>v+'%');setJustEvaluated(false);return}setExpr(v=>(justEvaluated&&!['+','−','×','÷'].includes(k)?'':v)+k);setJustEvaluated(false)};
 React.useEffect(()=>{const onKey=e=>{if(/^[0-9.]$/.test(e.key)||e.key==='('||e.key===')')press(e.key);else if(e.key==='+')press('+');else if(e.key==='-')press('−');else if(e.key==='*')press('×');else if(e.key==='/'){e.preventDefault();press('÷')}else if(e.key==='Enter'||e.key==='='){e.preventDefault();press('=')}else if(e.key==='Backspace')press('DEL');else if(e.key==='Escape')press('AC')};window.addEventListener('keydown',onKey);return()=>window.removeEventListener('keydown',onKey)});
 const functions=[['MC','#f97316'],['MR','#e879f9'],['M+','#06b6d4'],['M−','#8b5cf6'],['MS','#22c55e'],['AC','#ef4444'],['DEL','#f59e0b'],['%','#0ea5e9'],['( )','#14b8a6']];
 const digits=['1','2','3','4','5','6','7','8','9','0','.','+/-'];
 const operators=['+','−','×','÷'];
 return <div className="calcFull">
  <header className="calcHeader"><button className="calcBack" onClick={onBack}>‹</button><div className="calcBrand"><span className="calcBrandDot">S</span><span><b>Smart Calculator</b><small>3D • Full Screen Calculator</small></span></div><button className="calcHistoryToggle" onClick={()=>setShowHistory(v=>!v)}>History</button></header>
  <main className="calcMain">
   <section className="calcDisplay"><div className="calcDisplayTop"><span>CALCULATION</span><span>MEMORY: {fmt(memory)}</span></div><div className="calcExpression">{expr||'0'}</div><div className="calcResult">{result}</div></section>
   <section className="calcFunctionGrid">{functions.map(([k,color])=><button key={k} onClick={()=>press(k==='M−'?'M-':k==='( )'?'(':k)} style={{'--key-color':color}} className="calcFunctionKey">{k}</button>)}</section>
   <section className="calcDigitGrid">{digits.map(k=><button key={k} onClick={()=>press(k)} className={'calcDigitKey '+(k==='0'?'calcZero':'')+(k==='+/-'?' calcUtility':'')}>{k}</button>)}</section>
   <section className="calcOperatorGrid">{operators.map(k=><button key={k} onClick={()=>press(k)} className="calcOperatorKey">{k}</button>)}</section>
   <button className="calcEqualsKey" onClick={()=>press('=')}>= <span>CALCULATE</span></button>
   {showHistory&&<section className="calcHistoryPanel"><div className="calcHistoryHead"><b>Calculation History</b><button onClick={()=>{setHistory([]);localStorage.removeItem('smartoffice-calculator-history')}}>Clear all</button></div>{history.length?history.map((h,i)=><button key={i} className="calcHistoryItem" onClick={()=>{setExpr(h.expression);setResult(h.result);setShowHistory(false)}}><span>{h.expression}</span><b>{h.result}</b></button>):<p>No calculations yet.</p>}</section>}
   <div className="calcFoot">Smart Office • Fast, colorful & accurate</div>
  </main>
 </div>
}

function BikeRouteMap({ride}){
 const host=React.useRef(null),mapRef=React.useRef(null),lineRef=React.useRef(null),markersRef=React.useRef([]);
 React.useEffect(()=>{if(!host.current||mapRef.current)return;const map=L.map(host.current,{zoomControl:true,attributionControl:true}).setView([23.7,90.4],6);L.tileLayer('https://{s}.tile.openstreetmap.org/{z}/{x}/{y}.png',{maxZoom:19,attribution:'© OpenStreetMap'}).addTo(map);mapRef.current=map;const timer=setTimeout(()=>map.invalidateSize(),250);return()=>{clearTimeout(timer);map.remove();mapRef.current=null}},[]);
 React.useEffect(()=>{const map=mapRef.current;if(!map)return;const pts=(ride&&ride.points||[]).filter(p=>Number.isFinite(p.lat)&&Number.isFinite(p.lng));if(lineRef.current){lineRef.current.remove();lineRef.current=null}markersRef.current.forEach(m=>m.remove());markersRef.current=[];if(!pts.length)return;const latlngs=pts.map(p=>[p.lat,p.lng]);lineRef.current=L.polyline(latlngs,{color:'#1675e8',weight:5,opacity:.9}).addTo(map);markersRef.current.push(L.circleMarker(latlngs[0],{radius:7,color:'#fff',weight:2,fillColor:'#16a34a',fillOpacity:1}).addTo(map));if(pts.length>1)markersRef.current.push(L.circleMarker(latlngs[latlngs.length-1],{radius:7,color:'#fff',weight:2,fillColor:'#e43d3d',fillOpacity:1}).addTo(map));map.fitBounds(lineRef.current.getBounds(),{padding:[24,24],maxZoom:16})},[ride&&ride.points&&ride.points.length]);
 return <div className="bikeRealMap" ref={host}><div className="bikeMapHint">{ride&&ride.points&&ride.points.length?'GPS route':'Waiting for GPS points'}</div></div>;
}
function BikeTracker({onBack}){
 const [rides,setRides]=React.useState(()=>{try{return JSON.parse(localStorage.getItem('smartoffice-bike-rides')||'[]')}catch{return[]}});
 const [activeRide,setActiveRide]=React.useState(()=>{try{return JSON.parse(localStorage.getItem('smartoffice-bike-active')||'null')}catch{return null}});
 const [gpsStatus,setGpsStatus]=React.useState('GPS permission required');
 const [gpsError,setGpsError]=React.useState('');
 const [speed,setSpeed]=React.useState(0);
 const [distance,setDistance]=React.useState(0);
 const [elapsed,setElapsed]=React.useState(0);
 const [autoMode,setAutoMode]=React.useState(true);
 const [tab,setTab]=React.useState('ride');
 const [expandedRide,setExpandedRide]=React.useState(null);
 const [historyFilter,setHistoryFilter]=React.useState('all');
 const [historyFrom,setHistoryFrom]=React.useState('');
 const [historyTo,setHistoryTo]=React.useState('');
 const [reportFrom,setReportFrom]=React.useState('');
 const [reportTo,setReportTo]=React.useState('');
 const [fuelPrice,setFuelPrice]=React.useState(()=>Number(localStorage.getItem('smartoffice-bike-fuel-price')||120));
 const [motorcycles,setMotorcycles]=React.useState(()=>{try{const x=JSON.parse(localStorage.getItem('smartoffice-bike-list')||'[]');return x.length?x:[{id:'bike-default',name:'My Motorcycle'}]}catch{return[{id:'bike-default',name:'My Motorcycle'}]}});
 const [selectedBikeId,setSelectedBikeId]=React.useState(()=>localStorage.getItem('smartoffice-bike-selected')||'bike-default');
 const [fuelLogs,setFuelLogs]=React.useState(()=>{try{return JSON.parse(localStorage.getItem('smartoffice-bike-fuel-logs')||'[]')}catch{return[]}});
 const [maintenance,setMaintenance]=React.useState(()=>{try{return JSON.parse(localStorage.getItem('smartoffice-bike-maintenance')||'[]')}catch{return[]}});
 const [speedLimit,setSpeedLimit]=React.useState(()=>Number(localStorage.getItem('smartoffice-bike-speed-limit')||80));
 const [fuelLitres,setFuelLitres]=React.useState('');
 const [fuelOdometer,setFuelOdometer]=React.useState('');
 const [fuelFullTank,setFuelFullTank]=React.useState(true);
 const [maintenanceName,setMaintenanceName]=React.useState('');
 const [maintenanceDate,setMaintenanceDate]=React.useState('');
 const speedAlertRef=React.useRef(false);
 React.useEffect(()=>{localStorage.setItem('smartoffice-bike-list',JSON.stringify(motorcycles))},[motorcycles]);
 React.useEffect(()=>{localStorage.setItem('smartoffice-bike-selected',selectedBikeId)},[selectedBikeId]);
 React.useEffect(()=>{localStorage.setItem('smartoffice-bike-fuel-logs',JSON.stringify(fuelLogs))},[fuelLogs]);
 React.useEffect(()=>{localStorage.setItem('smartoffice-bike-maintenance',JSON.stringify(maintenance))},[maintenance]);
 React.useEffect(()=>{localStorage.setItem('smartoffice-bike-speed-limit',String(speedLimit))},[speedLimit]);

 const watchRef=React.useRef(null),activeRef=React.useRef(activeRide),stillRef=React.useRef(0),movingRef=React.useRef(0),lastFixRef=React.useRef(null);const finishingRef=React.useRef(false);
 React.useEffect(()=>{activeRef.current=activeRide;localStorage.setItem('smartoffice-bike-active',JSON.stringify(activeRide))},[activeRide]);
 React.useEffect(()=>{localStorage.setItem('smartoffice-bike-rides',JSON.stringify(rides))},[rides]);
 React.useEffect(()=>{localStorage.setItem('smartoffice-bike-fuel-price',String(fuelPrice))},[fuelPrice]);
 const haversine=(a,b)=>{const rad=n=>n*Math.PI/180,R=6371000,dLat=rad(b.lat-a.lat),dLon=rad(b.lng-a.lng),x=Math.sin(dLat/2)**2+Math.cos(rad(a.lat))*Math.cos(rad(b.lat))*Math.sin(dLon/2)**2;return 2*R*Math.asin(Math.sqrt(x))};
 const beginRide=(mode='manual',point=null)=>{
  if(activeRef.current)return;
  const now=Date.now(),p=point?{lat:point.coords.latitude,lng:point.coords.longitude,at:now,speed:Math.max(0,point.coords.speed||0)}:null;
  const ride={id:'ride-'+now,bikeId:selectedBikeId,startAt:new Date(now).toISOString(),endAt:null,distanceM:0,seconds:0,maxSpeed:0,points:p?[p]:[],mode};
  activeRef.current=ride;setActiveRide(ride);setDistance(0);setElapsed(0);setGpsError('');
  if(Capacitor.isNativePlatform())BikeTracking.startTracking().catch(()=>setGpsError('Native background tracking could not start. Check location permission.'));
 };
 const finishRide=async()=>{
  const ride=activeRef.current;if(!ride||finishingRef.current)return;finishingRef.current=true;
  let native={};if(Capacitor.isNativePlatform()){try{await BikeTracking.stopTracking();await new Promise(resolve=>setTimeout(resolve,300));native=await BikeTracking.getTrackData();}catch(e){}}
  let nativePoints=[];try{nativePoints=JSON.parse(native.pointsJson||'[]')}catch(e){}
  const finished={...ride,points:nativePoints.length?nativePoints:ride.points,distanceM:native.distanceM>0?native.distanceM:ride.distanceM,maxSpeed:Math.max(ride.maxSpeed||0,native.maxSpeed||0),endAt:new Date().toISOString(),seconds:Math.max(ride.seconds,Math.floor((Date.now()-Date.parse(ride.startAt))/1000))};
  if(finished.points.length>1||finished.distanceM>20){setRides(old=>[finished,...old].slice(0,2000))}
  activeRef.current=null;setActiveRide(null);setSpeed(0);setDistance(0);setElapsed(0);stillRef.current=0;movingRef.current=0;finishingRef.current=false;
 };
 React.useEffect(()=>{
  if(!navigator.geolocation){setGpsStatus('GPS not supported');return}
  setGpsStatus('Connecting GPS…');
  watchRef.current=navigator.geolocation.watchPosition(pos=>{
   const accuracy=Number(pos.coords.accuracy||999);
   if(accuracy>50){setGpsStatus('GPS signal weak');return}
   setGpsStatus('GPS connected');setGpsError('');
   const stamp=pos.timestamp||Date.now();
   const previousFix=lastFixRef.current;
   let kmh=pos.coords.speed!=null?Math.max(0,pos.coords.speed*3.6):0;
   if((pos.coords.speed==null||!Number.isFinite(kmh))&&previousFix){
    const seconds=(stamp-previousFix.at)/1000;
    if(seconds>0&&seconds<30){const metres=haversine(previousFix,{lat:pos.coords.latitude,lng:pos.coords.longitude});kmh=metres/seconds*3.6}
   }
   setSpeed(kmh);
   if(kmh>speedLimit&&!speedAlertRef.current){speedAlertRef.current=true;setGpsError('Speed alert: '+Math.round(kmh)+' km/h — limit '+speedLimit+' km/h');if(navigator.vibrate)navigator.vibrate([250,120,250]);}else if(kmh<speedLimit-5)speedAlertRef.current=false;
   const point={lat:pos.coords.latitude,lng:pos.coords.longitude,at:stamp,speed:kmh,accuracy};
   lastFixRef.current=point;
   let current=activeRef.current;
   if(!current&&autoMode){
    if(kmh>=5){movingRef.current+=1}else movingRef.current=0;
    if(movingRef.current>=2){beginRide('automatic',pos);movingRef.current=0}
   }else if(current){
    const points=current.points||[],last=points[points.length-1];
    if(last){const gap=haversine(last,point),elapsedSeconds=(point.at-(last.at||point.at))/1000,minimumMove=Math.max(3,((last.accuracy||0)+accuracy)*0.35);if(gap>=minimumMove&&gap<200&&elapsedSeconds>0){current={...current,distanceM:(current.distanceM||0)+gap,points:[...points,point],maxSpeed:Math.max(current.maxSpeed||0,kmh)}}}
    else current={...current,points:[point],maxSpeed:kmh};
    if(kmh<2){stillRef.current+=1}else stillRef.current=0;
    activeRef.current=current;setActiveRide(current);setDistance(current.distanceM||0);
    if(stillRef.current>=24&&current.mode==='automatic'){finishRide()}
   }
  },err=>{setGpsStatus('GPS unavailable');setGpsError(err.code===1?'Location permission denied. Allow location access in Android settings.':err.message||'GPS signal unavailable')},{enableHighAccuracy:true,maximumAge:1000,timeout:15000});
  return()=>{if(watchRef.current!==null)navigator.geolocation.clearWatch(watchRef.current)};
 },[autoMode]);
 React.useEffect(()=>{if(!activeRide)return;const tick=()=>{setElapsed(Math.floor((Date.now()-Date.parse(activeRide.startAt))/1000));setDistance(activeRef.current?.distanceM||0)};tick();const id=setInterval(tick,1000);return()=>clearInterval(id)},[!!activeRide]);
 React.useEffect(()=>{if(!Capacitor.isNativePlatform())return;let alive=true;const syncNative=async()=>{if(!activeRef.current)return;try{const data=await BikeTracking.getTrackData();if(!alive||!data.tracking)return;const points=JSON.parse(data.pointsJson||'[]');if(points.length){const current=activeRef.current;if(current){const merged={...current,points,distanceM:Math.max(current.distanceM||0,data.distanceM||0),maxSpeed:Math.max(current.maxSpeed||0,data.maxSpeed||0)};activeRef.current=merged;setActiveRide(merged);setDistance(merged.distanceM)}}}catch(e){}};syncNative();const id=setInterval(syncNative,5000);return()=>{alive=false;clearInterval(id)}},[]);

 const fmtTime=s=>{const h=Math.floor(s/3600),m=Math.floor(s%3600/60),sec=s%60;return [h,m,sec].map(n=>String(n).padStart(2,'0')).join(':')};
 const fmtDistance=m=>(m/1000).toFixed(2);
 const today=new Date().toLocaleDateString('en-CA');
 const bikeRides=rides.filter(r=>(r.bikeId||'bike-default')===selectedBikeId);
 const todays=bikeRides.filter(r=>new Date(r.startAt).toLocaleDateString('en-CA')===today);
 const filteredRides=bikeRides.filter(r=>{const d=new Date(r.startAt),now=new Date();if(historyFilter==='today')return d.toLocaleDateString('en-CA')===today;if(historyFilter==='month')return d.getMonth()===now.getMonth()&&d.getFullYear()===now.getFullYear();if(historyFilter==='week'){const start=new Date(now);start.setHours(0,0,0,0);start.setDate(start.getDate()-6);return d>=start}if(historyFilter==='custom'){const day=d.toLocaleDateString('en-CA');return (!historyFrom||day>=historyFrom)&&(!historyTo||day<=historyTo)}return true});
 const todayKm=todays.reduce((s,r)=>s+(r.distanceM||0),0)/1000+(activeRide?(activeRide.distanceM||0)/1000:0);
 const todaySeconds=todays.reduce((s,r)=>s+(r.seconds||0),0)+ (activeRide?elapsed:0);
 const allDistanceM=bikeRides.reduce((s,r)=>s+(r.distanceM||0),0)+(activeRide?(activeRide.distanceM||0):0);
 const allSeconds=bikeRides.reduce((s,r)=>s+(r.seconds||0),0)+(activeRide?elapsed:0);
 const allRideCount=bikeRides.length+(activeRide?1:0);
 const averageRideKm=allRideCount?allDistanceM/1000/allRideCount:0;
 const averageSpeed=allSeconds?allDistanceM/allSeconds*3.6:0;
 const reportRides=bikeRides.filter(r=>{const day=new Date(r.startAt).toLocaleDateString('en-CA');return (!reportFrom||day>=reportFrom)&&(!reportTo||day<=reportTo)});
 const reportDistanceM=reportRides.reduce((s,r)=>s+(r.distanceM||0),0);
 const reportSeconds=reportRides.reduce((s,r)=>s+(r.seconds||0),0);
 const reportCount=reportRides.length;
 const reportMaxSpeed=Math.max(0,...reportRides.map(r=>r.maxSpeed||0));
 const reportAvgSpeed=reportSeconds?reportDistanceM/reportSeconds*3.6:0;
 const reportDailyGroups=Object.values(reportRides.reduce((groups,r)=>{const key=new Date(r.startAt).toLocaleDateString('en-CA');if(!groups[key])groups[key]={date:key,count:0,distanceM:0,seconds:0,maxSpeed:0};const g=groups[key];g.count++;g.distanceM+=r.distanceM||0;g.seconds+=r.seconds||0;g.maxSpeed=Math.max(g.maxSpeed,r.maxSpeed||0);return groups;},{})).sort((a,b)=>b.date.localeCompare(a.date));
 const mapUrl=ride=>{const pts=ride.points||[];if(!pts.length)return '';const first=pts[0],last=pts[pts.length-1],middle=pts.slice(1,-1),waypoints=[];const count=Math.min(8,middle.length);for(let i=0;i<count;i++){const index=Math.round((i+1)*(middle.length/(count+1))-1);const p=middle[Math.max(0,index)];waypoints.push(p.lat+','+p.lng)}return 'https://www.google.com/maps/dir/?api=1&origin='+first.lat+','+first.lng+'&destination='+last.lat+','+last.lng+(waypoints.length?'&waypoints='+waypoints.join('%7C'):'')+'&travelmode=driving'};
 const openMap=ride=>{const url=mapUrl(ride);if(url)window.open(url,'_blank','noopener,noreferrer')};
 const deleteRide=ride=>{if(!window.confirm('এই রাইডের তথ্য স্থায়ীভাবে মুছে ফেলবেন?'))return;setRides(old=>old.filter(item=>item.id!==ride.id));if(expandedRide===ride.id)setExpandedRide(null)};
 const exportCsv=()=>{const rows=[['Date','Start','End','Distance KM','Duration seconds','Max speed KMH','Tracking mode','Google Maps URL'],...filteredRides.map(r=>[new Date(r.startAt).toLocaleDateString(),new Date(r.startAt).toLocaleTimeString(),r.endAt?new Date(r.endAt).toLocaleTimeString():'',fmtDistance(r.distanceM||0),r.seconds||0,r.maxSpeed||0,r.mode,mapUrl(r)])];const csv=rows.map(row=>row.map(v=>'"'+String(v??'').replace(/"/g,'""')+'"').join(',')).join('\n');const a=document.createElement('a'),url=URL.createObjectURL(new Blob(['\ufeff'+csv],{type:'text/csv;charset=utf-8'}));a.href=url;a.download='Bike-Tracker-Ride-History.csv';a.click();URL.revokeObjectURL(url)};
 const exportReportCsv=()=>{const rows=[['Report From',reportFrom||'All time'],['Report To',reportTo||'All time'],['Total Rides',reportCount],['Total Distance KM',(reportDistanceM/1000).toFixed(2)],['Total Ride Time Seconds',reportSeconds],['Average Speed KMH',reportAvgSpeed.toFixed(1)],['Top Speed KMH',Math.round(reportMaxSpeed)],[],['Date','Start','End','Distance KM','Duration seconds','Max speed KMH','Tracking mode','Google Maps URL'],...reportRides.map(r=>[new Date(r.startAt).toLocaleDateString(),new Date(r.startAt).toLocaleTimeString(),r.endAt?new Date(r.endAt).toLocaleTimeString():'',fmtDistance(r.distanceM||0),r.seconds||0,r.maxSpeed||0,r.mode,mapUrl(r)])];const csv=rows.map(row=>row.map(v=>'"'+String(v??'').replace(/"/g,'""')+'"').join(',')).join('\n');const a=document.createElement('a'),url=URL.createObjectURL(new Blob(['\ufeff'+csv],{type:'text/csv;charset=utf-8'}));a.href=url;a.download='Bike-Tracker-Report-'+(reportFrom||'start')+'-to-'+(reportTo||'latest')+'.csv';a.click();URL.revokeObjectURL(url)};

 const addFuelLog=()=>{const litres=Number(fuelLitres),odo=Number(fuelOdometer);if(!litres||litres<=0||!odo||odo<=0)return;setFuelLogs(old=>[{id:'fuel-'+Date.now(),bikeId:selectedBikeId,date:new Date().toISOString(),litres,cost:litres*fuelPrice,odometerKm:odo,fullTank:fuelFullTank},...old]);setFuelLitres('');setFuelOdometer('')};
 const bikeFuelLogs=fuelLogs.filter(x=>x.bikeId===selectedBikeId).sort((a,b)=>new Date(a.date)-new Date(b.date));
 const fullTankLogs=bikeFuelLogs.filter(x=>x.fullTank&&Number(x.odometerKm)>0);
 const fuelMileageData=fullTankLogs.slice(1).reduce((acc,x,i)=>{const prev=fullTankLogs[i],litres=Number(x.litres)||0,km=Number(x.odometerKm)-Number(prev.odometerKm);if(km>0&&litres>0){acc.km+=km;acc.litres+=litres}return acc},{km:0,litres:0});
 const actualMileage=fuelMileageData.litres?fuelMileageData.km/fuelMileageData.litres:0;
 const addMaintenance=()=>{if(!maintenanceName.trim()||!maintenanceDate)return;setMaintenance(old=>[{id:'maint-'+Date.now(),bikeId:selectedBikeId,name:maintenanceName.trim(),dueDate:maintenanceDate,done:false},...old]);setMaintenanceName('');setMaintenanceDate('')};
 const backupData=()=>{const payload={version:1,createdAt:new Date().toISOString(),rides,motorcycles,fuelLogs,maintenance,fuelPrice,speedLimit};const url=URL.createObjectURL(new Blob([JSON.stringify(payload,null,2)],{type:'application/json'}));const a=document.createElement('a');a.href=url;a.download='Smart-Office-Bike-Backup-'+today+'.json';a.click();URL.revokeObjectURL(url)};
 const restoreData=e=>{const file=e.target.files&&e.target.files[0];if(!file)return;const reader=new FileReader();reader.onload=()=>{try{const d=JSON.parse(reader.result);if(!Array.isArray(d.rides)||!Array.isArray(d.motorcycles))throw new Error('Invalid');setRides(d.rides);setMotorcycles(d.motorcycles);setFuelLogs(Array.isArray(d.fuelLogs)?d.fuelLogs:[]);setMaintenance(Array.isArray(d.maintenance)?d.maintenance:[]);if(d.fuelPrice)setFuelPrice(d.fuelPrice);if(d.speedLimit)setSpeedLimit(d.speedLimit);alert('Backup restored successfully.')}catch(err){alert('Invalid backup file.')}};reader.readAsText(file);e.target.value=''};
 const addMotorcycle=()=>{const name=window.prompt('Motorcycle name / model?');if(!name||!name.trim())return;const bike={id:'bike-'+Date.now(),name:name.trim()};setMotorcycles(old=>[...old,bike]);setSelectedBikeId(bike.id)};
 const reportGraphMax=Math.max(1,...reportDailyGroups.map(x=>x.distanceM));
 const monthly=bikeRides.filter(r=>new Date(r.startAt).getMonth()===new Date().getMonth()&&new Date(r.startAt).getFullYear()===new Date().getFullYear()).reduce((s,r)=>s+(r.distanceM||0),0)/1000;
 return <div className="bikeApp">
  <header className="bikeHeader"><button className="bikeBack" onClick={onBack}>‹</button><div className="bikeBrand"><span className="bikeBrandIcon"><Bike size={23}/></span><span><b>Bike Tracker</b><small>Smart Office • Motorcycle</small></span></div><span className={'gpsBadge '+(gpsStatus==='GPS connected'?'gpsOn':'')}><span/>{gpsStatus}</span></header>
  <main className="bikeMain">
   <section className="bikeHero"><div className="bikeHeroGlow"/><div className="bikeHeroTop"><span>YOUR RIDE TODAY</span><span className="bikeLive">{activeRide?'● RECORDING':'● READY'}</span></div><div className="bikeKm">{todayKm.toFixed(2)} <small>KM</small></div><div className="bikeHeroStats"><div><Timer size={17}/><b>{fmtTime(todaySeconds)}</b><small>Ride time</small></div><div><Route size={17}/><b>{todays.length+(activeRide?1:0)}</b><small>Rides today</small></div><div><Gauge size={17}/><b>{Math.round(speed)}</b><small>Speed km/h</small></div></div></section>
   <div className="bikeActionRow"><button className="bikeStart" disabled={!!activeRide} onClick={()=>beginRide('manual')}><Play size={19} fill="currentColor"/> START RIDE</button><button className="bikeStop" disabled={!activeRide} onClick={finishRide}><Square size={17} fill="currentColor"/> STOP</button></div>
   <div className="bikeMode"><div><b>Automatic ride detection</b><small>Detects movement using GPS when this screen is active</small></div><button className={'bikeSwitch '+(autoMode?'on':'')} onClick={()=>setAutoMode(v=>!v)} aria-label="Toggle automatic ride detection"><i/></button></div>
   {gpsError&&<div className="bikeError"><MapPin size={17}/>{gpsError}</div>}
   <div className="bikeTabs"><button className={tab==='ride'?'active':''} onClick={()=>setTab('ride')}>Ride</button><button className={tab==='history'?'active':''} onClick={()=>setTab('history')}>History</button><button className={tab==='report'?'active':''} onClick={()=>setTab('report')}>Reports</button><button className={tab==='settings'?'active':''} onClick={()=>setTab('settings')}>Settings</button></div>
   {tab==='ride'&&<><div className="bikeSectionTitle">Current ride <span>{activeRide?'In progress':'Not started'}</span></div><div className="bikeMetrics"><div><span><Route size={17}/></span><b>{fmtDistance(distance)} km</b><small>Distance</small></div><div><span><Timer size={17}/></span><b>{fmtTime(elapsed)}</b><small>Duration</small></div><div><span><TrendingUp size={17}/></span><b>{Math.round(activeRide?.maxSpeed||speed)} km/h</b><small>Max speed</small></div></div><div className="bikeMapCard"><BikeRouteMap ride={activeRide}/><p>ম্যাপে GPS রুট, সবুজ Start এবং লাল End পয়েন্ট দেখা যাবে। Map tiles-এর জন্য ইন্টারনেট প্রয়োজন।</p></div></>}
   {tab==='history'&&<><div className="bikeSectionTitle">Ride history <button onClick={exportCsv} className="bikeExport"><Download size={15}/> Excel / CSV</button></div><div className="bikeHistoryFilters">{[['today','Today'],['week','7 Days'],['month','This Month'],['custom','Date Range'],['all','All']].map(([key,label])=><button key={key} className={historyFilter===key?'active':''} onClick={()=>setHistoryFilter(key)}>{label}</button>)}<span>{filteredRides.length} rides</span></div>{historyFilter==='custom'&&<div className="bikeDateRange"><label>From<input type="date" value={historyFrom} onChange={e=>setHistoryFrom(e.target.value)}/></label><label>To<input type="date" value={historyTo} min={historyFrom||undefined} onChange={e=>setHistoryTo(e.target.value)}/></label><button onClick={()=>{setHistoryFrom('');setHistoryTo('')}}>Clear</button></div>}{filteredRides.length?filteredRides.map(r=><div className="bikeRideItem" key={r.id}><div className="bikeRideIcon"><Bike size={20}/></div><div className="bikeRideInfo"><b>{new Date(r.startAt).toLocaleDateString()} • {new Date(r.startAt).toLocaleTimeString([], {hour:'2-digit',minute:'2-digit'})}</b><small>{r.endAt?new Date(r.endAt).toLocaleTimeString([], {hour:'2-digit',minute:'2-digit'}):'In progress'} • {r.mode==='automatic'?'Automatic':'Manual'}</small></div><div className="bikeRideValue"><b>{fmtDistance(r.distanceM||0)} km</b><button onClick={()=>openMap(r)} disabled={!r.points?.length}><Navigation size={13}/> Map</button><button onClick={()=>setExpandedRide(expandedRide===r.id?null:r.id)}>{expandedRide===r.id?'Less':'Details'}</button><button className="bikeDeleteRide" onClick={()=>deleteRide(r)} aria-label="Delete ride"><Trash2 size={13}/></button></div>{expandedRide===r.id&&<div className="bikeRideDetails"><div><small>Duration</small><b>{fmtTime(r.seconds||0)}</b></div><div><small>Average speed</small><b>{r.seconds>0?((r.distanceM||0)/r.seconds*3.6).toFixed(1):'0.0'} km/h</b></div><div><small>Maximum speed</small><b>{Math.round(r.maxSpeed||0)} km/h</b></div><div><small>GPS points</small><b>{(r.points||[]).length}</b></div><button onClick={()=>openMap(r)} disabled={!r.points?.length}><Navigation size={14}/> Open detailed route in Google Maps</button></div>}</div>):<div className="bikeEmpty"><Route size={30}/><b>No rides in this period</b><small>Choose another date filter to see your ride history.</small></div>}</>}
   {tab==='report'&&<><div className="bikeSectionTitle">Ride reports</div><div className="bikeDateRange bikeReportRange"><label>From<input type="date" value={reportFrom} onChange={e=>setReportFrom(e.target.value)}/></label><label>To<input type="date" min={reportFrom||undefined} value={reportTo} onChange={e=>setReportTo(e.target.value)}/></label><button onClick={()=>{setReportFrom('');setReportTo('')}}>All Time</button></div><div className="bikeReportGrid"><div><small>SELECTED PERIOD</small><b>{(reportDistanceM/1000).toFixed(2)} km</b><span>{reportCount} rides</span></div><div><small>TOTAL RIDE TIME</small><b>{fmtTime(reportSeconds)}</b><span>Selected period</span></div><div><small>AVERAGE PER RIDE</small><b>{(reportCount?reportDistanceM/1000/reportCount:0).toFixed(2)} km</b><span>Selected period</span></div><div><small>AVERAGE SPEED</small><b>{reportAvgSpeed.toFixed(1)} km/h</b><span>Distance ÷ ride time</span></div><div><small>TOP SPEED</small><b>{Math.round(reportMaxSpeed)} km/h</b><span>Selected period</span></div><div><small>ALL-TIME DISTANCE</small><b>{(allDistanceM/1000).toFixed(2)} km</b><span>{allRideCount} recorded rides</span></div><div><small>ALL-TIME RIDE TIME</small><b>{fmtTime(allSeconds)}</b><span>Combined duration</span></div></div><div className="bikeSectionTitle bikeDailyTitle">Daily breakdown <span>{reportDailyGroups.length} days</span></div><div className="bikeDailyReport">{reportDailyGroups.length?reportDailyGroups.map(day=><div className="bikeDailyRow" key={day.date}><div className="bikeDailyDate"><b>{day.date}</b><small>{day.count} rides</small></div><div><b>{(day.distanceM/1000).toFixed(2)} km</b><small>Distance</small></div><div><b>{fmtTime(day.seconds)}</b><small>Ride time</small></div><div><b>{Math.round(day.maxSpeed)} km/h</b><small>Top speed</small></div></div>):<div className="bikeEmpty">No daily data for this date range.</div>}</div><div className="bikeSectionTitle bikeDailyTitle">Distance chart</div><div className="bikeDailyChart">{reportDailyGroups.slice(0,14).reverse().map(day=><div className="bikeChartBar" key={day.date}><span style={{height:Math.max(5,day.distanceM/reportGraphMax*100)+"%"}}/><small>{day.date.slice(5)}</small></div>)}</div><div className="bikeReportActions"><button className="bikeWideExport" onClick={exportReportCsv}><Download size={17}/> Export selected report to Excel-compatible CSV</button><button className="bikeWideExport" onClick={()=>window.print()}>Export / Save Report as PDF</button></div></>}
   {tab==='settings'&&<><div className="bikeSectionTitle">Motorcycles & Settings</div><div className="bikeSetting"><div><b>Active motorcycle</b><small>Ride and maintenance profile</small></div><select value={selectedBikeId} onChange={e=>setSelectedBikeId(e.target.value)}>{motorcycles.map(b=><option key={b.id} value={b.id}>{b.name}</option>)}</select><button onClick={addMotorcycle}>+ Add</button></div><div className="bikeSetting"><div><b>Speed alert limit</b><small>Alert when speed exceeds km/h</small></div><input type="number" min="20" max="200" value={speedLimit} onChange={e=>setSpeedLimit(Math.max(20,Number(e.target.value)))}/></div><div className="bikeSetting"><Fuel size={20}/><div><b>Fuel price</b><small>BDT per litre</small></div><input type="number" min="0" value={fuelPrice} onChange={e=>setFuelPrice(Math.max(0,Number(e.target.value)))}/></div><div className="bikeSectionTitle">Fuel & Mileage</div><div className="bikeSetting"><div><b>Add refuel</b><small>Litres and current odometer</small></div><input type="number" min="0" step="0.1" placeholder="Litres" value={fuelLitres} onChange={e=>setFuelLitres(e.target.value)}/><input type="number" min="0" step="0.1" placeholder="Odometer km" value={fuelOdometer} onChange={e=>setFuelOdometer(e.target.value)}/><label><input type="checkbox" checked={fuelFullTank} onChange={e=>setFuelFullTank(e.target.checked)}/> Full tank</label><button onClick={addFuelLog}>Save</button></div><div className="bikeReportGrid"><div><small>TOTAL FUEL</small><b>{fuelLogs.filter(x=>x.bikeId===selectedBikeId).reduce((a,x)=>a+x.litres,0).toFixed(2)} L</b></div><div><small>FUEL COST</small><b>৳{fuelLogs.filter(x=>x.bikeId===selectedBikeId).reduce((a,x)=>a+x.cost,0).toFixed(0)}</b></div><div><small>FULL-TANK MILEAGE</small><b>{actualMileage.toFixed(1)} km/L</b></div></div><div className="bikeSectionTitle">Maintenance reminders</div><div className="bikeSetting"><input placeholder="Service / Oil / Tyre" value={maintenanceName} onChange={e=>setMaintenanceName(e.target.value)}/><input type="date" value={maintenanceDate} onChange={e=>setMaintenanceDate(e.target.value)}/><button onClick={addMaintenance}>Add</button></div>{maintenance.filter(x=>x.bikeId===selectedBikeId).map(m=><div className="bikeSetting" key={m.id}><div><b>{m.name}</b><small>Due: {m.dueDate}</small></div><button onClick={()=>setMaintenance(old=>old.map(x=>x.id===m.id?{...x,done:!x.done}:x))}>{m.done?'Completed':'Mark done'}</button></div>)}<div className="bikeSectionTitle">Backup & Restore</div><div className="bikeSetting"><button onClick={backupData}>Export Full Backup (JSON)</button><label className="bikeRestoreLabel">Restore Backup<input type="file" accept="application/json,.json" onChange={restoreData}/></label></div><div className="bikeSectionTitle">Android Tracking Permissions</div><div className="bikeSetting"><button onClick={()=>Capacitor.isNativePlatform()?BikeTracking.openLocationSettings():alert("Available in the installed Android app.")}>Location / Background Permission</button><button onClick={()=>Capacitor.isNativePlatform()?BikeTracking.openBatterySettings():alert("Available in the installed Android app.")}>Battery Optimization</button></div><div className="bikeSettingNote">Background GPS uses Android foreground service while a ride is active. Android may still stop tracking if the user force-stops the app or disables location permission. Allow location and notification permissions. Battery settings can vary by phone brand.</div></>}
   <div className="bikeFooter">Smart Office • Bike Tracker • Your rides stay on this device</div>
  </main>
 </div>
}
function SmartOfficeSplash(){
 React.useEffect(()=>{const timer=window.setTimeout(()=>{const node=document.getElementById('smart-office-splash');if(node){node.classList.add('splashExit');window.setTimeout(()=>node.remove(),220)}} ,2000);return()=>window.clearTimeout(timer)},[]);
 const streams=Array.from({length:22},(_,i)=>({id:i,left:(i*47)%100,delay:-((i*137)%2000)/1000,duration:1.5+((i*29)%18)/10,text:Array.from({length:12},(_,j)=>((i+j*7)%2?'1':'0')).join('')}));
 return <div id="smart-office-splash" className="smartSplash" aria-label="Smart Office loading"><div className="matrixRain">{streams.map(s=><div key={s.id} className="matrixStream" style={{left:s.left+'%',animationDelay:s.delay+'s',animationDuration:s.duration+'s'}}>{s.text.split('').map((ch,i)=><span key={i}>{ch}</span>)}</div>)}</div><div className="binaryTunnel">{Array.from({length:28},(_,i)=><span key={i} style={{'--i':i}}>{i%2?'010101':'101010'}</span>)}</div><div className="splashGlow"></div><div className="splashLogoStage"><div className="splashLogo"><img src="/smart-office-logo-1.png" alt="Smart Office logo"/></div></div><div className="splashCaption">SMART OFFICE</div></div>
}
function App(){
 const [tab,setTab]=React.useState('Home');
 const [showSplash,setShowSplash]=React.useState(true);
 const [excel,setExcel]=React.useState(false);
 const [calculator,setCalculator]=React.useState(false);
 const [bikeTracker,setBikeTracker]=React.useState(false);
 const [officeFile,setOfficeFile]=React.useState(null);
 const [selectedTool,setSelectedTool]=React.useState(null);
 const [recentFiles,setRecentFiles]=React.useState(()=>{try{return JSON.parse(localStorage.getItem('smartoffice-recent-files')||'[]')}catch{return[]}});
 const cards=[{name:'Word',desc:'Create and edit documents',icon:FileText,color:'#2463c7',ext:'docx'},{name:'Excel',desc:'Spreadsheets and calculations',icon:Table2,color:'#16834b',ext:'xlsx'},{name:'PowerPoint',desc:'Presentations and slides',icon:Presentation,color:'#d66a28',ext:'pptx'},{name:'PDF Reader',desc:'View PDF documents',icon:File,color:'#c83b45',ext:'pdf'},{name:'Calculator',desc:'Fast everyday calculations',icon:Calculator,color:'#2563eb',ext:'calc'},{name:'Bike Tracker',desc:'Motorcycle GPS ride tracking',icon:Bike,color:'linear-gradient(145deg,#ff8a00,#ef4444)',ext:'bike'}];
 const refreshRecent=()=>{
  let list=[];try{list=JSON.parse(localStorage.getItem('smartoffice-recent-files')||'[]')}catch{}
  try{const wb=JSON.parse(localStorage.getItem('smartoffice-workbook')||'null');if(wb&&wb.workbookName){const item={id:'excel-current',name:(wb.workbookName||'Book1')+'.xlsx',type:'Excel',updatedAt:wb.savedAt||new Date().toISOString()};list=[item,...list.filter(x=>x.id!=='excel-current')];localStorage.setItem('smartoffice-recent-files',JSON.stringify(list));}}catch{}
  setRecentFiles(list);
 };
 React.useEffect(()=>{refreshRecent()},[]);
 React.useEffect(()=>{const timer=window.setTimeout(()=>setShowSplash(false),2000);return()=>window.clearTimeout(timer)},[]);
 if(showSplash)return <SmartOfficeSplash/>;
 if(calculator)return <CalculatorEditor onBack={()=>setCalculator(false)}/>;
 if(bikeTracker)return <BikeTracker onBack={()=>setBikeTracker(false)}/>;
 if(excel)return <ExcelEditor onBack={()=>{setExcel(false);refreshRecent()}}/>;
 if(officeFile?.type==='Word')return <WordEditor file={officeFile} onBack={()=>{setOfficeFile(null);refreshRecent()}}/>;
 if(officeFile?.type==='PowerPoint')return <PowerPointEditor file={officeFile} onBack={()=>{setOfficeFile(null);refreshRecent()}}/>;
 const openTool=tool=>{if(tool==='Calculator'){setCalculator(true);return}if(tool==='Bike Tracker'){setBikeTracker(true);return}setSelectedTool(tool);refreshRecent()};
 const openRecent=item=>{setSelectedTool(null);if(item.type==='Excel')setExcel(true);else if(item.type==='Word'||item.type==='PowerPoint')setOfficeFile(item);else alert('PDF Reader পরবর্তী ধাপে যুক্ত হবে।')};
 const createOfficeFile=type=>{const ext=type==='Word'?'docx':'pptx';const item={id:type+'-'+Date.now(),name:(type==='Word'?'Document':'Presentation')+'-'+new Date().toLocaleDateString('en-CA')+'.'+ext,type,updatedAt:new Date().toISOString()};const list=[item,...recentFiles.filter(x=>x.id!==item.id)];setRecentFiles(list);localStorage.setItem('smartoffice-recent-files',JSON.stringify(list));setSelectedTool(null);setOfficeFile(item)};
 const toolFiles=recentFiles.filter(f=>f.type===selectedTool);
 return <div className="app">
  <header><div className="brand"><div className="brandIcon">S</div><div><b>Smart Office</b><small>Work smarter, anywhere</small></div></div><button className="round" onClick={()=>setSelectedTool(null)}><Settings size={20}/></button></header>
  {selectedTool? <main>
   <div className="sectionTitle" style={{marginTop:20}}><button className="textBtn" onClick={()=>setSelectedTool(null)}>‹ Home</button><h2>{selectedTool} — Recent Files</h2><span>{toolFiles.length} files</span></div>
   <p style={{color:'#64748b',fontSize:13,margin:'8px 0 18px'}}>আগে খোলা {selectedTool} ফাইলগুলো এখানে দেখাবে। ফাইল নির্বাচন করলে সেটি খুলবে।</p>
   {toolFiles.length? <div style={{display:'grid',gap:10}}>{toolFiles.map(file=><button key={file.id} onClick={()=>openRecent(file)} style={{display:'flex',alignItems:'center',gap:12,textAlign:'left',padding:14,border:'1px solid #e2e8f0',borderRadius:14,background:'#fff',color:'#172033'}}><div className="quickIcon"><FileSpreadsheet size={22}/></div><span style={{flex:1,minWidth:0}}><b style={{display:'block',overflow:'hidden',textOverflow:'ellipsis',whiteSpace:'nowrap'}}>{file.name}</b><small style={{color:'#64748b'}}>শেষ খোলা: {new Date(file.updatedAt||Date.now()).toLocaleString()}</small></span><span>›</span></button>)}</div>:<div className="quick" style={{marginTop:18}}><div className="quickIcon"><FolderOpen/></div><div><b>কোনো Recent {selectedTool} ফাইল নেই</b><small>নতুন ফাইল তৈরি বা ওপেন করলে এখানে দেখা যাবে।</small></div></div>}
   {selectedTool==='Excel'&&<button className="openBtn" style={{marginTop:18,padding:'12px 18px'}} onClick={()=>{setExcel(true);setSelectedTool(null)}}>＋ নতুন Workbook / Excel খুলুন</button>}
   {selectedTool==='Word'&&<button className="openBtn" style={{marginTop:18,padding:'12px 18px'}} onClick={()=>createOfficeFile('Word')}>＋ নতুন Word Document তৈরি করুন</button>}
   {selectedTool==='PowerPoint'&&<button className="openBtn" style={{marginTop:18,padding:'12px 18px'}} onClick={()=>createOfficeFile('PowerPoint')}>＋ নতুন Presentation তৈরি করুন</button>}
   <button className="textBtn" style={{display:'block',marginTop:22}} onClick={()=>setSelectedTool(null)}>Dashboard-এ ফিরে যান</button>
  </main>:<><main><section className="welcome"><div><span className="eyebrow">YOUR PRODUCTIVITY SPACE</span><h1>Everything you need<br/>to get work done.</h1><p>Create, edit and manage your office files in one place.</p></div><div className="welcomeMark">SO</div></section><div className="sectionTitle"><h2>Office tools</h2><span>4 tools</span></div><div className="grid">{cards.map(c=><button className="tool" key={c.name} onClick={()=>openTool(c.name)}><div className="toolIcon" style={{background:c.color}}><c.icon size={25}/></div><b>{c.name}</b><small>{c.desc}</small><span className="arrow">›</span></button>)}</div><div className="sectionTitle recentTitle"><h2>Quick access</h2><button className="textBtn" onClick={()=>{setSelectedTool('Excel');refreshRecent()}}>See all</button></div><div className="quick"><div className="quickIcon"><FolderOpen/></div><div><b>Open a document</b><small>Browse files on your device</small></div><button className="openBtn" onClick={()=>document.getElementById('fileInput').click()}>Open</button><input id="fileInput" type="file" hidden accept=".doc,.docx,.xls,.xlsx,.ppt,.pptx,.pdf" onChange={e=>{const f=e.target.files?.[0];if(!f)return;const ext=f.name.split('.').pop().toLowerCase(),type=['xls','xlsx','csv'].includes(ext)?'Excel':['doc','docx'].includes(ext)?'Word':['ppt','pptx'].includes(ext)?'PowerPoint':'PDF Reader';const item={id:type+'-'+f.name+'-'+f.lastModified,name:f.name,type,updatedAt:new Date().toISOString()};const list=[item,...recentFiles.filter(x=>x.id!==item.id)];setRecentFiles(list);localStorage.setItem('smartoffice-recent-files',JSON.stringify(list));setSelectedTool(type);e.target.value=''}}/></div><div className="notice"><ShieldCheck size={18}/><span>Your files stay on your device. Offline-ready design.</span></div></main><nav>{[['Home',FolderOpen],['Recent',Clock3],['Files',FileText]].map(([n,I])=><button className={tab===n?'active':''} onClick={()=>{setTab(n);if(n!=='Home')setSelectedTool(n==='Recent'?'Excel':'Excel')}}><I size={20}/><span>{n}</span></button>)}</nav>{tab!=='Home'&&<div className="toast" onClick={()=>{setTab('Home');setSelectedTool(null)}}>{tab} — tap to return Home.</div>}</>}
 </div>
}
createRoot(document.getElementById('root')).render(<App/>);