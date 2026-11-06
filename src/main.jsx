import React from 'react';
import {createRoot} from 'react-dom/client';
import {FileText,Table2,Presentation,File,FolderOpen,Clock3,Settings,ShieldCheck,Grid2X2,Home,Folder,Eye,Plus,Database,ClipboardCheck,Save,Printer,Undo2,Redo2,Copy,Scissors,Clipboard,Paintbrush,Type,Bold,Italic,Underline,AlignLeft,AlignCenter,AlignRight,WrapText,Merge,Filter,ArrowDownUp,Search,Check,MessageSquare,Lock,BarChart3,Image,Link,Calculator,Percent,Coins,CalendarDays,SortAsc,FileSpreadsheet,X,ChevronDown} from 'lucide-react';
import './style.css';

const ribbonData = {
 Home: [
  {group:'Clipboard',items:[['Paste',Clipboard],['Scissors',Scissors],['Copy',Copy],['Format Painter',Paintbrush]]},
  {group:'Font',items:[['Font',Type],['Bold',Bold],['Italic',Italic],['Underline',Underline],['Font Color',Type]]},
  {group:'Alignment',items:[['Align Left',AlignLeft],['Center',AlignCenter],['Align Right',AlignRight],['Wrap Text',WrapText],['Merge Cells',Merge]]},
  {group:'Number',items:[['Number Format',Calculator],['Currency',Coins],['Percent',Percent],['Date',CalendarDays]]},
  {group:'Cells',items:[['Insert',Plus],['Delete',X],['Format',Grid2X2]]},
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
 let s=String(value).trim().replace(/[৳$€£,\\s]/g,'');
 if(s.endsWith('%')){const n=Number(s.slice(0,-1));return Number.isFinite(n)?n/100:null}
 if(/^\\(.*\\)$/.test(s))s='-'+s.slice(1,-1);
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
 const getRange=(token)=>{
  const t=token.trim().replace(/\\$/g,'');
  if(!t.includes(':'))return [calculateCell(t.toUpperCase(),cells,stack)];
  const parts=t.split(':');if(parts.length!==2)throw Error('Bad range');
  const a=cellToPosition(parts[0]),b=cellToPosition(parts[1]);if(!a||!b)throw Error('Bad reference');
  const values=[];
  for(let row=Math.min(a.row,b.row);row<=Math.max(a.row,b.row);row++)
   for(let col=Math.min(a.col,b.col);col<=Math.max(a.col,b.col);col++)
    values.push(calculateCell(positionToCell(col,row),cells,stack));
  return values;
 };
 try{
  const fn=source.trim().match(/^(SUM|AVERAGE|COUNT|MAX|MIN)\\s*\\((.*)\\)$/i);
  if(fn){
   const args=fn[2].split(',').flatMap(getRange);
   const nums=args.map(v=>typeof v==='number'?v:numericValue(v)).filter(v=>v!==null);
   const name=fn[1].toUpperCase();
   if(name==='COUNT')return nums.length;
   if(name==='SUM')return nums.reduce((a,b)=>a+b,0);
   if(name==='AVERAGE')return nums.length?nums.reduce((a,b)=>a+b,0)/nums.length:0;
   if(name==='MAX')return nums.length?Math.max(...nums):0;
   return nums.length?Math.min(...nums):0;
  }
  const expression=source.trim().replace(/([A-Z]+\\$?[1-9][0-9]*)/gi,(ref)=>{
   const normalized=ref.replace(/\\$/g,'').toUpperCase();
   const v=calculateCell(normalized,cells,stack);
   const n=typeof v==='number'?v:numericValue(v);
   if(n===null)throw Error('Non-numeric reference');
   return String(n);
  });
  if(!/^[0-9+\\-*/().\\s]+$/.test(expression))throw Error('Invalid formula');
  const tokens=expression.match(/\\d+(?:\\.\\d+)?|[()+\\-*/]/g)||[];
  let i=0;
  const primary=()=>{
   const t=tokens[i++];
   if(t==='('){const v=add();if(tokens[i++]!==')')throw Error('Missing )');return v}
   if(t==='+')return primary();
   if(t==='-')return -primary();
   const n=Number(t);if(!Number.isFinite(n))throw Error('Bad number');return n;
  };
  const mul=()=>{let v=primary();while(tokens[i]==='*'||tokens[i]==='/'){const op=tokens[i++],n=primary();v=op==='*'?v*n:v/n}return v};
  const add=()=>{let v=mul();while(tokens[i]==='+'||tokens[i]==='-'){const op=tokens[i++],n=mul();v=op==='+'?v+n:v-n}return v};
  const result=add();if(i!==tokens.length||!Number.isFinite(result))throw Error('Invalid result');return result;
 }catch{return '#VALUE!'}
}
function ExcelEditor({onBack}){
 const [active,setActive]=React.useState('Home');
 const [menuOpen,setMenuOpen]=React.useState(false);
 const [selected,setSelected]=React.useState('A1');
 const [formula,setFormula]=React.useState('');
 const [editingCell,setEditingCell]=React.useState(null);
 const [cells,setCells]=React.useState({});
 const [cellStyles,setCellStyles]=React.useState({});
 const [sheet,setSheet]=React.useState('Sheet1');
 const [notice,setNotice]=React.useState('');
 const [showFormula,setShowFormula]=React.useState(true);
 const [showGrid,setShowGrid]=React.useState(true);
 const [filter,setFilter]=React.useState(false);
 const colName=n=>{let s='';while(n>0){let r=(n-1)%26;s=String.fromCharCode(65+r)+s;n=Math.floor((n-1)/26)}return s};
 const notify=t=>{setNotice(t);setTimeout(()=>setNotice(''),1800)};
 const selectCell=(id)=>{setSelected(id);setFormula(cells[id]||'')};
 const updateCell=(id,value)=>setCells(p=>({...p,[id]:value}));
 const updateStyle=(patch)=>setCellStyles(p=>({...p,[selected]:{...(p[selected]||{}),...patch}}));
 const toggleStyle=(key)=>updateStyle({[key]:!cellStyles[selected]?.[key]});
 const action=(name)=>{
  if(name==='Save'||name==='Save As'){notify('Workbook save option is being prepared');return}
  if(name==='Open'){document.getElementById('excelFileInput')?.click();return}
  if(name==='Export PDF'||name==='Print'){notify(name+' option selected');return}
  if(name==='Gridlines'){setShowGrid(v=>!v);return}
  if(name==='Formula Bar'){setShowFormula(v=>!v);return}
  if(name==='Filter'){setFilter(v=>!v);return}
  if(name==='New Workbook'){setCells({});setSelected('A1');setFormula('');notify('New workbook created');return}
  if(name==='AutoSum'){const p=cellToPosition(selected);if(p){let start=p.row-1;while(start>1&&cells[positionToCell(p.col,start-1)]!==undefined&&cells[positionToCell(p.col,start-1)]!=='')start--;const end=p.row-1;const ref=end>=start?positionToCell(p.col,start)+':'+positionToCell(p.col,end):positionToCell(p.col,1)+':'+positionToCell(p.col,1);const value='=SUM('+ref+')';setFormula(value);updateCell(selected,value);notify('SUM formula added')}return}
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
 const rows=Array.from({length:24},(_,i)=>i+1),cols=Array.from({length:10},(_,i)=>i+1);
 return <div className="excelApp">
  <header className="excelTop"><button className="backBtn" onClick={onBack}>‹</button><div className="excelTitle"><b><Table2 size={19}/> Smart Excel</b><small>Workbook • {sheet}</small></div><button className="topIcon" onClick={()=>action('Save')}><Save size={19}/></button><button className="topIcon" onClick={()=>notify('More workbook options')}><Settings size={19}/></button></header>
  <div className="workbookBar"><span className="fileBadge"><FileSpreadsheet size={16}/></span><input aria-label="Workbook name" defaultValue="Book1" /><span className="savedState">● Saved locally</span></div>
  <div className="ribbonStrip">{ribbonData[active].map(group=><div className="ribbonGroup" key={group.group}><div className="ribbonButtons">{group.items.map(([label,Icon])=><button key={label} className="ribbonAction" onClick={()=>action(label)} title={label}><Icon size={17}/><span>{label}</span></button>)}</div><small>{group.group}</small></div>)}</div>
  {showFormula&&<div className="formulaBar"><span className="cellRef">{selected}</span><span className="fx">fx</span><input value={formula} onChange={e=>{setFormula(e.target.value);updateCell(selected,e.target.value)}} placeholder="Enter value or formula"/></div>}
  <div className="sheetScroll"><table className={showGrid?'sheetGrid':'sheetGrid noGrid'}><thead><tr><th className="corner"></th>{cols.map(c=><th key={c}>{colName(c)}</th>)}</tr></thead><tbody>{rows.map(r=><tr key={r}><th>{r}</th>{cols.map(c=>{const id=colName(c)+r;const style=cellStyles[id]||{};let display=cells[id]===undefined?'':String(calculateCell(id,cells));if(style.numberFormat==='currency'&&display!==''&&Number.isFinite(Number(display)))display='৳ '+Number(display).toFixed(2);if(style.numberFormat==='percent'&&display!==''&&Number.isFinite(Number(display)))display=(Number(display)*100).toFixed(2)+'%';return <td key={id} className={selected===id?'selectedCell':''}><input style={{fontWeight:style.bold?'700':'400',fontStyle:style.italic?'italic':'normal',textDecoration:style.underline?'underline':'none',textAlign:style.textAlign||'left',color:style.color||'inherit',backgroundColor:style.backgroundColor||'transparent'}} aria-label={id} value={editingCell===id?(cells[id]||''):display} onFocus={()=>{selectCell(id);setEditingCell(id)}} onBlur={()=>setEditingCell(null)} onChange={e=>{updateCell(id,e.target.value);if(selected===id)setFormula(e.target.value)}} onDoubleClick={()=>{selectCell(id);setEditingCell(id)}}/></td>})}</tr>)}</tbody></table></div>
  <div className="sheetTabs"><button className="sheetNav" onClick={()=>notify('Previous sheet')}>‹</button><button className="sheetTab activeSheet">{sheet}</button><button className="addSheet" onClick={()=>{setSheet('Sheet'+(Number(sheet.replace('Sheet',''))+1||2));notify('New sheet added')}}><Plus size={17}/></button><span className="sheetStatus">Ready</span></div>
  <button className="ribbonMenuButton" onClick={()=>setMenuOpen(v=>!v)} aria-label="Open ribbon menu"><Grid2X2 size={24}/></button>
  {menuOpen&&<><div className="menuScrim" onClick={()=>setMenuOpen(false)}></div><div className="ribbonMenuPanel"><div className="menuPanelTitle">Workbook Menu <button onClick={()=>setMenuOpen(false)}><X size={18}/></button></div>{menus.map(([name,Icon])=><button key={name} className={'menuChoice '+(active===name?'chosen':'')} onClick={()=>{setActive(name);setMenuOpen(false)}}><Icon size={19}/><span>{name}</span>{active===name&&<Check size={16}/>}</button>)}</div></>}
  <input id="excelFileInput" type="file" hidden accept=".xls,.xlsx,.csv" onChange={e=>e.target.files?.[0]&&notify('Selected '+e.target.files[0].name+' — import engine is next phase')}/>
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