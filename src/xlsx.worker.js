import * as XLSX from 'xlsx';

self.onmessage = (event) => {
  try {
    const data = new Uint8Array(event.data.buffer);
    self.postMessage({type:'progress',message:'Excel ফাইল আনজিপ ও বিশ্লেষণ হচ্ছে…'});
    const workbook = XLSX.read(data, {
      type:'array',
      cellFormula:true,
      dense:false,
      cellStyles:false,
      cellNF:false,
      cellHTML:false,
      bookFiles:false,
      bookVBA:false,
      WTF:false
    });
    const names = workbook.SheetNames || [];
    const activeIndex = workbook.Workbook?.Views?.[0]?.activeTab;
    const activeSheet = Number.isInteger(activeIndex) ? names[activeIndex] : names[0];
    if (!names.length) throw new Error('কোনো Worksheet পাওয়া যায়নি');
    const sheets = {};
    names.forEach((name,index) => {
      self.postMessage({type:'progress',message:(index+1)+'/'+names.length+'টি Worksheet প্রস্তুত হচ্ছে…'});
      const ws = workbook.Sheets[name] || {};
      const cells = {};
      for (const address of Object.keys(ws)) {
        if (!/^[A-Z]{1,3}[1-9][0-9]*$/.test(address)) continue;
        const cell = ws[address];
        if (!cell) continue;
        if (cell.f) cells[address] = '='+cell.f;
        else if (cell.v !== undefined && cell.v !== null) cells[address] = typeof cell.v === 'number' ? cell.v : String(cell.v);
      }
      sheets[name] = {cells,cellStyles:{}};
    });
    self.postMessage({type:'success',sheets,activeSheet});
  } catch (error) {
    self.postMessage({type:'error',message:error?.message || 'Unknown error'});
  }
};