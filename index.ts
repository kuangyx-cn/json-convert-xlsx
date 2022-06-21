import { utils, writeFile } from "xlsx";
import { WorkSheet, WorkBook } from "xlsx";


export default class JsonToXlsx{
  public originData: any;

  public ws: WorkSheet;
  public wb: WorkBook;

  constructor(data: any, sheetName = 'Sheet1'){
    this.originData = data;

    this.ws = utils.json_to_sheet(this.originData);
    this.wb = utils.book_new();

    utils.book_append_sheet(this.wb, this.ws, sheetName);
    
  }

  replaceHeader(obj: object): JsonToXlsx{
    
    const range = utils.decode_range(this.ws['!ref'] as any)

    for(let i = range.s.c; i <= range.e.c; i++){
      const h = utils.encode_col(i) + '1'
      obj[this.ws[h].v] && (this.ws[h].v = obj[this.ws[h].v])
    }
    return this;
  }

  download(fileName: string){
    fileName ??= 'excel' + Date.now();
    return writeFile(this.wb, fileName + ".xlsx")
  }
}
