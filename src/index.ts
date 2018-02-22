import _ = require('lodash');
import XLSX = require('@inspirationlabs/js-xlsx');

export class worksheet {
  protected name:string;
  protected headerColumns = [];
  protected rows = [];
  protected C = 0;
  protected R = 0;
  protected range = {s: {c:10000000, r:10000000}, e: {c:0, r:0 }};
  protected data = {};

  constructor(name) {
    this.name = name;
  }

  public datenum(v, date1904?): any {
  	if(date1904) v+=1462;
  	var epoch:any = Date.parse(v);
    let calcdate:any = new Date(Date.UTC(1899, 11, 30));
    let returnval:any = (epoch - calcdate) / (24 * 60 * 60 * 1000);
  	return returnval;
  }

  public setHeader(arr) {
    _.each(arr, (item) => {
      this.headerColumns.push(item);
    });
    _.each(this.headerColumns, (col) => {
      this.setCell(this.R, this.C, {v: col});
      this.C++;
    });
    this.R++;
    this.C = 0;
  }

  public addRow(row:any) {
    if(this.headerColumns.length) {
      _.each(this.headerColumns, (col) => {
        var cell: any = row[col];
        this.setCell(this.R, this.C, cell);
        this.C++;
      });
    } else {
      _.each(row, (cell) => {
        this.setCell(this.R, this.C, cell);
        this.C++;
      });
    }
    this.R++;
    this.C = 0;
  }

  public encodeCell(obj) {
    return XLSX.utils.encode_cell(obj);
  }

  protected setCell(R, C, cell) {
    var ws = {};
    if(this.range.s.r > R) this.range.s.r = R;
    if(this.range.s.c > C) this.range.s.c = C;
    if(this.range.e.r < R) this.range.e.r = R;
    if(this.range.e.c < C) this.range.e.c = C;
    var cell_ref = this.encodeCell({c:C, r:R});

    if(!_.isObject(cell)) {
      cell = {v: cell};
    }
    if(cell.v != null) {
      if(cell.v instanceof Date) {
				cell.t = 'n';
        if(cell.z == null) {
          cell.z = XLSX.SSF._table[14];
        }
				cell.v = this.datenum(cell.v);
      } else if(typeof cell.v === 'number') {
        cell.t = 'n';
      } else if(typeof cell.v === 'boolean') {
        cell.t = 'b';
      } else if(!cell.t) {
        cell.t = 's';
      }
      this.data[cell_ref] = cell;
    } else if(cell.f != null) {
      cell.t = 'f';
      this.data[cell_ref] = cell;
    }
  }

  protected write() {
    if(this.range.s.c < 10000000) {
      this.data['!ref'] = XLSX.utils.encode_range(this.range);
    }
  }
}

export class workbook {
  protected worksheets = [];

  public addSheet(worksheet) {
    this.worksheets.push(worksheet);
  }

  public write() {
    var wb = {
      SheetNames: [],
      Sheets: {}
    };
    _.each(this.worksheets, (sheet) => {
      sheet.write();
      wb.SheetNames.push(sheet.name);
      wb.Sheets[sheet.name] = sheet.data;
    });
    var wopts = { bookType:'xlsx', bookSST:false, type:'binary' };
    var output = XLSX.write(wb, wopts);
    return output;
  }

  public download(res, filename?) {
    var wbout = this.write();
    res.type('application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.attachment(filename);
    res.send(new Buffer(wbout, 'binary'));
  }
}

export class xslxtools {

  public xlsx_tools() {

  }

  public static workbook() {
    return new workbook;
  }

  public static worksheet(name):worksheet {
    return new worksheet(name);
  }

  public static init(m):xslxtools {
    return new xslxtools();
  }
}
