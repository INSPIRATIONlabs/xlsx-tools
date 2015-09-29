/// <reference path="./typings/node/node.d.ts"/>
/// <reference path="./typings/lodash/lodash.d.ts"/>
/// <reference path="./typings/xlsx/xlsx.d.ts"/>

import _ = require('lodash');
import XLSX = require('xlsx');

class worksheet {
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

  public datenum(v, date1904) {
  	if(date1904) v+=1462;
  	var epoch = Date.parse(v);
  	return (epoch - new Date(Date.UTC(1899, 11, 30))) / (24 * 60 * 60 * 1000);
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
    if(this.headerColumns) {
      _.each(this.headerColumns, (col) => {
        var item = row[col];
        var cell;
        if(typeof item === 'object' && item.v !== "undefined") {
          cell = item;
        } else {
          cell = {v: item};
        }
        this.setCell(this.R, this.C, cell);
        this.C++;
      });
      this.R++;
      this.C = 0;
    }
  }

  protected setCell(R, C, cell) {
    var ws = {};
    if(this.range.s.r > R) this.range.s.r = R;
    if(this.range.s.c > C) this.range.s.c = C;
    if(this.range.e.r < R) this.range.e.r = R;
    if(this.range.e.c < C) this.range.e.c = C;

    if(cell.v != null) {
      var cell_ref = XLSX.utils.encode_cell({c:C, r:R});
      if(typeof cell.v === 'number') cell.t = 'n';
			else if(typeof cell.v === 'boolean') cell.t = 'b';
      else if(cell.v instanceof Date) {
				cell.t = 'n';
        if(cell.z == null) {
          cell.z = XLSX.SSF._table[14];
        }
				cell.v = this.datenum(cell.v);
			} else {
        cell.t = 's';
      }
      this.data[cell_ref] = cell;
    }
  }

  protected write() {
    if(this.range.s.c < 10000000) {
      this.data['!ref'] = XLSX.utils.encode_range(this.range);
    }
  }
}

class workbook {
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

  public download(res, filename) {
    var wbout = this.write();
    res.type('application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.attachment(filename);
    res.send(new Buffer(wbout, 'binary'));
  }
}

class xslx_tools {

  public xlsx_tools() {

  }

  public static workbook() {
    return new workbook;
  }

  public static worksheet(name):worksheet {
    return new worksheet(name);
  }

  public static init(m):xslx_tools {
    return new xslx_tools();
  }

}

export = xslx_tools;
