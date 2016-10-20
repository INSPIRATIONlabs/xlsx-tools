"use strict";
const _ = require('lodash');
const XLSX = require('xlsx');
class worksheet {
    constructor(name) {
        this.headerColumns = [];
        this.rows = [];
        this.C = 0;
        this.R = 0;
        this.range = { s: { c: 10000000, r: 10000000 }, e: { c: 0, r: 0 } };
        this.data = {};
        this.name = name;
    }
    datenum(v, date1904) {
        if (date1904)
            v += 1462;
        var epoch = Date.parse(v);
        let calcdate = new Date(Date.UTC(1899, 11, 30));
        let returnval = (epoch - calcdate) / (24 * 60 * 60 * 1000);
        return returnval;
    }
    setHeader(arr) {
        _.each(arr, (item) => {
            this.headerColumns.push(item);
        });
        _.each(this.headerColumns, (col) => {
            this.setCell(this.R, this.C, { v: col });
            this.C++;
        });
        this.R++;
        this.C = 0;
    }
    addRow(row) {
        if (this.headerColumns.length) {
            _.each(this.headerColumns, (col) => {
                var cell = row[col];
                this.setCell(this.R, this.C, cell);
                this.C++;
            });
        }
        else {
            _.each(row, (cell) => {
                this.setCell(this.R, this.C, cell);
                this.C++;
            });
        }
        this.R++;
        this.C = 0;
    }
    encodeCell(obj) {
        return XLSX.utils.encode_cell(obj);
    }
    setCell(R, C, cell) {
        var ws = {};
        if (this.range.s.r > R)
            this.range.s.r = R;
        if (this.range.s.c > C)
            this.range.s.c = C;
        if (this.range.e.r < R)
            this.range.e.r = R;
        if (this.range.e.c < C)
            this.range.e.c = C;
        var cell_ref = this.encodeCell({ c: C, r: R });
        if (!_.isObject(cell)) {
            cell = { v: cell };
        }
        if (cell.v != null) {
            if (cell.v instanceof Date) {
                cell.t = 'n';
                if (cell.z == null) {
                    cell.z = XLSX.SSF._table[14];
                }
                cell.v = this.datenum(cell.v);
            }
            else if (typeof cell.v === 'number') {
                cell.t = 'n';
            }
            else if (typeof cell.v === 'boolean') {
                cell.t = 'b';
            }
            else if (!cell.t) {
                cell.t = 's';
            }
            this.data[cell_ref] = cell;
        }
        else if (cell.f != null) {
            cell.t = 'f';
            this.data[cell_ref] = cell;
        }
    }
    write() {
        if (this.range.s.c < 10000000) {
            this.data['!ref'] = XLSX.utils.encode_range(this.range);
        }
    }
}
exports.worksheet = worksheet;
class workbook {
    constructor() {
        this.worksheets = [];
    }
    addSheet(worksheet) {
        this.worksheets.push(worksheet);
    }
    write() {
        var wb = {
            SheetNames: [],
            Sheets: {}
        };
        _.each(this.worksheets, (sheet) => {
            sheet.write();
            wb.SheetNames.push(sheet.name);
            wb.Sheets[sheet.name] = sheet.data;
        });
        var wopts = { bookType: 'xlsx', bookSST: false, type: 'binary' };
        var output = XLSX.write(wb, wopts);
        return output;
    }
    download(res, filename) {
        var wbout = this.write();
        res.type('application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
        res.attachment(filename);
        res.send(new Buffer(wbout, 'binary'));
    }
}
exports.workbook = workbook;
class xslxtools {
    xlsx_tools() {
    }
    static workbook() {
        return new workbook;
    }
    static worksheet(name) {
        return new worksheet(name);
    }
    static init(m) {
        return new xslxtools();
    }
}
exports.xslxtools = xslxtools;
