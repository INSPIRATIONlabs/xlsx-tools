export declare class worksheet {
    protected name: string;
    protected headerColumns: any[];
    protected rows: any[];
    protected C: number;
    protected R: number;
    protected range: {
        s: {
            c: number;
            r: number;
        };
        e: {
            c: number;
            r: number;
        };
    };
    protected data: {};
    constructor(name: any);
    datenum(v: any, date1904?: any): any;
    setHeader(arr: any): void;
    addRow(row: any): void;
    encodeCell(obj: any): any;
    protected setCell(R: any, C: any, cell: any): void;
    protected write(): void;
}
export declare class workbook {
    protected worksheets: any[];
    addSheet(worksheet: any): void;
    write(): any;
    download(res: any, filename?: any): void;
}
export declare class xslxtools {
    xlsx_tools(): void;
    static workbook(): workbook;
    static worksheet(name: any): worksheet;
    static init(m: any): xslxtools;
}
