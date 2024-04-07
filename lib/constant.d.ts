export interface IColsGroup {
    content: string;
    origin: string;
    colStart: number;
    colEnd: number;
    rowStart: number;
    rowEnd: number;
}
export interface IStyle {
    rowStart: number;
    rowEnd?: number;
    colStart: number;
    colEnd?: number;
    type?: "b" | "e" | "n" | "d" | "s" | "z";
    format?: string;
    style: any;
}
export interface IConvertDataToExcel {
    data: Object[];
    fileExtension?: ".xlsx" | ".xls";
    startCellOfData?: string;
    colsGroup?: IColsGroup[];
    sheetName?: string;
    colWidths?: number[];
    rowHeights?: number[];
    fileName?: string;
    isBordered?: boolean;
    styles?: IStyle[];
}
