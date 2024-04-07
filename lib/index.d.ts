import { IConvertDataToExcel } from "./constant";
export default function convertDataToExcel({ data, fileExtension, startCellOfData, colsGroup, sheetName, colWidths, rowHeights, fileName, isBordered, styles, }: IConvertDataToExcel): Blob;
