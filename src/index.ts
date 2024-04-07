import * as XLSX from "xlsx";
import * as XLSXStyle from "xlsx-js-style";
import * as FileSaver from "file-saver";

import { IColsGroup, IConvertDataToExcel, IStyle } from "./constant";

export default function convertDataToExcel({
  data,
  fileExtension = ".xlsx",
  startCellOfData = "A1",
  colsGroup,
  sheetName = "Sheet 1",
  colWidths,
  rowHeights,
  fileName = "example",
  isBordered,
  styles,
}: IConvertDataToExcel) {
  const fileType =
    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;charset=UTF-8";

  const options: any = {
    origin: startCellOfData,
  };

  const ws = XLSX.utils.json_to_sheet(data, options);

  if (colsGroup?.length) {
    colsGroup?.forEach((colGroup: IColsGroup) => {
      ws["!merges"] = [
        ...(ws["!merges"] || []),
        {
          s: { c: colGroup.colStart, r: colGroup.rowStart },
          e: { c: colGroup.colEnd, r: colGroup.rowEnd },
        },
      ];
      XLSX.utils.sheet_add_aoa(ws, [[colGroup.content]], {
        origin: colGroup.origin,
      });
    });
  }

  if (styles?.length) {
    for (var i in ws) {
      const cell = XLSX.utils.decode_cell(i);

      styles.forEach((style: IStyle) => {
        let checkCellAvaiable = true;
        if (!style.rowEnd) {
          if (cell.r < style.rowStart) {
            checkCellAvaiable = false;
          }
        } else {
          if (cell.r < style.rowStart || cell.r > style.rowEnd) {
            checkCellAvaiable = false;
          }
        }

        if (!style.colEnd) {
          if (cell.c < style.colStart) {
            checkCellAvaiable = false;
          }
        } else {
          if (cell.c < style.colStart || cell.c > style.colEnd) {
            checkCellAvaiable = false;
          }
        }

        if (checkCellAvaiable) {
          ws[i].s = style.style;
          ws[i].t = style.type || "s";
          ws[i].z = style.format;
        }
      });
    }
  }

  if (isBordered) {
    for (var i in ws) {
      const cell = XLSX.utils.decode_cell(i);
      const rowOfHeader = Number(startCellOfData.split("").pop());
      if (cell.r >= rowOfHeader - 1) {
        ws[i].s = {
          ...ws[i].s,
          border: {
            right: { style: "thin" },
            left: { style: "thin" },
            top: { style: "thin" },
            bottom: { style: "thin" },
          },
        };
      }
    }
  }

  if (colWidths?.length) {
    const wscols = colWidths.map((colWidth) => ({
      wch: colWidth,
    }));
    ws["!cols"] = wscols;
  }

  if (rowHeights?.length) {
    const wsrows = rowHeights.map((rowHeight) => ({
      hpt: rowHeight,
    }));
    ws["!rows"] = wsrows;
  }

  const wb = { Sheets: { [sheetName]: ws }, SheetNames: [sheetName] };
  const excelBuffer = XLSXStyle.write(wb, {
    bookType: "xlsx",
    type: "array",
  });

  const dataConverted = new Blob([excelBuffer], { type: fileType });
  FileSaver.saveAs(dataConverted, fileName + fileExtension);
  return dataConverted;
}
