import convertDataToExcel from "./src/index";

const express = require("express");
const app = express();
const port = 3000;

app.get("/", (req: any, res: any) => {
  const data = convertDataToExcel({
    data: [
      {
        STT: 1,
        Tên: "Chiến",
        Tuổi: 18,
      },
      {
        STT: 1,
        Tên: "Công",
        Tuổi: 16,
      },
    ],
    fileExtension: ".xlsx",
    startCellOfData: "A5",
    colsGroup: [
      {
        content: "BÁO CÁO TEST",
        origin: "A1",
        colStart: 0,
        colEnd: 6,
        rowStart: 0,
        rowEnd: 1,
      },
      {
        content: "BÁO CÁO TEST 1",
        origin: "A3",
        colStart: 0,
        colEnd: 3,
        rowStart: 2,
        rowEnd: 2,
      },
    ],
    colWidths: [20, 30, 40, 50],
    rowHeights: [10, 20, 30, 40, 50, 60],
    sheetName: "Đây là sheet để test",
    isBordered: true,
    styles: [
      {
        rowStart: 4,
        colStart: 0,
        style: {
          alignment: {
            vertical: "center",
            horizontal: "center",
          },
        },
      },
    ],
  });

  res.type(data.type);
  data.arrayBuffer().then((buf) => {
    res.send(Buffer.from(buf));
  });
});

app.listen(port, () => {
  console.log(`App listening on port ${port}`);
});
