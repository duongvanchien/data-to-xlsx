# About
The JSON to Excel data conversion library offers an easy and efficient way to transform data structures from JSON format into Excel spreadsheets. With flexibility and high accuracy, this tool saves time and effort in processing and presenting information. Particularly useful for those working in data analysis and data science fields, it serves as a valuable asset in managing and analyzing data effectively.

# Installation
Install [npm](https://nodejs.org/)
```sh
npm i data-to-xlsx
```

Additional typescript definitions
```sh
npm i @types/data-to-xlsx --save-dev
```

Import
------
```js
import convertDataToExcel from 'data-to-xlsx';
```

```js
const convertDataToExcel = require("data-to-xlsx")
```

Examples
--------

### Using convertDataToExcel in ReactJs

```js
import convertDataToExcel from 'data-to-xlsx';

const App = () => {
    const onDownload = () => {
        convertDataToExcel({
            data: [
              {
                Index: 1,
                Name: "Jindo Katory",
                Age: 18
              },
              {
                Index: 2,
                Name: "Katory Jindo",
                Age: 17
              },
            ],
            fileName: "example"
        })
    }

    return(
        <button onClick={onDownload}>Download</button>
    )
}

export default App;
```

# API

| Props        |  Type | Require | Default value | Description
| -------------- | ------------- | ------------ | ------------ | ------------
| data    | `Object[]`          | Yes | | Values of table in file excel
| fileExtension   | `.xlsx`, `.xls`    | No | `.xlsx` | The file type will be downloaded after conversion          
| startCellOfData   | `string`    | No | `A1` | The starting cell name of the data when saved in the excel file
| colsGroup   | `IColsGroup`    | No | | Use to group specific cells in a file
| sheetName   | `string`    | No | `Sheet 1` | Name of the sheet file
| colWidths   | `number[]`    | No | | The size of the columns in the file starts from the first column (A)
| rowHeights   | `number[]`    | No | | The size of the rows in the file starts from the first row (1)
| fileName   | `string`    | No | `example` | The name of the file after it is downloaded
| isBordered   | `boolean`    | No | `false` | Determines whether to fill the data cells with borders
| styles   | `IStyle`    | No |  | Style for cells in the file

In the data of `data` elements, the keys in the object are used as the header of the data table

# IColsGroup
| Props        |  Type | Require | Default value | Description
| -------------- | ------------- | ------------ | ------------ | ------------
| content    | `string`          | Yes | | Content of cell group
| origin   | `string`    | Yes |  | The starting cell name of the cell group          
| colStart   | `number`    | Yes |  | Column numbers start from 0
| colEnd   | `number`    | Yes |  | Column number end
| rowStart   | `number`    | Yes |  | Row numbers start from 0
| rowEnd   | `number`    | Yes |  | Row numbers end from 0

In case you only want to group in the same row or same column, you can configure the parameters `colStart` = `colEnd`, `rowStart` = `rowEnd`

# IStyle
| Props        |  Type | Require | Default value | Description
| -------------- | ------------- | ------------ | ------------ | ------------
| rowStart    | `number`          | Yes | | Row numbers start from 0
| rowEnd   | `number`    | Yes |  | Row numbers end       
| colStart   | `number`    | Yes |  | Column numbers start from 0
| colEnd   | `number`    | Yes |  | Column number end
| type   | `b`, `e`, `n`, `d`, `s`, `z`   | No | `s` | Cell value type
| format   | `number`    | No |  | Format cell values
| style   | [`Cell Style Properties`](https://www.npmjs.com/package/xlsx-js-style#cell-style-properties)    | No |  | Configure the style of the cell
