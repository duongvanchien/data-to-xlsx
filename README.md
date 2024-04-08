## About
The JSON to Excel data conversion library offers an easy and efficient way to transform data structures from JSON format into Excel spreadsheets. With flexibility and high accuracy, this tool saves time and effort in processing and presenting information. Particularly useful for those working in data analysis and data science fields, it serves as a valuable asset in managing and analyzing data effectively.

## Installation
Install [npm](https://nodejs.org/)
```sh
npm i data-to-xlsx
```

Additional typescript definitions
```sh
npm i @types/data-to-xlsx --save-dev
```

## Demo
A minimal demo page can be found in `src` directory.
[Online demo](https://codesandbox.io/p/sandbox/data-to-xlsx-example-xjnwlx?layout=%257B%2522sidebarPanel%2522%253A%2522EXPLORER%2522%252C%2522rootPanelGroup%2522%253A%257B%2522direction%2522%253A%2522horizontal%2522%252C%2522contentType%2522%253A%2522UNKNOWN%2522%252C%2522type%2522%253A%2522PANEL_GROUP%2522%252C%2522id%2522%253A%2522ROOT_LAYOUT%2522%252C%2522panels%2522%253A%255B%257B%2522type%2522%253A%2522PANEL_GROUP%2522%252C%2522contentType%2522%253A%2522UNKNOWN%2522%252C%2522direction%2522%253A%2522vertical%2522%252C%2522id%2522%253A%2522cluqgbmb60006356kw3x3xq78%2522%252C%2522sizes%2522%253A%255B100%252C0%255D%252C%2522panels%2522%253A%255B%257B%2522type%2522%253A%2522PANEL_GROUP%2522%252C%2522contentType%2522%253A%2522EDITOR%2522%252C%2522direction%2522%253A%2522horizontal%2522%252C%2522id%2522%253A%2522EDITOR%2522%252C%2522panels%2522%253A%255B%257B%2522type%2522%253A%2522PANEL%2522%252C%2522contentType%2522%253A%2522EDITOR%2522%252C%2522id%2522%253A%2522cluqgbmb60002356kct6smxxt%2522%257D%255D%257D%252C%257B%2522type%2522%253A%2522PANEL_GROUP%2522%252C%2522contentType%2522%253A%2522SHELLS%2522%252C%2522direction%2522%253A%2522horizontal%2522%252C%2522id%2522%253A%2522SHELLS%2522%252C%2522panels%2522%253A%255B%257B%2522type%2522%253A%2522PANEL%2522%252C%2522contentType%2522%253A%2522SHELLS%2522%252C%2522id%2522%253A%2522cluqgbmb60003356kw6w8sg1k%2522%257D%255D%252C%2522sizes%2522%253A%255B100%255D%257D%255D%257D%252C%257B%2522type%2522%253A%2522PANEL_GROUP%2522%252C%2522contentType%2522%253A%2522DEVTOOLS%2522%252C%2522direction%2522%253A%2522vertical%2522%252C%2522id%2522%253A%2522DEVTOOLS%2522%252C%2522panels%2522%253A%255B%257B%2522type%2522%253A%2522PANEL%2522%252C%2522contentType%2522%253A%2522DEVTOOLS%2522%252C%2522id%2522%253A%2522cluqgbmb60005356kea3dr2vp%2522%257D%255D%252C%2522sizes%2522%253A%255B100%255D%257D%255D%252C%2522sizes%2522%253A%255B50%252C50%255D%257D%252C%2522tabbedPanels%2522%253A%257B%2522cluqgbmb60002356kct6smxxt%2522%253A%257B%2522id%2522%253A%2522cluqgbmb60002356kct6smxxt%2522%252C%2522tabs%2522%253A%255B%255D%257D%252C%2522cluqgbmb60005356kea3dr2vp%2522%253A%257B%2522tabs%2522%253A%255B%257B%2522id%2522%253A%2522cluqgbmb60004356k3cu9jdzy%2522%252C%2522mode%2522%253A%2522permanent%2522%252C%2522type%2522%253A%2522UNASSIGNED_PORT%2522%252C%2522port%2522%253A0%252C%2522path%2522%253A%2522%252F%2522%257D%255D%252C%2522id%2522%253A%2522cluqgbmb60005356kea3dr2vp%2522%252C%2522activeTabId%2522%253A%2522cluqgbmb60004356k3cu9jdzy%2522%257D%252C%2522cluqgbmb60003356kw6w8sg1k%2522%253A%257B%2522tabs%2522%253A%255B%255D%252C%2522id%2522%253A%2522cluqgbmb60003356kw6w8sg1k%2522%257D%257D%252C%2522showDevtools%2522%253Atrue%252C%2522showShells%2522%253Afalse%252C%2522showSidebar%2522%253Atrue%252C%2522sidebarPanelSize%2522%253A15%257D) is also available!

## Import
------
```js
import convertDataToExcel from 'data-to-xlsx';
```

```js
const convertDataToExcel = require("data-to-xlsx")
```

## Examples
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

## API

| Props        |  Type | Require | Default value | Description
| -------------- | ------------- | ------------ | ------------ | ------------
| data    | `Object[]`          | Yes | | Values of table in file excel
| fileExtension   | `.xlsx`, `.xls`    | No | `.xlsx` | The file type will be downloaded after conversion          
| startCellOfData   | `string`    | No | `A1` | The starting cell name of the data when saved in the excel file
| colsGroup   | `IColsGroup[]`    | No | | Use to group specific cells in a file
| sheetName   | `string`    | No | `Sheet 1` | Name of the sheet file
| colWidths   | `number[]`    | No | | The size of the columns in the file starts from the first column (A)
| rowHeights   | `number[]`    | No | | The size of the rows in the file starts from the first row (1)
| fileName   | `string`    | No | `example` | The name of the file after it is downloaded
| isBordered   | `boolean`    | No | `false` | Determines whether to fill the data cells with borders
| styles   | `IStyle[]`    | No |  | Style for cells in the file

In the data of `data` elements, the keys in the object are used as the header of the data table

## IColsGroup
| Props        |  Type | Require | Default value | Description
| -------------- | ------------- | ------------ | ------------ | ------------
| content    | `string`          | Yes | | Content of cell group
| origin   | `string`    | Yes |  | The starting cell name of the cell group          
| colStart   | `number`    | Yes |  | Column numbers start from 0
| colEnd   | `number`    | Yes |  | Column number end
| rowStart   | `number`    | Yes |  | Row numbers start from 0
| rowEnd   | `number`    | Yes |  | Row numbers end from 0

In case you only want to group in the same row or same column, you can configure the parameters `colStart` = `colEnd`, `rowStart` = `rowEnd`

## IStyle
| Props        |  Type | Require | Default value | Description
| -------------- | ------------- | ------------ | ------------ | ------------
| rowStart    | `number`          | Yes | | Row numbers start from 0
| rowEnd   | `number`    | Yes |  | Row numbers end       
| colStart   | `number`    | Yes |  | Column numbers start from 0
| colEnd   | `number`    | Yes |  | Column number end
| type   | `b`, `e`, `n`, `d`, `s`, `z`   | No | `s` | Cell value type
| format   | `number`    | No |  | Format cell values
| style   | [`Cell Style Properties`](https://www.npmjs.com/package/xlsx-js-style#cell-style-properties)    | No |  | Configure the style of the cell
