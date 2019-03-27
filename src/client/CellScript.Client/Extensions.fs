module CellScript.Client.Extensions
open ExcelDna.Integration
open System.IO
open CellScript.Core.Extensions
open ExcelProcess


[<RequireQualifiedAccess>]
module ExcelRangeBase =
    let mapByXlRef (xlRef: ExcelReference) mapping =
        let retrievedSheetName = XlCall.Excel(XlCall.xlSheetNm,xlRef) :?> string

        let sheetPath =
            let dir = XlCall.Excel(XlCall.xlfGetDocument,2,retrievedSheetName) :?> string
            let name = XlCall.Excel(XlCall.xlfGetDocument,88,retrievedSheetName) :?> string
            Path.Combine(dir,name)
                
        let sheetName =
            String.rightOf "]" retrievedSheetName

        use p = Excel.getWorksheetByName sheetName sheetPath

        p.Cells.[xlRef.RowFirst + 1,xlRef.ColumnFirst + 1,xlRef.RowLast + 1,xlRef.ColumnLast + 1]
        |> mapping