// Learn more about F# at http://fsharp.org

open System
open CellScript.Core

[<EntryPoint>]
let main argv =
    //let m = ExcelRangeContactInfo.readFromFile (UserRange) (SheetNameOrSheetIndex("Sheet1", 0)) "datas/emptySheet.xlsx"
    let m = ExcelRangeContactInfo.readFromFile (UserRange) (SheetName("Sheet1")) "datas/EU80-20255~20264 装箱明细.xlsx"
    let table = Table.OfArray2D m.Content
    printfn "Hello World from F#!"
    0 // return an integer exit code
