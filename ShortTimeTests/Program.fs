// Learn more about F# at http://fsharp.org

open System
open CellScript.Core
open CellScript.Core.Extensions
open Deedle
open Shrimp.FSharp.Plus
open System.Collections.Generic
open System.IO
open Hyperion
open Akka.Util
open Akkling
open Newtonsoft.Json
open OfficeOpenXml
open Shrimp.FSharp.Plus.Operators




type KK =
    | A
    | B 

type M =
    { 版面: int 
      订单数量: int 
      拼版个数: int 
      纸张数量: int }
with 
    member x.XXNumber = 600
[<EntryPoint>]
let main argv =
    let tb = Table.OfXlsxFile (XlsxFile @"D:\VsCode\Workspace\CellScript\ShortTimeTests\datas\testData.xlsx")

    let p = Table.fillEmptyUp tb
    let kb = p.ToArray2D()
    let rows: list<Series<_, _>> = []

    let m = 
        rows
        |> Table.OfRowsOrdinal

    //let excelPackage = new ExcelPackage(FileInfo @"D:\Users\Jia\Documents\MyData\Docs\2017\健耐\嘴唇\数据库.xlsx")
    //let sheet = excelPackage.GetValidWorksheet(SheetGettingOptions.SheetName (StringIC"shEEt1")).Value.Cells.[1, 1, 1, 10]
    眼镜复杂外箱贴.Module_眼镜复杂外箱贴.storages()


    printfn "Hello World from F#!"
    0 // return an integer exit code
