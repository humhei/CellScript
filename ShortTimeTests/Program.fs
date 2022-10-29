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
open FParsec
open FParsec.CharParsers
open System.Collections.Concurrent

let private hyperionSerializer = Hyperion.Serializer()
let private serialize (v: obj) =
    use stream = new MemoryStream()
    hyperionSerializer.Serialize(v, stream)
    let bytes = stream.ToArray()
    bytes

let private deserialize (bytes: byte []) =
    use stream = new MemoryStream(bytes)

    hyperionSerializer.Deserialize(stream)
    |> unbox

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

    //let m = 
    //    StringIC "Hello"

    //let texts =
    //    let m =
    //        List.replicate 15 ["Yes"; "PPPP"; "HSASA"; "GoGo"; "Hello1" ]
    //        |> List.concat
    //    m @ ["Hello"]
    //    |> List.map StringIC

    //let cache = ConcurrentDictionary<StringIC, bool>()

    //let p =
    //    List.replicate  1000000 m
    //    |> List.map(fun m ->
    //        match texts with 
    //        | List.Contains m -> true
    //        | _ -> false
    //        //cache.GetOrAdd(m, fun _ ->
    //        //    match texts with 
    //        //    | List.Contains m -> true
    //        //    | _ -> false
    //        //)
    //    )


    //let parser =
    //    (pstringCI "Column" .>> pint32) 
    //    <|>
    //    (pstringCI "CellScriptColumn" .>> pint32) 


    //let m = run parser "CellScriptColumn1"


    let tb = Table.OfXlsxFile (
        XlsxFile @"D:\Users\Jia\Documents\MyData\Docs\2017\健耐\Treeker\Input\Orders\健耐 22TW30\Sheet1@健耐 22TW30.output.xlsx"
    )

    tb.SaveToXlsx(@"C:\Users\Administrator\Desktop\新建 Microsoft Excel Worksheet.xlsx", TableXlsxSavingOptions.Create(cellFormat = CellSavingFormat.ODBC))

    //let m = serialize tb
    //File.WriteAllBytes(@"C:\Users\Jia\Desktop\新建文本文档.txt", m)
    let m = File.ReadAllBytes(@"C:\Users\Jia\Desktop\新建文本文档.txt")
    let p1 = deserialize m
    let p2 = deserialize m
    let excelPackage = new ExcelPackage(new FileInfo(@"D:\Users\Jia\Documents\MyData\Docs\2017\健耐\Markham\数据库.xlsx"))
    let p = excelPackage.GetValidWorksheet(SheetGettingOptions.DefaultValue)

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
