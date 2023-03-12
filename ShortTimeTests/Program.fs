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
    let m = Table.OfXlsxFile (XlsxFile @"D:\Users\Jia\Documents\MyData\Docs\2017\健耐\Shasa\数据库.xlsx")
    let m = SearchableTable.OfCsvFile (CsvFile  @"D:\Users\Jia\Documents\Shrimp.Workflow.TypeProvider\颜色翻译.csv")
    let p = m.["全黑", "Color"]

    0