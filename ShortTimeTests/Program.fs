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
    let template = 
        @"\\2021-pc\JobData\2023年-管家婆\模板\错误修复模板.xlsx"
        |> XlsxFile

    let values = 
        @"\\2021-pc\2023-已完成出CTP文件2\自动打印\资料\中达产品资料.xlsx"
        |> XlsxFile
        |> Table.OfXlsxFile 
        |> Table.mapFrame (Frame.filterCols(fun col _ ->
            match col.Value with 
            | "零售价" -> false
            | _ -> true
        ))

    let addressedArray = 
        { AddressedArray.ofTable "A1" (values) with 
            SpecificName = Some "产品资料"
        } 
        |> List.singleton
        |> AddressedArrays.OfList

    let named =
        { NamedAddressedArrays.SheetName = "产品资料修复"
          SpecificTargetSheetName = None 
          AddressedArrays = addressedArray }

    let targetXlsxPath =
        @"C:\Users\Administrator\Desktop\中达产品结果.xlsx"
        |> XlsxPath

    named.SaveToXlsx_FromTemplate(template, targetXlsxPath)


    0