namespace CellScript.Core
open Deedle
open CellScript.Core.Extensions
open Shrimp.FSharp.Plus
open OfficeOpenXml

type SearchableTable = SearchableTable of Frame<StringIC, StringIC>
with 
    static member OfCsvFile(csvFile) =
        let frame = Table.OfCsvFile(csvFile).AsFrame

        let headers =
            frame.ColumnKeys
            |> List.ofSeq
        frame
        |> Frame.indexRows headers.[0]
        |> Frame.mapRowKeys (string >> StringIC)
        |> SearchableTable


    member x.AsFrame =
        let (SearchableTable v) = x
        v

    member x.Item(rowKey, columnKey) =
        x.AsFrame.[StringIC columnKey, StringIC rowKey]
        |> string
        
          
