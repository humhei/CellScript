namespace CellScript.Core
open Deedle
open Registration
open CellScript.Core.Extensions
open Types

[<CustomParamConversion>]
type Table = Table of ExcelFrame<int,string>
with
    member x.Headers =
        let (Table excelFrame) = x
        excelFrame.AsFrame.ColumnKeys

    interface ICustomReturn with
        member x.ReturnValue() =
            let (Table excelFrame) = x
            ExcelFrame.toArray2DWithHeader excelFrame

    static member Convert (input: obj[,]) =
        ExcelFrame.ofArray2DWithHeader input 
        |> Table

[<RequireQualifiedAccess>]
module Table =
    let value (Table table) = table

    let mapFrame mapping (Table table) =
        ExcelFrame.map mapping table
        |> Table

    let item (indexes: seq<string>) table =
        let mapping =
            Frame.filterCols (fun key _ ->
                Seq.exists (String.equalIgnoreCaseAndEdgeSpace key) indexes
            )

        mapFrame mapping table
