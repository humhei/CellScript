namespace CellScript.Core
open Deedle
open CellScript.Core.Extensions
open Types
open Akka.Util

type Table = Table of ExcelFrame<int, string>
with 
    member x.AsExcelFrame = 
        let (Table frame) = x
        frame

    member x.AsFrame = 
        let (Table frame) = x
        frame.AsFrame
                   
    member x.ToArray2D() =
        ExcelFrame.toArray2DWithHeader x.AsExcelFrame

    static member Convert array2D =
        let frame = ExcelFrame.ofArray2DWithHeader array2D
        Table frame

    static member TempData =
        let param =
            array2D [[box "Header1";box "Header2"]; [1; 2]]

        Table.Convert param


    interface ISurrogated with 
        member x.ToSurrogate(system) = 
            TableSurrogate (x.ToArray2D()) :> ISurrogate

and private TableSurrogate = TableSurrogate of obj [,]
with 
    interface ISurrogate with 
        member x.FromSurrogate(system) = 
            let (TableSurrogate array2D) = x
            Table.Convert array2D :> ISurrogated


[<RequireQualifiedAccess>]
module Table =

    let mapFrame mapping (table: Table) =
        mapping table.AsFrame
        |> ExcelFrame
        |> Table

    let item (indexes: seq<string>) table =
        let mapping =
            Frame.filterCols (fun key _ ->
                Seq.exists (String.equalIgnoreCaseAndEdgeSpace key) indexes
            )

        mapFrame mapping table
