namespace CellScript.Core
open Deedle
open CellScript.Core.Extensions
open Akka.Util
open System

type Table = Table of ExcelFrame<int, string>
with 
    member x.AsExcelFrame = 
        let (Table frame) = x
        frame

    member x.AsFrame = 
        let (Table frame) = x
        frame.AsFrame

    member private table.MapFrame(mapping) =
        ExcelFrame.mapFrame mapping table.AsExcelFrame
        |> Table

    static member Convert array2D =
        let frame = ExcelFrame.ofArray2DWithHeader array2D
        Table frame

    member x.Headers =
        x.AsFrame.ColumnKeys

    member x.GetHeadersLC() =
        x.AsFrame.ColumnKeys
        |> Seq.map (fun m -> m.ToLower())

    member table.ConvertHeadersToLC() =
        table.MapFrame (fun frame ->
            let originColumnKeys = frame.ColumnKeys |> List.ofSeq
            let lowerColumnkeys = originColumnKeys |> List.map (fun m -> m.ToLower())
            Frame.indexColsWith lowerColumnkeys frame
        )


    member x.ToArray2D() =
        ExcelFrame.toArray2DWithHeader x.AsExcelFrame


    interface IToArray2D with 
        member x.ToArray2D() =
            ExcelFrame.toArray2DWithHeader x.AsExcelFrame

    interface ISurrogated with 
        member x.ToSurrogate(system) = 
            TableSurrogate (x.ToArray2D()) :> ISurrogate

and private TableSurrogate = TableSurrogate of obj[,]
with 
    interface ISurrogate with 
        member x.FromSurrogate(system) = 
            let (TableSurrogate array2D) = x
            Table.Convert array2D :> ISurrogated


[<RequireQualifiedAccess>]
module Table =

    let private mapFrame mapping (table: Table) =
        ExcelFrame.mapFrame mapping table.AsExcelFrame
        |> Table

    let convertHeadersToLC (table: Table) =
        table.ConvertHeadersToLC()
        

    /// mapping frame using lower case headers
    let mapFrameUsingLC mapping table =
        mapFrame (fun frame ->
            let originColumnKeys = frame.ColumnKeys |> List.ofSeq
            let lowerColumnkeys = originColumnKeys |> List.map (fun m -> m.ToLower())
            Frame.indexColsWith lowerColumnkeys frame
            |> mapping
            |> Frame.mapColKeys(fun key ->
                originColumnKeys
                |> List.tryFind (fun m -> String.Compare(m, key, true) = 0)
                |> function
                    | Some key -> key
                    | None -> key
            )
        ) table

    let getColumns (indexes: seq<string>) table =
        let mapping =
            Frame.filterCols (fun key _ ->
                Seq.exists (String.equalIgnoreCase key) indexes
            )

        mapFrame mapping table

    let asExcelArray (Table excelFrame) =
        excelFrame
        |> ExcelFrame.mapFrame (fun frame ->
            let sequenceHeaders = 
                frame.ColumnKeys
                |> Seq.mapi (fun i _ -> i)
            Frame.indexColsWith sequenceHeaders frame
        )
        |> ExcelArray