namespace CellScript.Core
open Deedle
open CellScript.Core.Extensions
open Akka.Util
open System
open Shrimp.FSharp.Plus

/// Ignore case
[<CustomEquality; CustomComparison>]
type TableHeader = TableHeader of string
with 
    member x.ValueLC =
        let (TableHeader v) = x
        v.ToLower()

    override x.Equals(yobj) =  
       match yobj with 
       | :? TableHeader as y -> 
            x.ValueLC = y.ValueLC
       | _ -> false 
    override x.GetHashCode() = x.LowerValue.GetHashCode()

    interface System.IComparable with 
       member x.CompareTo yobj =  

         match yobj with 
         | :? TableHeader as y -> compare x.LowerValue y.LowerValue
         | _ -> invalidArg "yobj" "cannot compare value of different types" 


type Table = private Table of ExcelFrame<int, TableHeader>
with 
    member internal x.AsExcelFrame = 
        let (Table frame) = x
        frame

    member x.AsFrame = 
        let (Table frame) = x
        frame.AsFrame

    member private table.MapFrame(mapping) =
        ExcelFrame.mapFrame mapping table.AsExcelFrame
        |> Table

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

    member x.GetColumns (indexes: seq<string>)  =
        let mapping =
            Frame.filterCols (fun key _ ->
                Seq.exists (fun m -> String.Compare(m, key, true) = 0) indexes
            )

        x.MapFrame mapping 

    interface IToArray2D with 
        member x.ToArray2D() =
            ExcelFrame.toArray2DWithHeader x.AsExcelFrame

    interface ISurrogated with 
        member x.ToSurrogate(system) = 
            TableSurrogate (x.ToArray2D()) :> ISurrogate

    static member OfArray2D array2D =
        let frame = ExcelFrame.ofArray2DWithHeader array2D
        Table frame

    static member OfRecords records =
        ExcelFrame.ofRecords records
        |> Table



and private TableSurrogate = TableSurrogate of obj[,]
with 
    interface ISurrogate with 
        member x.FromSurrogate(system) = 
            let (TableSurrogate array2D) = x
            Table.OfArray2D array2D :> ISurrogated


[<RequireQualifiedAccess>]
module Table =
    let mapFrame mapping (table: Table) =
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

    let fillEmptyUp (table: Table) =
        table 
        |> mapFrame Frame.fillEmptyUp

    let splitRowToMany addtionalHeaders mapping (table: Table) =
        table 
        |> mapFrame (Frame.splitRowToMany addtionalHeaders mapping)

    let removeEmptyRows (table: Table) =
        table 
        |> mapFrame(
            Frame.filterRowValues(fun row ->
                row.GetAllValues()
                |> Seq.exists (fun v ->
                    match OptionalValue.asOption v with 
                    | CellValue.HasSense _ -> true
                    | _ -> false
                )
            )
        )

    let getColumns (indexes: seq<string>) (table: Table) =
        table.GetColumns(indexes)


    let toExcelArray (Table excelFrame) =
        excelFrame
        |> ExcelFrame.mapFrame (fun frame ->
            let sequenceHeaders = 
                frame.ColumnKeys
                |> Seq.mapi (fun i _ -> i)
            Frame.indexColsWith sequenceHeaders frame
        )
        |> ExcelArray

    let concat (tables: AtLeastOneList<Table>) =

        let headerListsLower =
            tables.AsList
            |> List.map (fun m ->
                m.GetHeadersLC()
                |> Set.ofSeq
            )

        headerListsLower
        |> List.reduce(fun headers1 headers2 -> 
            if headers1 <> headers2 then failwithf "headers1 %A <> headers2 %A when concating table" headers1 headers2
            else headers2
        )
        |> ignore

        let rows =
            tables.AsList
            |> List.collect (fun table ->
                table.AsFrame.Rows.Values
                |> List.ofSeq
                |> List.map (fun row -> row)
            )

        Frame.ofRowsOrdinal rows
        |> Frame.mapRowKeys int