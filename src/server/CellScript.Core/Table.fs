namespace CellScript.Core
open Deedle
open CellScript.Core.Extensions
open Akka.Util
open System
open Shrimp.FSharp.Plus

/// Ignore case
[<CustomEquality; CustomComparison>]
type KeyIC = KeyIC of string
with 

    member x.LowerValue =
        let (KeyIC v) = x
        v.ToLower()

    override x.ToString() =
        let (KeyIC v) = x
        v

    override x.Equals(yobj) =  
       match yobj with 
       | :? KeyIC as y -> 
            x.LowerValue = y.LowerValue
       | _ -> false 
    override x.GetHashCode() = x.LowerValue.GetHashCode()

    interface System.IComparable with 
       member x.CompareTo yobj =  

         match yobj with 
         | :? KeyIC as y -> compare x.LowerValue y.LowerValue
         | _ -> invalidArg "yobj" "cannot compare value of different types" 


type Table = private Table of ExcelFrame<int, KeyIC>
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

    member x.ToArray2D() =
        ExcelFrame.toArray2DWithHeader x.AsExcelFrame

    member x.GetColumns (indexes: seq<string>)  =

        let indexes =
            indexes
            |> Seq.map KeyIC

        let mapping =
            Frame.filterCols (fun key _ ->
                Seq.exists (fun index -> key = index) indexes
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
        frame
        |> ExcelFrame.mapFrame(Frame.mapColKeys KeyIC)
        |> Table

    static member OfRecords records =
        ExcelFrame.ofRecords records
        |> ExcelFrame.mapFrame(Frame.mapColKeys KeyIC)
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

    let fillEmptyUp (table: Table) =
        table 
        |> mapFrame Frame.fillEmptyUp

    let splitRowToMany addtionalHeaders mapping (table: Table) =

        let addtionalHeaders =
            addtionalHeaders
            |> List.map KeyIC

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

        let headerLists =
            tables.AsList
            |> List.map (fun m ->
                m.Headers
                |> Set.ofSeq
            )

        headerLists
        |> List.reduce(fun headers1 headers2 -> 
            if headers1 <> headers2 then failwithf "headers1 %A <> headers2 %A when concating table" headers1 headers2
            else headers2
        )
        |> ignore

        tables.Header
        |> mapFrame(fun _ ->
            let rows =
                tables.AsList
                |> List.collect (fun table ->
                    table.AsFrame.Rows.Values
                    |> List.ofSeq
                )

            Frame.ofRowsOrdinal rows
            |> Frame.mapRowKeys int
        )




