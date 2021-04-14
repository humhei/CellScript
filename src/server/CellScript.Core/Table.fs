namespace CellScript.Core
open Deedle
open CellScript.Core.Extensions
open Akka.Util
open System
open Shrimp.FSharp.Plus
open System.IO
open OfficeOpenXml
open Newtonsoft.Json


type Table = private Table of ExcelFrame<stringIC>
with 
    member internal x.AsExcelFrame = 
        let (Table frame) = x
        frame

    member x.AsFrame = 
        let (Table frame) = x
        frame.AsFrame


    member table.MapFrame(mapping) =
        ExcelFrame.mapFrame mapping table.AsExcelFrame
        |> Table

    member x.Headers =
        x.AsFrame.ColumnKeys

    member x.ToArray2D() =
        ExcelFrame.toArray2DWithHeader x.AsExcelFrame

    member x.SaveToXlsx(path: string, ?isOverride: bool, ?sheetName: string, ?columnAutofitOptions, ?tableName: string, ?tableStyle, ?customCellMapping: obj -> obj) =
        let isOverride = defaultArg isOverride true

        let sheetName = defaultArg sheetName "Sheet1"

        if isOverride
        then File.Delete path


        let excelPackage = new ExcelPackage(FileInfo(path))

        let worksheet = 
            excelPackage.Workbook.Worksheets.Add(sheetName)
            |> VisibleExcelWorksheet.Create

        let array2D =
            let array2D =
                x.ToArray2D()

            match customCellMapping with 
            | Some mapping ->
                array2D
                |> Array2D.map mapping

            | None -> array2D

        worksheet.LoadFromArraysAsTable(array2D, ?columnAutofitOptions = columnAutofitOptions, ?tableName = tableName, ?tableStyle = tableStyle)
        
        excelPackage.Save()
        excelPackage.Dispose()



    member x.GetColumns (indexes: seq<stringIC>)  =

        let mapping =
            Frame.filterCols (fun key _ ->
                Seq.exists (fun index -> key = index) indexes
            )

        x.MapFrame mapping 



    static member OfArray2D array2D =
        let frame = ExcelFrame.ofArray2DWithHeader array2D
        frame
        |> ExcelFrame.mapFrame(Frame.mapColKeys stringIC)
        |> Table.Table

    static member OfRecords records =
        ExcelFrame.ofRecords records
        |> ExcelFrame.mapFrame(Frame.mapColKeys stringIC)
        |> Table.Table

    static member OfFrame (frame: Frame<int, string>) =
        frame
        |> ExcelFrame.ofFrameWithHeaders 
        |> ExcelFrame.mapFrame(Frame.mapColKeys stringIC)
        |> Table.Table

    static member OfFrame (frame: Frame<int, stringIC>) =
        frame
        |> Frame.mapColKeys stringIC.value
        |> Table.OfFrame

    static member OfXlsxFile(xlsxFile: XlsxFile, ?rangeGettingOptions, ?sheetGettingOptions) =
        let rangeGettingOptions = 
            defaultArg rangeGettingOptions RangeGettingOptions.UserRange

        let sheetGettingOptions =
            defaultArg sheetGettingOptions SheetGettingOptions.DefaultValue

        let excelRangeInfo = ExcelRangeContactInfo.readFromFile rangeGettingOptions sheetGettingOptions xlsxFile
        
        Table.OfArray2D(excelRangeInfo.Content)

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
            Table.OfArray2D array2D :> ISurrogated


[<RequireQualifiedAccess>]
module Table =
    let mapFrame mapping (table: Table) =
        table.MapFrame mapping

    let fillEmptyUp (table: Table) =
        table 
        |> mapFrame Frame.fillEmptyUp

    let splitRowToMany addtionalHeaders mapping (table: Table) =
        let addtionalHeaders =
            addtionalHeaders
            |> List.map stringIC

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

    let getColumns (indexes) (table: Table) =
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

    let concat_RemoveRowKeys (tables: AtLeastOneList<Table>) =
        tables.Head
        |> mapFrame(fun _ ->
            tables
            |> AtLeastOneList.map (fun table -> table.AsFrame)
            |> Frame.Concat_RemoveRowKeys
        )



