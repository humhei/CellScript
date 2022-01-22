namespace CellScript.Core
open Deedle
open CellScript.Core.Extensions
open Akka.Util
open System
open Shrimp.FSharp.Plus
open System.IO
open OfficeOpenXml
open Newtonsoft.Json
open System.Runtime.CompilerServices
open FParsec
open FParsec.CharParsers
open CellScript.Core.Constrants

type ITableCell =
    abstract member CellText: unit -> string


type ITableCellValue =
    abstract member Value: unit -> IConvertible

[<RequireQualifiedAccess>]
module ITableCell =
    let getCellText (cell: #ITableCell) =
        cell.CellText()



[<Extension>]
type _ITableCellExtensions =
    [<Extension>] static member CellText(cell: #ITableCell) = cell.CellText()


type TableXlsxSavingOptions =
    { IsOverride: bool 
      SheetName: string
      ColumnAutofitOptions: ColumnAutofitOptions
      TableName: string
      AutoNumbericColumn: bool
      TableStyle: Table.TableStyles
      }
with 
    static member Create(?isOverride: bool, ?sheetName: string, ?columnAutofitOptions, ?tableName: string, ?tableStyle, ?automaticNumbericColumn) =
        {
            IsOverride = defaultArg isOverride true
            SheetName = defaultArg sheetName SHEET1
            ColumnAutofitOptions = defaultArg columnAutofitOptions ColumnAutofitOptions.DefaultValue
            TableName = defaultArg tableName DefaultTableName
            AutoNumbericColumn = defaultArg automaticNumbericColumn false
            TableStyle = defaultArg tableStyle <| DefaultTableStyle()
        }

    static member DefaultValue = TableXlsxSavingOptions.Create()       


type ITableColumnKey =
    abstract member StringIC: unit -> StringIC


type ITableColumnKey<'Value when 'Value :> IConvertible> =
    inherit ITableColumnKey

type ITableUnionColumnKey<'Value when 'Value :> IConvertible> =
    inherit ITableColumnKey<'Value>
    abstract member Unbox: IConvertible -> 'Value

[<AutoOpen>]
module __ITableColumnKeyExtensions =

    [<Extension>]
    type _ITableColumnKeyExtensions = 
        [<Extension>]
        static member GetBy(row: ObjectSeries<StringIC>, columnKey: ITableColumnKey<'Value>) =
            match columnKey with 
            | :? ITableUnionColumnKey<'Value> as iTableRowMapping ->
                row.GetAs(columnKey.StringIC())
                |> iTableRowMapping.Unbox
                
            | _ -> row.GetAs<'Value>(columnKey.StringIC())

        //[<Extension>]
        //static member TryGetBy(row: ObjectSeries<StringIC>, columnKey: #ITableColumnKey<'Value>) =
        //    row.TryGetAs<'Value>(columnKey.StringIC())

        [<Extension>]
        static member GetColumnBy(frame: Frame<_ ,StringIC>, columnKey: ITableColumnKey<'Value>) =
            match columnKey with 
            | :? ITableUnionColumnKey<'Value> as iTableRowMapping ->
                frame.GetColumn(columnKey.StringIC())
                |> Series.mapValues iTableRowMapping.Unbox
                
            | _ -> frame.GetColumn<'Value>(columnKey.StringIC())

        [<Extension>]
        static member TryGetColumnBy(frame: Frame<_ ,StringIC>, columnKey: ITableColumnKey<'Value>) =
            match columnKey with 
            | :? ITableUnionColumnKey<'Value> as iTableRowMapping ->
                frame.TryGetColumn(columnKey.StringIC(), Lookup.Exact)
                |> OptionalValue.asOption
                |> Option.map (Series.mapValues iTableRowMapping.Unbox)
                
            | _ -> 
                frame.TryGetColumn<'Value>(columnKey.StringIC(), Lookup.Exact)
                |> OptionalValue.asOption
            

        [<Extension>]
        static member GroupRowsBy(frame: Frame<_ ,StringIC>, columnKey: #ITableColumnKey<'Value>) =
            frame
            |> Frame.groupRowsUsing(fun _ row ->
                row.GetBy(columnKey)
            )

    type Frame with
        static member GroupRowsBy(columnKey: #ITableColumnKey<'Value>) =
            fun (frame: Frame<_, StringIC>) ->
                frame.GroupRowsBy(columnKey)





type Table = private Table of ExcelFrame<StringIC>
with 
    member internal x.AsExcelFrame = 
        let (Table frame) = x
        frame

    member x.AsFrame = 
        let (Table frame) = x
        frame.AsFrame

    member x.Rows = x.AsFrame.Rows

    member x.Columns = x.AsFrame.Columns

    member x.GetColumnBy(columnKey) = x.AsFrame.GetColumnBy(columnKey)


    member table.MapFrame(mapping) =
        ExcelFrame.mapFrame mapping table.AsExcelFrame
        |> Table

    member x.RowCount = x.AsFrame.RowCount

    member x.Headers =
        x.AsFrame.ColumnKeys

    member x.AutomaticNumbericColumn() =
        (ExcelFrame.autoNumbericColumns) x.AsExcelFrame
        |> Table
        

    member x.AddColumns(columns, ?position) =
        let position = 
            defaultArg position DataAddingPosition.AfterLast

        x.MapFrame(fun frame ->

            let originColumnKeys = frame.ColumnKeys |> Seq.indexed |> Seq.map (fun (i, v) -> (i, -1), v) |> List.ofSeq
            
            let addingColumns =
                let addingIndex = 
                    match position with 
                    | DataAddingPosition.BeforeFirst -> -1
                    | DataAddingPosition.ByIndex i -> i 
                    | DataAddingPosition.AfterLast -> originColumnKeys.Length - 1
                    | DataAddingPosition.After k ->
                        originColumnKeys
                        |> List.tryFindIndex(fun (_, k') -> k' = k)
                        |> function
                            | Some i -> i
                            | None -> failwithf "Invalid column adding position %A, all column keys are %A" position originColumnKeys

                    | DataAddingPosition.Before k ->
                        originColumnKeys
                        |> List.tryFindIndex(fun (_, k') -> k' = k)
                        |> function
                            | Some i -> i-1
                            | None -> failwithf "Invalid column adding position %A, all column keys are %A" position originColumnKeys


                columns
                |> Seq.mapi (fun i (columnKey, column) ->
                    ((addingIndex, i), columnKey), column
                )

            let frame = Frame.indexColsWith originColumnKeys frame

            let newFrame = 
                (frame, addingColumns)
                ||> Seq.fold(fun frame (addingColumnKey, addingColumn) ->
                    Frame.addCol addingColumnKey addingColumn frame
                )
                |> Frame.sortColsByKey
                |> Frame.mapColKeys snd

            newFrame
        )

    member x.AddColumn(key, values, ?position) =
        x.AddColumns([key, values], ?position = position)

    member x.AddColumn(key, values, ?position) =
        let column = Series.ofValues values |> Series.indexWith (x.AsFrame.RowKeys)
        x.AddColumn(key, column, ?position = position)

    member x.AddColumn_SingtonValue(key, value, ?position) =
        let values = List.replicate x.AsFrame.RowCount value
        x.AddColumn(key, values, ?position = position)



    member x.ToArray2D(?usingITableCellMapping:bool, ?automaticNumbericColumn: bool) =
        let x =
            match defaultArg automaticNumbericColumn false with 
            | true -> x.AutomaticNumbericColumn()
            | false -> x

        
        let usingITableCellMapping = defaultArg  usingITableCellMapping false
        let array2D =
            ExcelFrame.toArray2DWithHeader x.AsExcelFrame

        match usingITableCellMapping with 
        | true ->
            array2D
            |> Array2D.map (fun v ->
                match v with 
                | :? ITableCellValue as cell -> cell.Value()
                | :? ITableCell as cell -> (cell.CellText() :> IConvertible)
                | _ -> v
            )

        | false -> array2D
        
    
    member x.ToExcelArray() =
        let array2D = x.ToArray2D()
        let newArray2D = array2D.[1.., *]
        ExcelArray.Convert newArray2D


    member x.AddToExcelPackage
        (excelPackage: ExcelPackage,
         sheetName,
         tableName,
         ?addr, 
         ?usingITableCellMapping:bool,
         ?columnAutofitOptions,
         ?automaticNumbericColumn,
         ?tableStyle) =

        let worksheet = 
            let worksheets = excelPackage.Workbook.Worksheets
            worksheets
            |> Seq.tryFind(fun sheet -> StringIC sheet.Name = StringIC sheetName)
            |> function
                | Some sheet -> VisibleExcelWorksheet.Create sheet 
                | None ->
                    excelPackage.Workbook.Worksheets.Add(sheetName)
                    |> VisibleExcelWorksheet.Create

        let array2D = x.ToArray2D(?usingITableCellMapping = usingITableCellMapping, ?automaticNumbericColumn = automaticNumbericColumn)

        worksheet.LoadFromArraysAsTable(
            array2D |> Array2D.map box,
            columnAutofitOptions = defaultArg columnAutofitOptions ColumnAutofitOptions.DefaultValue,
            tableName = tableName, 
            tableStyle = defaultArg tableStyle (DefaultTableStyle()),
            ?addr = addr
        )

        array2D
        |> Array2D.iteri(fun row col v ->
            match v with 
            | :? DateTime -> worksheet.Value.Cells.[row+1, col+1].Style.Numberformat.Format <- System.Globalization.DateTimeFormatInfo.CurrentInfo.ShortDatePattern
            | _ -> ()
        )


    member x.SaveToXlsx(path: string, ?usingITableCellMapping:bool, ?tableXlsxSavingOptions: TableXlsxSavingOptions) =
        do (path |> XlsxPath |> ignore)
        let savingOptions = defaultArg tableXlsxSavingOptions TableXlsxSavingOptions.DefaultValue

        if savingOptions.IsOverride
        then File.Delete path

        let excelPackage = new ExcelPackage(FileInfo(path))
        x.AddToExcelPackage(
            excelPackage,
            savingOptions.SheetName,
            savingOptions.TableName,
            ?usingITableCellMapping = usingITableCellMapping,
            columnAutofitOptions = savingOptions.ColumnAutofitOptions,
            automaticNumbericColumn = savingOptions.AutoNumbericColumn,
            tableStyle = savingOptions.TableStyle)

        excelPackage.Save()
        excelPackage.Dispose()


    member x.SaveToXlsx_ITableCellMapping(path: string, ?tableXlsxSavingOptions: TableXlsxSavingOptions) =
        x.SaveToXlsx(path = path, usingITableCellMapping = true, ?tableXlsxSavingOptions = tableXlsxSavingOptions)

    member x.ToText(?usingITableCellMapping:bool) =
        let array2D = x.ToArray2D(?usingITableCellMapping = usingITableCellMapping)

        let contents = 
            array2D
            |> Array2D.map string
            |> Array2D.toLists
            |> List.map (String.concat ", ")
            |> String.concat "\n"

        contents

    member x.SaveToTxt(path: TxtPath, ?usingITableCellMapping:bool, ?isOverride) =
        let path = path.Path
        let isOverride = defaultArg isOverride true
        if isOverride
        then File.Delete path

        let contents = x.ToText(?usingITableCellMapping = usingITableCellMapping)
      
        File.WriteAllText(path, contents, Text.Encoding.UTF8)

    member x.ToText_ITableCellMapping() =
        x.ToText(usingITableCellMapping = true)

    member x.SaveToTxt_ITableCellMapping(path: TxtPath, ?isOverride) =
        x.SaveToTxt(path = path, usingITableCellMapping = true, ?isOverride = isOverride)


    member x.GetColumns (indexes: seq<StringIC>)  =

        let mapping =
            Frame.filterCols (fun key _ ->
                Seq.exists (fun index -> key = index) indexes
            )

        x.MapFrame mapping 


    static member OfArray2D (array2D: IConvertible[,]) =
        let frame = ExcelFrame.ofArray2DWithHeader_Convititable array2D
        frame
        |> ExcelFrame.mapFrame(Frame.mapColKeys StringIC)
        |> Table.Table

    static member OfArray2D (array2D: obj[,]) =
        array2D
        |> Array2D.map fixContent
        |> Table.OfArray2D
   


    static member OfRecords (records: seq<'record>) =
        
        ExcelFrame.ofRecords records
        |> ExcelFrame.mapFrame(fun frame ->
            Frame.mapColKeys StringIC frame
            //let colKeys = 
            //    typeof<'record>.GetProperties() |> Array.map (fun m -> m.Name)
            //frame 
            //|> Frame.mapColKeys (fun colKey ->
            //    Seq.findIndex ((=) colKey) colKeys, StringIC colKey
            //)
            //|> Frame.sortColsByKey
            //|> Frame.mapColKeys snd
        )
        |> Table.Table

    static member OfFrame (frame: Frame<int, string>) =
        frame
        |> ExcelFrame.ofFrameWithHeaders 
        |> ExcelFrame.mapFrame(Frame.mapColKeys StringIC)
        |> Table.Table

    static member OfFrame (frame: Frame<int, StringIC>) =
        frame
        |> Frame.mapColKeys StringIC.value
        |> Table.OfFrame

    static member OfRowsOrdinal (rows: Series<StringIC, IConvertible> seq) =
        Frame.ofRowsOrdinal rows
        |> Frame.mapRowKeys int
        |> Table.OfFrame

    static member OfRowsOrdinal (rows: seq<Observations>) =
        rows
        |> Seq.map (fun observations ->
            observations.ObservationValues
            |> series
        )
        |> Table.OfRowsOrdinal


    static member OfExcelPackage(excelPackage: ExcelPackageWithXlsxFile, ?rangeGettingOptions, ?sheetGettingOptions) =
        let rangeGettingOptions = 
            defaultArg rangeGettingOptions RangeGettingOptions.UserRange

        let sheetGettingOptions =
            defaultArg sheetGettingOptions SheetGettingOptions.DefaultValue

        let excelRangeInfo = ExcelRangeContactInfo.readFromExcelPackages rangeGettingOptions sheetGettingOptions excelPackage
        
        Table.OfArray2D(excelRangeInfo.Content)

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


and private TableSurrogate = TableSurrogate of IConvertible[,]
with 
    interface ISurrogate with 
        member x.FromSurrogate(system) = 
            let (TableSurrogate array2D) = x
            Table.OfArray2D array2D :> ISurrogated


type Tables = Tables of Table al1List
with    
    
    member x.SaveToXlsx(path: string, ?usingITableCellMapping:bool,  ?tableXlsxSavingOptions: TableXlsxSavingOptions) =
        let (Tables v) = x
        do (path |> XlsxPath |> ignore)
        let savingOptions = 
            { defaultArg tableXlsxSavingOptions TableXlsxSavingOptions.DefaultValue with 
                SheetName = "Sheet"
                TableName = "Table" }

        if savingOptions.IsOverride
        then File.Delete path

        let excelPackage = new ExcelPackage(FileInfo(path))

        v.AsList
        |> List.iteri (fun i table ->
            let i = i + 1
            let sheetName = sprintf "%s%d" savingOptions.SheetName i
            let tableName = sprintf "%s%d" savingOptions.TableName i
            table.AddToExcelPackage(
                excelPackage,
                sheetName,
                tableName,
                ?usingITableCellMapping = usingITableCellMapping,
                columnAutofitOptions = savingOptions.ColumnAutofitOptions,
                automaticNumbericColumn = savingOptions.AutoNumbericColumn,
                tableStyle = savingOptions.TableStyle)
        )

        excelPackage.Save()
        excelPackage.Dispose()


    member x.SaveToXlsx_ITableCellMapping(path: string,  ?tableXlsxSavingOptions: TableXlsxSavingOptions) =
        x.SaveToXlsx(path = path, usingITableCellMapping = true, ?tableXlsxSavingOptions = tableXlsxSavingOptions)


[<RequireQualifiedAccess>]
module Table =
    let [<Literal>] ``#N/A`` = "#N/A"

    let mapFrame mapping (table: Table) =
        table.MapFrame mapping

    let fillEmptyUp (table: Table) =
        table 
        |> mapFrame Frame.fillEmptyUp

    let splitRowToMany addtionalHeaders mapping (table: Table) =
        let addtionalHeaders =
            addtionalHeaders
            |> List.map StringIC

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

    let removeAutoNamedColumns (table: Table) =
        table 
        |> mapFrame(
            Frame.filterCols(fun col _ ->
                let parser =
                    (pstringCI "Column" .>> pint32) 
                    <|>
                    (pstringCI CELL_SCRIPT_COLUMN .>> pint32) 
                match  run (parser) col.Value with 
                | Success _ -> false
                | _ -> true
            )
        )

    let toExcelArray (Table excelFrame) =
        excelFrame
        |> ExcelFrame.mapFrame (fun frame ->
            let sequenceHeaders = 
                frame.ColumnKeys
                |> Seq.mapi (fun i _ -> i)
            Frame.indexColsWith sequenceHeaders frame
        )
        |> ExcelArray

    let concat_RefreshRowKeys (tables: AtLeastOneList<Table>) =
        tables.Head
        |> mapFrame(fun _ ->
            tables
            |> AtLeastOneList.map (fun table -> table.AsFrame)
            |> Frame.Concat_RefreshRowKeys
        )


    [<System.Obsolete("This method is obsolted, using concat_RefreshRowKeys instead")>]
    let concat_RemoveRowKeys (tables: AtLeastOneList<Table>) =
        concat_RefreshRowKeys(tables)


    let concat_GroupByHeaders_RefreshRowKeys (tables: AtLeastOneList<Table>) =
        tables.AsList
        |> List.groupBy(fun m ->
            Set.ofSeq m.Headers
        )
        |> List.map(fun (_, tables) ->
            tables
            |> AtLeastOneList.Create
            |> concat_RefreshRowKeys
        )
        |> AtLeastOneList.Create
        |> Tables

    [<System.Obsolete("This method is obsolted, using concat_RefreshRowKeys instead")>]
    let concat_GroupByHeaders_RemoveRowKeys (tables: AtLeastOneList<Table>) =
        concat_GroupByHeaders_RefreshRowKeys(tables)

      

type TableConcatProxy<'T>(getter: 'T -> Table, setter: Table -> 'T) =
    member x.Concat_RefreshRowKeys(tables: 'T al1List) =
        tables
        |> AtLeastOneList.map(getter)
        |> Table.concat_RefreshRowKeys
        |> setter

    member x.Concat_GroupByHeaders_RefreshRowKeys(tables: 'T al1List) =
        let (Tables tables) = 
            tables
            |> AtLeastOneList.map(getter)
            |> Table.concat_GroupByHeaders_RefreshRowKeys

        tables
        |> AtLeastOneList.map setter

    member x.Bind(getter': 'T2 -> 'T, setter': 'T -> 'T2) =
        TableConcatProxy(getter' >> getter, setter >> setter')


        

