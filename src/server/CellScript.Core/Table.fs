namespace CellScript.Core
#nowarn "0104"
open Deedle
open CellScript.Core.Extensions
open Akka.Util
open System
open Shrimp.FSharp.Plus
open System.IO
open OfficeOpenXml
open System.Runtime.CompilerServices
open FParsec
open CellScript.Core.Constrants
open System.Collections.Generic



type TableXlsxSavingOptions =
    { IsOverride: bool 
      SheetName: string
      ColumnAutofitOptions: ColumnAutofitOptions
      TableName: string
      TableStyle: Table.TableStyles
      IncludingFormula: bool
      CellSavingFormat: CellSavingFormat }
with 
    static member Create(?isOverride: bool, ?sheetName: string, ?columnAutofitOptions, ?tableName: string, ?tableStyle, ?cellFormat, ?includingFormula) =
        {
            IsOverride = defaultArg isOverride true
            SheetName = defaultArg sheetName SHEET1
            ColumnAutofitOptions = defaultArg columnAutofitOptions ColumnAutofitOptions.DefaultValue
            TableName = defaultArg tableName DefaultTableName
            TableStyle = defaultArg tableStyle <| DefaultTableStyle()
            CellSavingFormat = defaultArg cellFormat CellSavingFormat.KeepOrigin
            IncludingFormula = defaultArg includingFormula false
        }



    static member DefaultValue = TableXlsxSavingOptions.Create()       




type ITableColumnKey =
    abstract member StringIC: unit -> StringIC

type ITableColumnKey<'Value when 'Value :> IConvertible> =
    inherit ITableColumnKey
    //abstract member ToConvertible: 'Value -> IConvertible


type ITableColumnKeyEx<'Value when 'Value :> ICellValue> =
    inherit ITableColumnKey
    abstract member Unbox: IConvertible -> 'Value


[<AutoOpen>]
module __ITableColumnKeyExtensions =
    
    [<Extension>]
    type _ITableColumnKeyExtensions = 
        [<Extension>]
        static member TryGetAsEx<'Value>(row: ObjectSeries<StringIC>, key: StringIC, ?checkKeyValid): OptionalValue<'Value> =
            match defaultArg checkKeyValid false with 
            | true ->
                match row.Index.Locate(key) with 
                | EqualTo Addressing.Address.invalid -> 
                    raise (new KeyNotFoundException(sprintf "The key %O is not present in the index" key))
                    
                | addr -> 
                    row.Vector.GetValue(addr)
                    |> OptionalValue.map (fun v -> 
                        let conversionKind = ConversionKind.Flexible
                        Deedle.Internal.Convert.convertType<'Value> conversionKind v)
            | false -> row.TryGetAs<'Value>(key)

        [<Extension>]
        static member GetBy(row: ObjectSeries<StringIC>, columnKey: ITableColumnKey<'Value>) =
            row.GetAs<'Value>(columnKey.StringIC())

        [<Extension>]
        static member GetField<'T>(row: ObjectSeries<StringIC>, columnKey: string) =
            row.GetAs<'T>(StringIC columnKey)

        [<Extension>]
        static member TryGetBy(row: ObjectSeries<StringIC>, columnKey: ITableColumnKey<'Value>, ?checkKeyValid) =
            row.TryGetAsEx<'Value>(columnKey.StringIC(), ?checkKeyValid = checkKeyValid)


        [<Extension>]
        static member GetAsDateTime(row: ObjectSeries<StringIC>, key: StringIC) =
            row.GetAs<obj>(key)
            |> CellText.parseToDateTime


        [<Extension>]
        static member GetAsConveritableUnion(row: ObjectSeries<StringIC>, key: StringIC) =
            match row.GetAs<obj>(key) with 
            | :? IConvertible as convertible -> convertible 
            | :? ICellValue as v -> v.Convertible
            | v -> failwithf "type of cell value %A should either be ITableCellValue or IConvertible" (v.GetType())
            |> ConvertibleUnion.Convert

        [<Extension>]
        static member TryGetAsConveritableUnion(row: ObjectSeries<StringIC>, key: StringIC, ?checkKeyValid) =
            match row.TryGetAsEx<obj>(key, ?checkKeyValid = checkKeyValid) with 
            | OptionalValue.Present convertible ->
                match convertible with
                | :? IConvertible as convertible -> convertible 
                | :? ICellValue as v -> v.Convertible
                | v -> failwithf "type of cell value %A should either be ITableCellValue or IConvertible" (v.GetType())
                |> ConvertibleUnion.Convert
            | OptionalValue.Missing _ -> ConvertibleUnion.Missing

        [<Extension>]
        static member GetBy(row: ObjectSeries<StringIC>, columnKey: ITableColumnKeyEx<'Value>) =
            let convertible = row.GetAs<obj>(columnKey.StringIC())
            match convertible with 
            | :? IConvertible as convertible -> convertible |> columnKey.Unbox
            | :? ICellValue as v -> v :?> 'Value
            | _ -> failwithf "type of cell value %A should either be ITableCellValue or IConvertible" (convertible.GetType())



        [<Extension>]
        static member TryGetBy(row: ObjectSeries<StringIC>, columnKey: ITableColumnKeyEx<'Value>, ?checkKeyValid) =
            let convertible = row.TryGetAsEx<obj>(columnKey.StringIC(), ?checkKeyValid = checkKeyValid)
            match convertible with
            | OptionalValue.Present convertible ->
                match convertible with
                | :? IConvertible as convertible -> convertible |> columnKey.Unbox
                | :? ICellValue as v -> v :?> 'Value
                | _ -> failwithf "type of cell value %A should either be ITableCellValue or IConvertible" (convertible.GetType())
                |> Some
            | OptionalValue.Missing _ -> None

        [<Extension>]
        static member ObservationsAllEx(row: ObjectSeries<string>) =
            row.ObservationsAll
            |> List.ofSeq
            |> List.map(fun pair ->
                let value =
                    match OptionalValue.asOption pair.Value with 
                    | Some convertible ->
                        match convertible with
                        | :? IConvertible as convertible -> convertible |> ConvertibleUnion.Convert
                        | :? ICellValue as v -> v.Convertible |> ConvertibleUnion.Convert
                        | _ -> failwithf "type of cell value %A should either be ITableCellValue or IConvertible" (convertible.GetType())
                    | None -> ConvertibleUnion.Missing

                observation(StringIC pair.Key, value)
            )
            |> observations


        [<Extension>]
        static member ObservationsAllEx(row: Series<StringIC, obj>) =
            row.ObservationsAll
            |> List.ofSeq
            |> List.map(fun pair ->
                let value =
                    match OptionalValue.asOption pair.Value with 
                    | Some convertible ->
                        match convertible with
                        | :? IConvertible as convertible -> convertible |> ConvertibleUnion.Convert
                        | :? ICellValue as v -> v.Convertible |> ConvertibleUnion.Convert
                        | _ -> failwithf "type of cell value %A should either be ITableCellValue or IConvertible" (convertible.GetType())
                    | None -> ConvertibleUnion.Missing

                observation(pair.Key, value)
            )
            |> observations

        [<Extension>]
        static member ObservationsAllEx(row: ObjectSeries<StringIC>) =
            row.ObservationsAll
            |> List.ofSeq
            |> List.map(fun pair ->
                let value =
                    match OptionalValue.asOption pair.Value with 
                    | Some convertible ->
                        match convertible with
                        | :? IConvertible as convertible -> convertible |> ConvertibleUnion.Convert
                        | :? ICellValue as v -> v.Convertible |> ConvertibleUnion.Convert
                        | _ -> failwithf "type of cell value %A should either be ITableCellValue or IConvertible" (convertible.GetType())
                    | None -> ConvertibleUnion.Missing

                observation(pair.Key, value)
            )
            |> observations
          

        [<Extension>]
        static member GetColumnBy(frame: Frame<_ ,StringIC>, columnKey: ITableColumnKey<'Value>) =
            frame.GetColumn<'Value>(columnKey.StringIC())

        [<Extension>]
        static member GetColumnAsConveritableUnion(frame: Frame<_ ,StringIC>, key: StringIC) =
            frame.GetColumn<obj>(key)
            |> Series.mapValues (fun convertible ->
                match convertible with 
                | :? IConvertible as convertible -> convertible
                | :? ICellValue as v -> v.Convertible
                | _ -> failwithf "type of cell value %A should either be ITableCellValue or IConvertible" (convertible.GetType())
                |> ConvertibleUnion.Convert
            )

        [<Extension>]
        static member GetColumnBy(frame: Frame<_ ,StringIC>, columnKey: ITableColumnKeyEx<'Value>) =
            frame.GetColumn<obj>(columnKey.StringIC())
            |> Series.mapValues (fun convertible ->
                match convertible with 
                | :? IConvertible as convertible -> convertible |> columnKey.Unbox
                | :? ICellValue as v -> v :?> 'Value
                | _ -> failwithf "type of cell value %A should either be ITableCellValue or IConvertible" (convertible.GetType())
            )


        [<Extension>]
        static member TryGetColumnBy(frame: Frame<_ ,StringIC>, columnKey: ITableColumnKey<'Value>) =
            frame.TryGetColumn<'Value>(columnKey.StringIC(), Lookup.Exact)
            |> OptionalValue.asOption
         
        [<Extension>]
        static member TryGetColumnBy(frame: Frame<_ ,StringIC>, columnKey: ITableColumnKeyEx<'Value>) =
            frame.TryGetColumn<obj>(columnKey.StringIC(), Lookup.Exact)
            |> OptionalValue.asOption
            |> Option.map (Series.mapValues (fun convertible ->
                match convertible with 
                | :? IConvertible as convertible -> convertible |> columnKey.Unbox
                | :? ICellValue as v -> v.Convertible :?> 'Value
                | _ -> failwithf "type of cell value %A should either be ITableCellValue or IConvertible" (convertible.GetType())
            ))

        [<Extension>]
        static member TryGetColumnAsConveritableUnion(frame: Frame<_ ,StringIC>, key: StringIC) =
            frame.TryGetColumn<obj>(key, Lookup.Exact)
            |> OptionalValue.asOption
            |> Option.map(fun series ->
                series
                |> Series.mapValues (fun convertible ->
                    match convertible with 
                    | :? IConvertible as convertible -> convertible
                    | :? ICellValue as v -> v.Convertible
                    | _ -> failwithf "type of cell value %A should either be ITableCellValue or IConvertible" (convertible.GetType())
                    |> ConvertibleUnion.Convert
                )
            )

        [<Extension>]
        static member GroupRowsBy(frame: Frame<_ ,StringIC>, columnKey: ITableColumnKey<'Value>) =
            frame
            |> Frame.groupRowsUsing(fun _ row ->
                row.GetBy(columnKey)
            )

        [<Extension>]
        static member GroupRowsBy(frame: Frame<_ ,StringIC>, columnKey: ITableColumnKeyEx<'Value>) =
            frame
            |> Frame.groupRowsUsing(fun _ row ->
                row.GetBy(columnKey)
            )

    type Frame with
        static member GroupRowsBy(columnKey: ITableColumnKey<'Value>) =
            fun (frame: Frame<_, StringIC>) ->
                frame.GroupRowsBy(columnKey)

        static member GroupRowsBy(columnKey: ITableColumnKeyEx<'Value>) =
            fun (frame: Frame<_, StringIC>) ->
                frame.GroupRowsBy(columnKey)


type private HeadOrEnd =
    | Head = 0
    | End = 1

type internal XLColumn =
    { Header: IConvertible 
      Contents: IConvertible list }

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

    member x.GetColumnBy(columnKey: ITableColumnKey<_>) = x.AsFrame.GetColumnBy(columnKey)
    member x.GetColumnBy(columnKey: ITableColumnKeyEx<_>) = x.AsFrame.GetColumnBy(columnKey)

    [<System.ObsoleteAttribute>]
    static member Internal_OfFrame(frame: Frame<int, StringIC>) =
        ExcelFrame frame
        |> Table

    member table.MapFrame(mapping) =
        ExcelFrame.mapFrame mapping table.AsExcelFrame
        |> Table

    member x.RowCount = x.AsFrame.RowCount

    member x.IsEmpty = x.RowCount = 0
    member x.Headers = x.AsFrame.ColumnKeys

    member private x.MoveColumnsTo(columnKeys: StringIC list, headOrEnd: HeadOrEnd) =

        let originHeaders = 
            x.Headers
            |> List.ofSeq

        let alreadyDone = 
            (
                match originHeaders.Length > columnKeys.Length with 
                | true ->
                    match headOrEnd with 
                    | HeadOrEnd.Head -> List.take columnKeys.Length originHeaders = columnKeys
                    | HeadOrEnd.End -> 
                        List.skip (originHeaders.Length - columnKeys.Length) originHeaders = columnKeys
                | false -> false
            )
            || (
                let intersected = Set.intersect (Set.ofList originHeaders) (Set.ofList columnKeys)
                intersected.Count = 0
            )

        match alreadyDone with 
        | true -> x
        | false ->
            let originHeaders =
                originHeaders
                |> List.mapi(fun originIndex originHeader ->
                    let index = 
                        columnKeys
                        |> List.tryFindIndex(fun m -> m = originHeader)

                    let index = 
                        match index with 
                        | Some index -> 
                            let sortingIndex = 
                                match headOrEnd with 
                                | HeadOrEnd.Head -> -1
                                | HeadOrEnd.End -> 1

                            sortingIndex, index
                        | None -> 0, originIndex

                    originHeader, index
                )
                |> dict

            x.MapFrame(fun frame ->
                frame
                |> Frame.mapColKeys(fun colKey -> originHeaders.[colKey], colKey)
                |> Frame.sortColsByKey
                |> Frame.mapColKeys snd
            )

        //match columnKeys.Length <= originHeaders.Length with 
        //| true -> 


        //| false -> 
        //    let notExistsHeader =
        //        columnKeys
        //        |> List.find(fun m -> not (List.contains m originHeaders))

        //    failwithf "%A not exists in headers" notExistsHeader


    member x.MoveColumnsToHead(columnKeys: StringIC list) =
        x.MoveColumnsTo(columnKeys, HeadOrEnd.Head)

    member x.MoveColumnsToEnd(columnKeys: StringIC list) =
        x.MoveColumnsTo(columnKeys, HeadOrEnd.End)

    member x.AddColumns(columns, ?position) =
        let columns = List.ofSeq columns 
        match columns with 
        | [] -> x
        | _ ->
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
                    |> Seq.mapi (fun i (columnKey, column: #Series<_, _>) ->
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

    member tb.AddColumns_SingtonValue(keys, value, ?position) =
        let rowCount = tb.RowCount

        let columns =
            let rowKeys =
                tb.AsFrame.RowKeys

            let columnValues =
                let values = List.replicate rowCount value
                Series.ofValues values
                |> Series.indexWith rowKeys

            keys
            |> List.map(fun column ->
                column => columnValues
            )

        tb.AddColumns(columns, ?position = position)

    member x.AddColumn(key, values, ?position) =
        x.AddColumns([key, values], ?position = position)

    member x.AddColumn(key, values, ?position) =
        let column = Series.ofValues values |> Series.indexWith (x.AsFrame.RowKeys)
        x.AddColumn(key, column, ?position = position)

    member x.AddColumn_SingtonValue(key, value, ?position) =
        let values = List.replicate x.AsFrame.RowCount value
        x.AddColumn(key, values, ?position = position)

    member x.AddColumn_FirstRow(key, value: 'a, ?position) =
        let values = 
            value :: (List.replicate (x.AsFrame.RowCount-1) (Unchecked.defaultof<'a>))
        x.AddColumn(key, values, ?position = position)


    /// with header
    member frame.ToArray2D(?eraseAutomaticHeaders: bool) = 
        let headers = 
            frame.Headers
            |> List.ofSeq


        let headers = 
            match defaultArg eraseAutomaticHeaders false with 
            | true -> 
                headers
                |> List.map(fun m -> 
                    match m.Value with 
                    | String.StartsWithIC "CellScriptColumn" -> ""
                    | _ -> m.Value
                )

            | false -> 
                headers
                |> List.map(fun m -> m.Value)

        let header = 
            headers 
            |> List.map (fun m -> (m :> IConvertible))


        let contents = 
            frame.AsExcelFrame
            |> ExcelFrame.toArray2D
            |> Array2D.toLists

        let result =
            List.append [header] contents
            |> array2D
    
        result

    member private x.FormatText = 
        x.AsFrame.FillMissing("").Format(50)
    
    member internal x.Array2DValue = x.ToArray2D()

    member internal x.XLColumnValues = 
        let array2D = 
            x.ToArray2D()
            |> Array2D.transpose

        array2D
        |> Array2D.toLists
        |> List.map(fun datas ->
            { Header = datas.Head 
              Contents = datas.Tail }
        )

    member internal x.XLRowValues = 
        let array2D = 
            x.ToArray2D()

        array2D
        |> Array2D.toLists
    

    /// without header
    member x.ToExcelArray(?withHeader) =
        
        let array2D = x.ToArray2D()
        let newArray2D = array2D.[1.., *]
        ExcelArray.Convert newArray2D


    member x.ConvertDataForSave(cellSavingFormat: CellSavingFormat) =
        let x = 
            x.MapFrame(fun frame ->
                let frame = Frame.indexRowsOrdinally frame

                frame
                |> Frame.mapCols(fun colkey colValues ->
                    match cellSavingFormat with 
                    | CellSavingFormat.KeepOrigin -> colValues :> ISeries<_>
                    | CellSavingFormat.ODBC -> 
                        let colValues = 
                            colValues.ValuesAll
                            |> List.ofSeq
                            |> List.map readContentAsConvertible 

                        let numbers =
                            colValues
                            |> List.choose(fun m -> 
                                let m = m.ToString()
                                CellText.getAsODBCNumber m
                            )

                        match numbers.Length = colValues.Length with 
                        | true -> 
                            Series.ofValues numbers :> ISeries<_>
                            
                        | false ->
                            let series = 
                                colValues
                                |> List.map (fun m ->
                                    match m with 
                                    | null -> ""
                                    | _ -> string m
                                )
                                |> Series.ofValues

                            series :> ISeries<_>

                        

                     
                    | CellSavingFormat.AutomaticNumberic ->
                        colValues
                        |> Series.mapValues(fun m ->
                            let m = readContentAsConvertible m
                            match System.Double.TryParse (m.ToString()) with 
                            | true, v -> v :> IConvertible
                            | false, _ -> m
                        ) :> ISeries<_>

                    | CellSavingFormat.AllText excludings ->
                        match List.contains colkey excludings with 
                        | true -> colValues :> ISeries<_>
                        | false ->
                            colValues
                            |> Series.mapValues(fun m ->
                                let m = readContentAsConvertible m
                                m.ToString() :> IConvertible
                            ) :> ISeries<_>
                )
            )

        x



    member x.AddToExcelPackage
        (excelPackage: ExcelPackage,
         sheetName,
         tableName,
         ?includingFormula,
         ?addr, 
         ?columnAutofitOptions,
         ?tableStyle,
         ?cellFormat) =

        let worksheet = excelPackage.GetOrAddWorksheet(sheetName)
            
        let cellFormat = defaultArg cellFormat CellSavingFormat.KeepOrigin

        let x = x.ConvertDataForSave(cellFormat)

        let array2D = x.ToArray2D()


        worksheet.LoadFromArraysAsTable(
            array2D,
            columnAutofitOptions = defaultArg columnAutofitOptions ColumnAutofitOptions.DefaultValue,
            tableName = tableName, 
            tableStyle = defaultArg tableStyle (DefaultTableStyle()),
            ?addr = addr,
            ?includingFormula = includingFormula
        )

        array2D
        |> Array2D.iteri(fun row col v ->
            match v with 
            | :? DateTime -> worksheet.Value.Cells.[row+1, col+1].Style.Numberformat.Format <- System.Globalization.DateTimeFormatInfo.CurrentInfo.ShortDatePattern
            | _ -> 
                //match cellFormat with 
                //| CellSavingFormat.AllText -> 
                    
                ()
        )


    member x.SaveToXlsx(path: string, ?tableXlsxSavingOptions: TableXlsxSavingOptions) =    
        match x.RowCount with 
        | 0 -> failwith "Table is empty"
        | _ -> ()

        do (path |> XlsxPath |> ignore)
        let savingOptions = defaultArg tableXlsxSavingOptions TableXlsxSavingOptions.DefaultValue

        if savingOptions.IsOverride
        then 
            if File.Exists path
            then File.Delete path

        let excelPackage = new ExcelPackage(FileInfo(path))
        x.AddToExcelPackage(
            excelPackage,
            savingOptions.SheetName,
            savingOptions.TableName,
            includingFormula = savingOptions.IncludingFormula,
            columnAutofitOptions = savingOptions.ColumnAutofitOptions,
            tableStyle = savingOptions.TableStyle,
            cellFormat = savingOptions.CellSavingFormat)

        excelPackage.Save()
        excelPackage.Dispose()



    member x.ToText() =
        let array2D = x.ToArray2D()

        let contents = 
            array2D
            |> Array2D.map string
            |> Array2D.toLists
            |> List.map (String.concat ", ")
            |> String.concat "\n"

        contents

    member x.SaveToTxt(path: TxtPath, ?isOverride) =
        let path = path.Path
        let isOverride = defaultArg isOverride true
        if isOverride
        then File.Delete path

        let contents = x.ToText()
      
        File.WriteAllText(path, contents, Text.Encoding.UTF8)


    member x.GetColumns (indexes: seq<StringIC>)  =

        let mapping =
            Frame.filterCols (fun key _ ->
                Seq.exists (fun index -> key = index) indexes
            )

        x.MapFrame mapping 

    static member private OfArray2D (array: IConvertible[,]) =
        let fixHeaders headers =
            let headers = List.ofSeq headers
            headers |> List.mapi (fun i (header: IConvertible) ->
                match header with
                | null -> 
                    (sprintf "%s%d" CELL_SCRIPT_COLUMN i)
                | _ -> 
                    let isEmptyText = 
                        match header with 
                        | :? string as v -> v.Trim() = ""
                        | _ -> false

                    match isEmptyText with 
                    | true -> (sprintf "%s%d" CELL_SCRIPT_COLUMN i)
                    | false -> 

                        headers.[0 .. i - 1]
                        |> List.filter(fun preHeader -> preHeader = header)
                        |> function
                            | matchHeaders when matchHeaders.Length > 0 ->
                                let headerText = header.ToString()
                                sprintf "%s%d" headerText (matchHeaders.Length + 1) 
                            | [] -> header.ToString()
                            | _ -> failwith "invalid token"
            )
            |> List.map StringIC
    
        let array = Array2D.rebase array
    
        let headers = array.[0,*] |> fixHeaders 
    
        ExcelFrame.ofArray2DWithConvitable array.[1..,*] 
        |> ExcelFrame.mapFrame(Frame.indexColsWith headers)
        |> Table.Table

    static member OfArray2D (array2D: ConvertibleUnion[,]) =
        array2D
        |> Array2D.map (fun m -> m.Value)
        |> Table.OfArray2D

    static member OfArray2D (array2D: obj[,]) =
        array2D
        |> Array2D.map fixRawContent
        |> Table.OfArray2D


    static member private OfFrame_AllowICellValue (frame: Frame<int, StringIC>) =
        frame
        |> Frame.mapValues fixRawContent_AllowCellValue
        |> ExcelFrame
        |> Table.Table

    static member private OfFrame (frame: Frame<int, StringIC>) =
        frame
        |> Frame.mapValues fixRawContent
        |> ExcelFrame
        |> Table.Table

    static member private OfFrame (frame: Frame<int, string>) =
        frame
        |> Frame.mapColKeys StringIC
        |> Table.OfFrame


    static member OfRecords (records: seq<'record>) =
        (Frame.ofRecords records)
        |> Frame.mapColKeys StringIC
        |> Table.OfFrame

    static member OfRecords_AllowICellValue (records: seq<'record>) =
        (Frame.ofRecords records)
        |> Frame.mapColKeys StringIC
        |> Table.OfFrame_AllowICellValue

    static member OfColumns (columns: seq<StringIC * Series<int, ConvertibleUnion>>) =
        let columns =
            columns
            |> Seq.map(fun (key, series) ->
                key, series |> Series.mapValues (fun m -> m.Value)
            )

        (Frame.ofColumns columns)
        |> Table.OfFrame

    static member private OfRowsOrdinal (rows: Series<StringIC, IConvertible> seq) =
        Frame.ofRowsOrdinal rows
        |> Frame.mapRowKeys int
        |> Table.OfFrame

    static member OfRowsOrdinal (rows: Series<StringIC, ConvertibleUnion> seq) =
        let rows = rows |> Seq.map (Series.mapValues(fun m -> m.Value))
        Table.OfRowsOrdinal rows

    static member OfRowsOrdinal_AllowICellValue (rows: Series<StringIC, obj> seq) =
        let frame = 
            Frame.ofRowsOrdinal rows
            |> Frame.mapRowKeys int

        frame
        |> Table.OfFrame_AllowICellValue

    static member OfRowsOrdinal (rows: seq<Observations>) =
        let rows = 
            rows
            |> Seq.map (fun observations ->
                observations.ObservationValues
                |> series
            )
        
        Table.OfRowsOrdinal rows


    static member OfExcelPackage(excelPackage: ExcelPackageWithXlsxFile, ?rangeGettingOptions, ?sheetGettingOptions) =
        let rangeGettingOptions = 
            defaultArg rangeGettingOptions RangeGettingOptions.UserRange

        let sheetGettingOptions =
            defaultArg sheetGettingOptions SheetGettingOptions.DefaultValue

        let excelRangeInfo = ExcelRangeContactInfo.readFromExcelPackages rangeGettingOptions sheetGettingOptions excelPackage
        
        Table.OfArray2D(excelRangeInfo.Content |> Array2D.map (fun m -> m.Value))

    static member OfXlsxFileWithSheetName(xlsxFile: XlsxFile, ?rangeGettingOptions, ?sheetGettingOptions) =
        let rangeGettingOptions = 
            defaultArg rangeGettingOptions RangeGettingOptions.UserRange

        let sheetGettingOptions =
            defaultArg sheetGettingOptions SheetGettingOptions.DefaultValue

        let excelRangeInfo = ExcelRangeContactInfo.readFromFile rangeGettingOptions sheetGettingOptions xlsxFile
        let sheetName = StringIC excelRangeInfo.SheetName
        sheetName => Table.OfArray2D(excelRangeInfo.Content |> Array2D.map(fun m -> m.Value))

    static member OfXlsxFile(xlsxFile: XlsxFile, ?rangeGettingOptions, ?sheetGettingOptions) =
        Table.OfXlsxFileWithSheetName(xlsxFile, ?rangeGettingOptions = rangeGettingOptions, ?sheetGettingOptions = sheetGettingOptions)
        |> snd

    static member OfCsvFile(csvFile: CsvFile) =
        Frame.ReadCsv(csvFile.Path)
        |> Frame.mapColKeys StringIC
        |> ExcelFrame
        |> Table

    static member OfCsvFile_TXT(txtFile: TxtFile) =
        Frame.ReadCsv(txtFile.Path)
        |> Frame.mapColKeys StringIC
        |> ExcelFrame
        |> Table

    interface IToArray2D with 
        member x.ToArray2D() =
            x.ToArray2D()
            |> Array2D.map ConvertibleUnion.Convert

    interface ISurrogated with 
        member x.ToSurrogate(system) = 
            TableSurrogate ((x :> IToArray2D).ToArray2D()) :> ISurrogate


and private TableSurrogate = TableSurrogate of ConvertibleUnion[,]
with 
    interface ISurrogate with 
        member x.FromSurrogate(system) = 
            let (TableSurrogate array2D) = x
            Table.OfArray2D array2D :> ISurrogated


type Tables = Tables of Table al1List
with    
    
    static member OfXlsxFile_ZippedArgs(xlsxFile: XlsxFile, ?sheetAndRangeOptions: list<SheetGettingOptions * RangeGettingOptions>) =
        use xlsxFile = ExcelPackageWithXlsxFile.Create xlsxFile

        let r = 
            let sheets =
                match sheetAndRangeOptions with 
                | None -> 
                    xlsxFile.ExcelPackage.GetValidWorksheets()
                    |> List.map(fun sheet ->
                        sheet, RangeGettingOptions.UserRange
                    )
                | Some sheetAndRangeOptions ->
                    sheetAndRangeOptions
                    |> List.map(fun (sheetGettingOptions, rangeGettingOptions) ->
                        let sheet = xlsxFile.ExcelPackage.GetValidWorksheet(sheetGettingOptions)
                        sheet, rangeGettingOptions
                    )
                    
            sheets
            |> List.map(fun (sheet, rangeGettingOptions) ->
                let datas = sheet.ReadDatas(rangeGettingOptions)
                StringIC sheet.Name => Table.OfArray2D(datas.Content)
            )

        r
        |> dict

    static member OfXlsxFile(xlsxFile: XlsxFile, ?rangeGettingOptions, ?sheets) =
        use xlsxFile = ExcelPackageWithXlsxFile.Create xlsxFile
        let rangeGettingOptions = 
            defaultArg rangeGettingOptions RangeGettingOptions.UserRange

        let r = 
            let sheets =
                match sheets with 
                | None -> xlsxFile.ExcelPackage.GetValidWorksheets()
                | Some sheetNames ->
                    sheetNames
                    |> List.map(fun sheetName ->
                        xlsxFile.ExcelPackage.GetValidWorksheet(sheetName)
                    )
                    
            sheets
            |> List.map(fun sheet ->
                let datas = sheet.ReadDatas(rangeGettingOptions)
                StringIC sheet.Name => Table.OfArray2D(datas.Content)
            )

        r
        |> dict

    static member OfXlsxFile_ChooseSheetNames(xlsxFile: XlsxFile, sheets: string list, ?rangeGettingOptions) =
        use xlsxFile = ExcelPackageWithXlsxFile.Create xlsxFile
        let rangeGettingOptions = 
            defaultArg rangeGettingOptions RangeGettingOptions.UserRange

        let sheetNames = 
            [
                for sheet in xlsxFile.ExcelPackage.Workbook.Worksheets do
                    yield sheet.Name
            ]
            |> List.map StringIC

        let r = 
            let sheets =
                sheets
                |> List.map StringIC
                |> List.choose(fun sheetName ->
                    match List.contains (sheetName) sheetNames with 
                    | true -> 
                        xlsxFile.ExcelPackage.GetValidWorksheet(SheetGettingOptions.SheetName sheetName)
                        |> Some
                    | false -> None
                )
                    
            sheets
            |> List.map(fun sheet ->
                let datas = sheet.ReadDatas(rangeGettingOptions)
                StringIC sheet.Name => Table.OfArray2D(datas.Content)
            )

        r
        |> dict


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
                columnAutofitOptions = savingOptions.ColumnAutofitOptions,
                includingFormula = savingOptions.IncludingFormula,
                tableStyle = savingOptions.TableStyle,
                cellFormat = savingOptions.CellSavingFormat)
        )

        excelPackage.Save()
        excelPackage.Dispose()


    member x.SaveToXlsx_ITableCellMapping(path: string,  ?tableXlsxSavingOptions: TableXlsxSavingOptions) =
        x.SaveToXlsx(path = path, usingITableCellMapping = true, ?tableXlsxSavingOptions = tableXlsxSavingOptions)


[<RequireQualifiedAccess>]
module Table =
    
    type Internal_EmptyRecord =
        { EmptyText: string }

    let empty = 
        lazy
            Table.OfRecords<Internal_EmptyRecord> []

    let mapFrame mapping (table: Table) =
        table.MapFrame mapping

    let fillEmptyUp (table: Table) =
        table 
        |> mapFrame Frame.fillEmptyUp

    let fillEmptyUpForColumns columnKeys (table: Table) =
        table 
        |> mapFrame (Frame.fillEmptyUpForColumns columnKeys)

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
                    | Some v -> true
                    | None -> false
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

    /// no allow subsets
    /// Frame.indexRowsOrdinally
    let private concat_RefreshRowKeys_raw (tables: AtLeastOneList<Table>) =
        tables.Head
        |> mapFrame(fun _ ->
            tables
            |> AtLeastOneList.map (fun table -> table.AsFrame)
            |> Frame.Concat_RefreshRowKeys
        )

    let concat_RefreshRowKeys_ForceSameHeaders (tables: AtLeastOneList<Table>) =
        let allHeaders_List =
            tables.AsList
            |> List.collect(fun m -> List.ofSeq m.Headers)
            |> List.distinct

        let allHeaders_Indexed = 
            allHeaders_List 
            |> List.indexed


        let allHeaders = Set.ofList allHeaders_List


        let newTables =
            tables
            |> AtLeastOneList.map(fun tb ->
                let headers = tb.Headers |> Set.ofSeq
                let diff = Set.difference allHeaders headers 
                match diff.IsEmpty with 
                | true -> tb
                | false -> 
                    let diff = 
                        diff
                        |> Set.toList
                        |> List.map(fun m -> 
                            allHeaders_Indexed
                            |> List.find(fun (_, n) -> m = n)
                        )
                        |> List.sortBy fst
                        |> List.map snd

                    tb.AddColumns_SingtonValue(diff, "")
            )

        concat_RefreshRowKeys_raw newTables

    /// allow header subsets
    let concat_RefreshRowKeys (tables: AtLeastOneList<Table>) =
        match tables with 
        | AtLeastOneList.One table -> table
        | AtLeastOneList.Many tables ->
            let headerLists =
                tables.AsList
                |> List.map (fun m ->
                    m.Headers
                    |> Set.ofSeq
                )

            headerLists
            |> List.sortBy(fun m -> m.Count)
            |> List.iter(fun m ->
                let headers_length_biggerThan_this =
                    headerLists
                    |> List.filter(fun m2 -> m2.Count >= m.Count)

                headers_length_biggerThan_this
                |> List.iter(fun (headers_length_biggerThan_this) ->
                    match m.IsSubsetOf headers_length_biggerThan_this with 
                    | true -> ()
                    | false -> 
                        failwithf "headers1 %A <> headers2 %A when concating table" (Set.toList m) (Set.toList headers_length_biggerThan_this)
                )
            )

            concat_RefreshRowKeys_ForceSameHeaders tables.AsAtLeastOneList
            

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
            |> concat_RefreshRowKeys_raw
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




        

