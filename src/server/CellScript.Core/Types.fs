namespace CellScript.Core
#nowarn "0104"
open Deedle
open Extensions
open System
open OfficeOpenXml
open System.IO
open System.Collections.Generic
open Shrimp.FSharp.Plus
open Constrants

[<AutoOpen>]
module Types = 
    
    
    type ICellValue =
        abstract member Convertible: IConvertible




    let internal fixRawContent_AllowCellValue (content: obj) =
        match content with
        | :? ICellValue as v -> box v
        | _ -> fixRawContent content



    let private fixTableName (tableName: string) =
        tableName
            .Replace(' ','_')


    let private userRange_maxColumnWidth =
        lazy
            config.Value.GetInt("CellScript.Core.UserRangeMaxColumnIndex")

    type ValidTableName (v: string) =
        inherit POCOBaseV<StringIC>(StringIC(fixTableName v))

        member x.OriginName = v

        member x.StringIC = x.VV.Value

        member x.Name = x.StringIC.Value

        static member Convert(tableName: string) = ValidTableName(tableName).Name



    let internal readContentAsConvertible (content: obj) =
        match content with 
        | null -> "" :> System.IConvertible
        | :? IConvertible as convertible -> convertible 
        | :? ICellValue as v -> v.Convertible 
        | _ -> failwithf "type of cell value %A should either be ICellValue or IConvertible" (content.GetType())


    type ExcelPackageWithXlsxFile = private ExcelPackageWithXlsxFile of XlsxFile * ExcelPackage
    with 
        
        member x.XlsxFile = 
            let (ExcelPackageWithXlsxFile (v, _))  = x
            v

        member x.ExcelPackage = 
            let (ExcelPackageWithXlsxFile (_, v))  = x
            v

        interface System.IDisposable with 
            member x.Dispose() = (x.ExcelPackage :> IDisposable).Dispose()
                

        static member Create(xlsxFile: XlsxFile) =
            let package = new ExcelPackage(FileInfo xlsxFile.Path)
            ExcelPackageWithXlsxFile(xlsxFile, package)

    let internal DefaultTableStyle() = Table.TableStyles.Medium2


    type SheetReference =
        { WorkbookPath: string 
          SheetName: string }

    type DimensionEnum =
        | FixedDimension = 0
        /// ///<summary>
        ///            Dimension address for the worksheet for cells with a value different than null. 
        ///            Top left cell to Bottom right.
        ///            If the worksheet has no cells, null is returned        
        ///            </summary>
        | DimensionByValue = 1
    
    type RangeGettingOptions =
        | RangeIndexerCase of string * includeHided: bool
        /// MaxColumnIndex: cfg(CellScript.Core.UserRangeMaxColumnIndex)
        | UserRangeCase of includeHided: bool * dimension: DimensionEnum
        | UserRange_SkipRowsCase of int * includeHided: bool * dimension: DimensionEnum
        | TableNameCase         of string * includeHided: bool
        | TableNameExprCase         of TextSelector * includeHided: bool
        | TableLeftTopCase      of TextSelector * includeHided: bool
    with 
        member x.IncludeHided =
            match x with 
            | RangeIndexerCase(_, b)
            | UserRangeCase(b, _)
            | UserRange_SkipRowsCase(_, b, _)
            | TableNameCase(_, b)
            | TableNameExprCase(_, b)
            | TableLeftTopCase(_, b) -> b

        member x.SetIncludeHided(b) =
            match x with 
            | RangeIndexerCase(v, _) -> 
                (v, b)
                |> RangeIndexerCase 

            | UserRangeCase(_, removeEndingEmptys) -> UserRangeCase(b, removeEndingEmptys)
            | UserRange_SkipRowsCase(v, _, removeEndingEmptys) ->
                (v, b, removeEndingEmptys)
                |> UserRange_SkipRowsCase

            | TableNameCase(v, _) ->
                (v, b)
                |> TableNameCase


            | TableNameExprCase(v, _) ->
                (v, b)
                |> TableNameExprCase


            | TableLeftTopCase(v, _) ->
                (v, b)
                |> TableLeftTopCase

        static member UserRange = 
            RangeGettingOptions.UserRangeCase(
                includeHided = false,
                dimension = DimensionEnum.DimensionByValue
            )

        static member RangeIndexer(indexer, ?includeHided) = 
            RangeGettingOptions.RangeIndexerCase(indexer, defaultArg includeHided false)

        static member TableName(tableName, ?includeHided) = 
            RangeGettingOptions.TableNameCase(tableName, defaultArg includeHided false)

        static member TableNameExpr(tableNameTextSelector, ?includeHided) = 
            RangeGettingOptions.TableNameExprCase(tableNameTextSelector, defaultArg includeHided false)

        static member TableLeftTop(tableName, ?includeHided) = 
            RangeGettingOptions.TableLeftTopCase(tableName, defaultArg includeHided false)

        static member UserRange_SkipRows(skipRowsCount, ?includeHided, ?removeEndingEmptys) = 
            RangeGettingOptions.UserRange_SkipRowsCase(
                skipRowsCount,
                defaultArg includeHided false,
                DimensionEnum.DimensionByValue
            )


    [<RequireQualifiedAccess>]
    type ColumnAutofitOptions =
        | None
        | AutoFit of minimumSize: float * maximumSize: float
    with    
        /// ColumnAutofitOptions.AutoFit(10., 50.)
        static member DefaultValue = ColumnAutofitOptions.AutoFit(10., 50.)

    
    [<RequireQualifiedAccess>]
    type CellSavingFormat =
        | AutomaticNumberic 
        | AllText of excludingCols: StringIC list
        | KeepOrigin 
        | ODBC

    [<RequireQualifiedAccess>]
    module FormulaText =
        let Create(text: string) =
            sprintf "Formula(%s)" text



    let private (|FormulaText|_|) (text: string) =
        match text with 
        | String.TrimStartIC "Formula(" v -> Some (v.TrimEnding(")"))
        | _ -> None


    type SheetContentsEmptyException(error: string) =
        inherit Exception(error)

    type OfficeOpenXml.Table.ExcelTable with 
        member x.DeleteRowsBack(rowsCount) =
            let rows = x.Address.Rows
            x.DeleteRow(rows-rowsCount-1, rowsCount)

        member x.DeleteColumnsBack(columnsCount) =
            let columns = x.Address.Columns
            x.Columns.Delete(columns-columnsCount-1, columnsCount)

    type ExcelWorksheet with 
        member sheet.FixedDimension =
            let dimension = sheet.Dimension
            let ending = 
                let addr = dimension.End
                ExcelCellAddress(addr.Row, min userRange_maxColumnWidth.Value addr.Column)

            let start = dimension.Start

            ExcelAddress(
                start.Row,
                start.Column,
                ending.Row,
                ending.Column
            )

        member x.GetDimensionByEnum(dimensionEnum: DimensionEnum) =
            match dimensionEnum with 
            | DimensionEnum.FixedDimension -> x.FixedDimension
            | DimensionEnum.DimensionByValue -> x.FixedDimension
                //let dimension = x.DimensionByValue
                //let ending = 
                //    let addr = dimension.End
                //    ExcelCellAddress(addr.Row, min userRange_maxColumnWidth.Value addr.Column)

                //let start = dimension.Start

                //ExcelAddress(
                //    start.Row,
                //    start.Column,
                //    ending.Row,
                //    ending.Column
                //)
                

    type TableNameNotFoundException(tableName, allTableNames: string list) =
        inherit Exception(sprintf "Cannot found any table named %s, avaliable tables are %A" tableName allTableNames)

    /// Default ColumnPastingOptions.Directly
    type ColumnPastingOptions =
        | Directly = 0
        | ByColumnName = 1

    type VisibleExcelWorksheet = private VisibleExcelWorksheet of ExcelWorksheet
    with 
        member x.Value =
            let (VisibleExcelWorksheet v) = x
            v

        member x.Name = x.Value.Name

        member x.TryGetRange options = 
            let sheet = x.Value

            match options with
            | RangeIndexerCase (indexer, _) ->
                sheet.Cells.[indexer] :> ExcelRangeBase
                |> Result.Ok

            | UserRangeCase(_, dimension) ->
                let dimension = 
                    sheet.GetDimensionByEnum(dimension)

                let indexer = dimension.Start.Address + ":" + dimension.Address
                sheet.Cells.[indexer] :> ExcelRangeBase
                |> Result.Ok

            | UserRange_SkipRowsCase(skipRows, _, dimension) ->
                let dimension = 
                    sheet.GetDimensionByEnum(dimension)

                let start = 
                    let start = dimension.Start
                    ExcelCellAddress(start.Row + skipRows, start.Column)

                let indexer = start.Address + ":" + dimension.End.Address
                sheet.Cells.[indexer] :> ExcelRangeBase
                |> Result.Ok

            | TableNameCase(tbName, _) ->
                let findedTable =
                    x.Value.Tables
                    |> Seq.tryFind(fun m -> StringIC m.Name = StringIC tbName)

                match findedTable with 
                | Some table -> table.Range |> Result.Ok
                | None ->   
                    let allTableNames =
                        x.Value.Tables
                        |> List.ofSeq
                        |> List.map(fun m -> m.Name)

                    (TableNameNotFoundException(tbName, allTableNames) :> System.Exception)
                    |> Result.Error

            | TableNameExprCase(tbName, _) ->
                let findedTable =
                    x.Value.Tables
                    |> Seq.tryFind(fun m -> 
                        tbName.Predicate(m.Name)
                        //StringIC m.Name = StringIC tbName
                    )

                match findedTable with 
                | Some table -> table.Range |> Result.Ok
                | None ->   
                    let allTableNames =
                        x.Value.Tables
                        |> List.ofSeq
                        |> List.map(fun m -> m.Name)

                    (TableNameNotFoundException(tbName.MethodLiteralText, allTableNames) :> System.Exception)
                    |> Result.Error

            | TableLeftTopCase(expr, _) ->
                let userRange = x.TryGetRange (RangeGettingOptions.UserRange)
                userRange
                |> Result.map(fun userRange ->
                    let r =
                        userRange
                        |> Seq.tryPick(fun range ->
                            match expr.Predicate range.Text with 
                            | true ->  
                                let finedTable = 
                                    x.Value.Tables
                                    |> List.ofSeq
                                    |> List.tryFind(fun tb ->
                                        tb.Address.Start.Address = range.Start.Address
                                    )
                                finedTable
                            | false -> None
                        )
                        
                    match r with 
                    | None -> failwithf "Cannot finded table by left top indexer %s" (expr.MethodLiteralText)
                    | Some r -> r.Range
                )


        member x.GetRange options = x.TryGetRange(options) |> Result.getOrRaise


        member x.LoadFromArrays(array2D: IConvertible [, ], ?addr, ?includingFormula) =
            
            let worksheet = x.Value
            let addr = defaultArg addr "A1"

            let range = worksheet.Cells.[addr].LoadFromArray2D(array2D) 
                
            match defaultArg includingFormula false with 
            | false -> ()
            | true ->
                let range = worksheet.Cells.[range.Start.Address]
                array2D
                |> Array2D.toLists
                |> List.iteri(fun rowNum row ->
                    row
                    |> List.iteri(fun colNum v ->
                        match v with 
                        | :? string as v -> 
                            match v with 
                            | FormulaText v -> 
                                let range = range.Offset(rowNum, colNum)
                                //range.Value <- v
                                range.Formula <- v
                            | _ -> ()
                        | _ -> ()
                    )
                )

      

        member x.LoadFromArraysAsTable(datas: IConvertible [, ], ?columnAutofitOptions, ?tableName: string, ?tableStyle: Table.TableStyles, ?addr, ?includingFormula, ?allowRerangeTable, ?columnPastingOptions) =
            //let allowRerangeTable = None
            let __checkDataValid =
                match Array2D.length1 datas, Array2D.length2 datas with 
                | BiggerThan 1, BiggerThan 0 -> ()
                | l1, l2 -> failwithf "Invalid table data length %A" (l1, l2)

            let fixHeaders (datas: IConvertible [,]) =
                let lists = Array2D.toLists datas
                let headers = lists.[0]
                let contents = lists.[1..]

                let headers = 
                    let headers =
                        headers
                        |> List.map(fun m ->
                            match m with 
                            | null -> ""
                            | _ -> m.ToString()
                        )


                    let automaticColumnHeaders =
                        headers
                        |> List.choose(fun m -> 
                            match m with
                            | String.TrimStartIC "Column" i ->
                                match System.Int32.TryParse i with 
                                | true, i -> Some i
                                | _ -> None

                            | _ -> None
                        )
                        |> HashSet

                    let generateId() =
                        [1..1000]
                        |> List.find(fun m ->
                            match automaticColumnHeaders.Contains m with
                            | true -> false
                            | false -> 
                                automaticColumnHeaders.Add m |> ignore
                                true
                        )


                    headers
                    |> List.map(fun m ->
                        match m.Trim() with 
                        | "" -> ("Column" + generateId().ToString()) :> IConvertible
                        | _ -> m :> IConvertible
                    )
                
                let datas =
                    headers :: contents
                    |> array2D

                headers, datas

            let headers, datas = fixHeaders datas

            let headers = headers |> List.map (fun header ->
                (ConvertibleUnion.Convert header).Text
                |> StringIC
            )

            let worksheet = x.Value
            let allowRerangeTable = defaultArg allowRerangeTable false

            let tableName = defaultArg tableName DefaultTableName

            let findedTable =
                worksheet.Tables
                |> Seq.tryFind(fun m -> StringIC m.Name = StringIC tableName)

            //let avaliableTableNames = 
            //    worksheet.Tables
            //    |> List.ofSeq
            //    |> List.map(fun m -> m.Name)

            let columnPastingOptions = defaultArg columnPastingOptions ColumnPastingOptions.Directly


            let addr = 
                match findedTable, allowRerangeTable with  
                | Some table, true -> table.Address.Start.Address
                | _ -> defaultArg addr "A1"

            let range = worksheet.Cells.[addr]

            let tab = 
                match findedTable, allowRerangeTable with 
                | Some findedTable, true -> 
                    let array2D_rows    = Array2D.length1 datas
                    let array2D_columns = Array2D.length2 datas

                    let addr = findedTable.Address
                    
                    let rows    = addr.Rows
                    let columns = addr.Columns

                    let row_substract = rows - array2D_rows

                    match row_substract with 
                    | 0 -> ()
                    | BiggerThan 0 ->
                        findedTable.DeleteRowsBack(abs row_substract)
                        |> ignore

                    | SmallerThan 0 -> 
                        let row = findedTable.AddRow(abs row_substract)

                        let keepHeights = 
                            let beforeRow = row.Offset(-1, 0, 1, 1)
                            let beforeRowHeight = beforeRow.EntireRow.Height
                            let rowStart = row.Start
                            let rowEnd = row.End

                            [rowStart.Row .. rowEnd.Row]
                            |> List.iter(fun row ->
                                worksheet.Row(row).Height <- beforeRowHeight
                            )

                        row
                        |> ignore
                    
                    | _ -> failwith "Invalid token"

                    match columnPastingOptions with 
                    | ColumnPastingOptions.Directly ->
                        
                        let column_substract = columns - array2D_columns
                        match column_substract with 
                        | 0 -> ()
                        | BiggerThan 0 ->
                            findedTable.DeleteColumnsBack(abs column_substract)
                            |> ignore

                        | SmallerThan 0 -> 
                            findedTable.Columns.Add(abs column_substract)
                            |> ignore
                    
                        | _ -> failwith "Invalid token"

                        let __checkTableValid =
                            let addr = findedTable.Address
                            match addr.Rows = array2D_rows, addr.Columns = array2D_columns with 
                            | true, true -> ()
                            | _ -> failwithf "Invalid token, new table addr %A is not consistent to %A" addr.Address (array2D_columns, array2D_rows)

                        
                        worksheet.Cells.[addr.Address].LoadFromArray2D(datas)
                        |> ignore

                        findedTable

                    | ColumnPastingOptions.ByColumnName ->
                        let __checkTableValid =
                            let addr = findedTable.Address
                            match addr.Rows = array2D_rows(*, addr.Columns = array2D_columns*) with 
                            | true(*, true*) -> ()
                            | _ -> failwithf "Invalid token, new table addr %A is not consistent to %A" addr.Address (array2D_columns, array2D_rows)

                        
                        //let datas = datas.[1.., *]

                        let addr = 
                            findedTable.Address.Address
                            |> ComparableExcelAddress.OfAddress
                        //let originWidth  = worksheet.Cells.[addr.Address].EntireColumn.Width
                        //let originWidth  = worksheet.Cells.["A1"].EntireColumn.Width
                 
                        let columns = 
                            let headers = 
                                let range = 
                                    let columnCount = findedTable.Columns.Count
                                    findedTable.Range.Offset(0, 0, 1, columnCount)
                            
                                range
                                |> List.ofSeq
                                |> List.map(fun m -> m.Text)

                            let columns = 
                                findedTable.Columns
                                |> List.ofSeq

                            (headers, columns)
                            ||> List.map2(fun header column ->
                                {|
                                    Name = header
                                    Column = column
                                |}
                            )

                        let headers = 
                            match columns with 
                            | [column] -> [StringIC column.Name]
                            | _ -> headers

                        let columnNames =
                            columns
                            |> List.map(fun m -> StringIC m.Name)

                        columns
                        |> List.iteri(fun columnID column ->
                            let columnName = column.Name
                            let finedHeader =
                                headers
                                |> List.tryFindIndex(fun header -> header = StringIC columnName)

                            match finedHeader with 
                            | None -> ()
                            | Some finedHeader ->
                                let datas = datas.[1.., finedHeader]
                                let addr2 = addr.Offset(1, columnID, datas.Length-1, 0)
                                let cells = worksheet.Cells.[addr2.Address]
                                match cells.Formula with 
                                | "" ->
                                    cells.LoadFromCollection(datas)
                                    |> ignore

                                | _ -> 
                                    match cells.Columns, cells.Rows with 
                                    | 1, EqualTo datas.Length ->
                                        let range0 = worksheet.Cells.[cells.Start.Address]
                                        datas
                                        |> Array.iteri(fun i data ->
                                            match data with 
                                            | :? string as data ->
                                                match data with 
                                                | FormulaText formula ->
                                                    let range = range0.Offset(i, 0)
                                                    range.Formula <- formula

                                                | _ -> ()
                                            | _ -> ()
                                        )
                                    | _ -> ()
                        )

                        let addOtherColumn =
                            headers
                            |> List.indexed
                            |> List.filter(fun (i, header) ->
                                List.contains header columnNames
                                |> not
                            )
                            |> List.iter(fun (i, header) ->
                                let newColumn = findedTable.Columns.Add(1)
                                let datas = datas.[0.., i]
                                newColumn.LoadFromCollection(datas)
                                |> ignore
                            )

           

                        //worksheet.Cells.[addr.Address].LoadFromArray2D(datas)
                        //|> ignore


                        findedTable

                | Some tb, false ->
                    failwithf "Duplicate table name %s" tb.Name

                | None, false
                | None, true ->
                    let range = worksheet.Cells.[addr].LoadFromArray2D(datas)
                    try
                        worksheet.Tables.Add(ExcelAddress range.Address, tableName)

                    with ex ->
                        let message = 
                            let avaliableTableNames = 
                                worksheet.Tables
                                |> List.ofSeq
                                |> List.map(fun m -> m.Name)

                            sprintf "%s\nWhen adding table ‘%s’\nAvaliable table names are %A" ex.Message tableName avaliableTableNames
                        let ex = new System.Exception(message)
                        raise ex
                
            match defaultArg includingFormula false with 
            | false -> ()
            | true ->
                match columnPastingOptions with 
                | ColumnPastingOptions.Directly ->
                    let range = worksheet.Cells.[range.Start.Address]
                    datas
                    |> Array2D.toLists
                    |> List.iteri(fun rowNum row ->
                        row
                        |> List.iteri(fun colNum v ->
                            match v with 
                            | :? string as v -> 
                                match v with 
                                | FormulaText v -> 
                                    let offsetedRange = range.Offset(rowNum, colNum)
                                    offsetedRange.Formula <- v
                                | _ -> ()
                            | _ -> ()
                        )
                    )

                | ColumnPastingOptions.ByColumnName -> ()


            tab.TableStyle <- 
                defaultArg tableStyle <| DefaultTableStyle()

            match defaultArg columnAutofitOptions ColumnAutofitOptions.DefaultValue with 
            | ColumnAutofitOptions.AutoFit(minimumSize, maximumSize) ->
                range.AutoFitColumns(minimumSize, maximumSize)

            | ColumnAutofitOptions.None -> ()

        static member Create(excelworksheet: ExcelWorksheet) =
            match excelworksheet with 
            | null -> failwithf "Cannot create VisibleExcelWorksheet: excelworksheet is null"
            | _ -> ()

            match excelworksheet.Hidden with 
            | eWorkSheetHidden.Visible -> ()
            | _ -> failwithf "Cannot create VisibleExcelWorksheet: excelworksheet %s is hidden" excelworksheet.Name

            VisibleExcelWorksheet excelworksheet


    /// both visible and having contents
    type ValidExcelWorksheet private (visibleExcelWorksheet: VisibleExcelWorksheet) =

        member x.Value = visibleExcelWorksheet.Value

        member x.Name = x.Value.Name

        member x.VisibleExcelWorksheet = visibleExcelWorksheet

        member x.GetRange rangeGettingOptions = visibleExcelWorksheet.GetRange(rangeGettingOptions)

        member sheet.TryReadDatasWithUserState(rangeGettingOptions: RangeGettingOptions, fUserState) =
            let includeHided = rangeGettingOptions.IncludeHided
            let sheet = sheet.VisibleExcelWorksheet
            let mergedCellAddrs = sheet.Value.TryGetMergeCellAddrs()


            let range = sheet.TryGetRange(rangeGettingOptions)
            range
            |> Result.map(fun range ->
                let rowStart = range.Start.Row
                let rowEnd = range.End.Row
                let columnStart = range.Start.Column
                let columnEnd = range.End.Column
    
                let content = range.ReadDatasWithUserState_TrackMergeRange(fUserState, includeHided, mergedCellAddrs)
                let content, reducedNums = 
                    match rangeGettingOptions with 
                    | RangeGettingOptions.UserRangeCase(_, dimension) 
                    | RangeGettingOptions.UserRange_SkipRowsCase(_, _, dimension) ->
                        match dimension with 
                        | DimensionEnum.FixedDimension -> content, None
                        | DimensionEnum.DimensionByValue ->
                            let l1 = Array2D.length1 content
                            let l2 = Array2D.length2 content
                            let l2_down = 
                                [0..l2-1]
                                |> List.rev
                                |> List.takeWhile(fun l2 ->
                                    [0..l1-1]
                                    |> List.forall(fun l1 ->
                                        content.[l1, l2].Content.IsStringEmpty
                                    )
                                )
                                |> List.tryLast
                                |> Option.defaultValue (l2-1)

                            let l1_down = 
                                [0..l1-1]
                                |> List.rev
                                |> List.takeWhile(fun l1 ->
                                    [0..l2-1]
                                    |> List.forall(fun l2 ->
                                        content.[l1, l2].Content.IsStringEmpty
                                    )
                                )
                                |> List.tryLast
                                |> Option.defaultValue (l1-1)

                            let newContent = content.[0..l1_down, 0..l2_down]
                            newContent, Some (l1-1-l1_down, l2-1-l2_down)

                    | _ -> content, None

                let rowEnd, columnEnd = 
                    match reducedNums with 
                    | None -> rowEnd, columnEnd
                    | Some (reducedRow, reducedCol) ->
                        rowEnd - reducedRow, columnEnd - reducedCol

                {|
                    RowStart = rowStart
                    RowEnd = rowEnd
                    ColumnStart = columnStart
                    ColumnEnd = columnEnd
                    Content = content
                |}
            )



        member sheet.ReadDatasWithUserState(rangeGettingOptions, fUserState) =
            sheet.TryReadDatasWithUserState(rangeGettingOptions, fUserState)
            |> Result.getOrFail

        member sheet.TryReadDatas(rangeGettingOptions) =
            let r = sheet.TryReadDatasWithUserState(rangeGettingOptions, ignore)
            r
            |> Result.map(fun r ->
                {| r with 
                    Content = 
                        r.Content
                        |> Array2D.map(fun m -> m.Content)
                    
                |}
            )


        member sheet.ReadDatas(rangeGettingOptions) =
            let r = sheet.ReadDatasWithUserState(rangeGettingOptions, ignore)
            {| r with 
                Content = 
                    r.Content
                    |> Array2D.map(fun m -> m.Content)
                
            |}

        static member TryCreate(visibleExcelWorksheet: VisibleExcelWorksheet) =
            match visibleExcelWorksheet.Value.Dimension with
            | null -> Result.Error (sprintf "Cannot create VisibleExcelWorksheet: Contents in sheet %s is empty" visibleExcelWorksheet.Value.Name)
            | _ -> 
                Result.Ok (ValidExcelWorksheet visibleExcelWorksheet)

        static member Create(visibleExcelWorksheet: VisibleExcelWorksheet) =
            ValidExcelWorksheet.TryCreate(visibleExcelWorksheet)
            |> Result.getOrFail

        static member Create (excelworksheet: ExcelWorksheet) =
            ValidExcelWorksheet.Create(VisibleExcelWorksheet.Create(excelworksheet))
        


    [<RequireQualifiedAccess>]
    type SheetGettingOptions =
        | SheetName of StringIC
        | SheetIndex of int
        | SheetNameOrSheetIndex of sheetName: StringIC * index: int
    with 
        /// SheetGettingOptions.SheetNameOrSheetIndex (StringIC SHEET1, 0)
        static member DefaultValue = 
            SheetGettingOptions.SheetNameOrSheetIndex (StringIC SHEET1, 0)
            //SheetGettingOptions.SheetIndex 0



    type ExcelPackage with

        member excelPackage.GetVisibleWorksheets() =
            excelPackage.Workbook.Worksheets
            |> Seq.filter(fun m -> m.Hidden = eWorkSheetHidden.Visible)
            |> Seq.map VisibleExcelWorksheet

        member excelPackage.TryGetVisibleSheetByIndex(index) =
            let worksheet =
                excelPackage.Workbook.Worksheets
                |> Seq.filter(fun m -> m.Hidden = eWorkSheetHidden.Visible)
                |> Seq.tryItem index

            match worksheet with 
            | Some worksheet -> Result.Ok worksheet
            | None -> 
                sprintf "Cannot get visible worksheet %d from %s, please check xlsx file" index excelPackage.File.FullName
                |> Result.Error

        member excelPackage.GetVisibleSheetByIndex(index) =
            excelPackage.TryGetVisibleSheetByIndex(index)
            |> Result.getOrFail

        member excelPackage.GetVisibleWorksheet (options) =
            let excelworksheet = 
                match options with 
                | SheetGettingOptions.SheetName sheetName -> 
                    let sheet = excelPackage.Workbook.Worksheets.[sheetName.Value]
                    match sheet with 
                    | null -> failwithf "No sheet named %s was found in %A" sheetName.Value (excelPackage.File)
                    | _ -> sheet

                | SheetGettingOptions.SheetIndex index -> excelPackage.GetVisibleSheetByIndex index
                | SheetGettingOptions.SheetNameOrSheetIndex (sheetName, index) ->
                    let worksheets = excelPackage.Workbook.Worksheets
                    match Seq.tryFind (fun (worksheet: ExcelWorksheet) -> StringIC worksheet.Name = sheetName && worksheet.Hidden = eWorkSheetHidden.Visible) worksheets with 
                    | Some worksheet -> 
                        let sheet_byName = worksheet
                        match sheet_byName.View.TabSelected with 
                        | true -> sheet_byName
                        | false -> 
                            let sheet_byIndex = excelPackage.GetVisibleSheetByIndex index
                            match sheet_byIndex.View.TabSelected with 
                            | true -> sheet_byIndex
                            | false -> sheet_byName

                    | None -> excelPackage.GetVisibleSheetByIndex index

  

            VisibleExcelWorksheet.Create excelworksheet

        member excelPackage.TryGetValidWorksheet (options) =

            let r = 
                try
                    excelPackage.GetVisibleWorksheet(options)
                    |> ValidExcelWorksheet.TryCreate
                with ex ->
                    Result.Error (ex.Message)

            match r with 
            | Result.Ok r -> Result.Ok r
            | Result.Error _ -> 
                match options with 
                | SheetGettingOptions.SheetName _
                | SheetGettingOptions.SheetIndex _ -> r
                | SheetGettingOptions.SheetNameOrSheetIndex (sheetName, index) ->
                    excelPackage.GetVisibleSheetByIndex index
                    |> VisibleExcelWorksheet.Create
                    |> ValidExcelWorksheet.TryCreate

        member excelPackage.GetValidWorksheet (options) =
            excelPackage.GetVisibleWorksheet(options)
            |> ValidExcelWorksheet.Create

        member excelPackage.GetValidWorksheets_Seq() =
            excelPackage.GetVisibleWorksheets()
            |> Seq.choose (ValidExcelWorksheet.TryCreate >> Result.toOption)

        member excelPackage.GetValidWorksheets() =
            excelPackage.GetValidWorksheets_Seq()
            |> List.ofSeq

        member excelPackage.GetOrAddWorksheet(sheetName: string) =
            let worksheet = 
                let worksheets = excelPackage.Workbook.Worksheets
                worksheets
                |> Seq.tryFind(fun sheet -> StringIC sheet.Name = StringIC sheetName)
                |> function
                    | Some sheet -> VisibleExcelWorksheet.Create sheet 
                    | None ->
                        excelPackage.Workbook.Worksheets.Add(sheetName)
                        |> VisibleExcelWorksheet.Create

            worksheet

    type ExcelRangeContactInfo =
        { ColumnFirst: int
          RowFirst: int
          ColumnLast: int
          RowLast: int
          XlsxFile: XlsxFile
          SheetName: string
          Content: ConvertibleUnion[,] }
    with 
        member x.WorkbookPath = x.XlsxFile.Path

        member xlRef.CellAddress = ExcelAddress(xlRef.RowFirst, xlRef.ColumnFirst, xlRef.RowLast, xlRef.ColumnLast)
    
        override xlRef.ToString() = sprintf "%s %s (%d, %d, %d, %d)" xlRef.WorkbookPath xlRef.SheetName xlRef.RowFirst xlRef.ColumnFirst xlRef.RowLast xlRef.ColumnLast
    
    


    [<RequireQualifiedAccess>]
    module ExcelRangeContactInfo =
    
        let cellAddress (xlRef: ExcelRangeContactInfo) = xlRef.CellAddress
    
        let readFromExcelPackages (rangeGettingOptions: RangeGettingOptions) (sheetGettingArgs: SheetGettingOptions) (excelPackage: ExcelPackageWithXlsxFile) =
            let xlsxFile = excelPackage.XlsxFile
            let excelPackage = excelPackage.ExcelPackage
            let sheet = excelPackage.GetValidWorksheet sheetGettingArgs
    
            let datas = sheet.ReadDatas(rangeGettingOptions) 
    
            { ColumnFirst = datas.ColumnStart
              RowFirst = datas.RowStart
              ColumnLast = datas.ColumnEnd
              RowLast = datas.RowEnd
              XlsxFile = xlsxFile
              SheetName = sheet.Name
              Content = datas.Content }   

        let readFromFile (rangeGettingArg: RangeGettingOptions) (sheetGettingArgs: SheetGettingOptions) (xlsxFile: XlsxFile) =
            use excelPackage = ExcelPackageWithXlsxFile.Create xlsxFile
            readFromExcelPackages rangeGettingArg sheetGettingArgs excelPackage
    

    type ExcelPackage with 
        member excelPackage.GetWorkSheet(xlRef: ExcelRangeContactInfo) =
            excelPackage.Workbook.Worksheets
            |> Seq.find (fun worksheet -> worksheet.Name = xlRef.SheetName)


        member excelPackage.GetExcelRange(xlRef: ExcelRangeContactInfo) =
            let workbookSheet = excelPackage.GetWorkSheet xlRef
            workbookSheet.Cells.[xlRef.RowFirst, xlRef.ColumnFirst, xlRef.RowLast, xlRef.ColumnLast]

    
    type SerializableExcelReference = 
        { ColumnFirst: int
          RowFirst: int
          ColumnLast: int
          RowLast: int
          XlsxFile: XlsxFile
          SheetName: string }
    with 
        member x.WorkbookPath = x.XlsxFile.Path

        member xlRef.CellAddress = ExcelAddress(xlRef.RowFirst, xlRef.ColumnFirst, xlRef.RowLast, xlRef.ColumnLast)
    
    [<RequireQualifiedAccess>]
    module SerializableExcelReference =
        let ofExcelRangeContactInfo (xlRef: ExcelRangeContactInfo) =
            { ColumnFirst = xlRef.ColumnFirst
              RowFirst = xlRef.RowFirst
              ColumnLast = xlRef.ColumnLast
              RowLast = xlRef.RowLast
              XlsxFile = xlRef.XlsxFile
              SheetName = xlRef.SheetName }
    
        let toExcelRangeContactInfo content (xlRef: SerializableExcelReference) : ExcelRangeContactInfo =
            { ColumnFirst = xlRef.ColumnFirst
              RowFirst = xlRef.RowFirst
              ColumnLast = xlRef.ColumnLast
              RowLast = xlRef.RowLast
              XlsxFile = xlRef.XlsxFile
              SheetName = xlRef.SheetName
              Content = content }
    
        let cellAddress xlRef =
            ExcelAddress(xlRef.RowFirst, xlRef.ColumnFirst, xlRef.RowLast, xlRef.ColumnLast)
    
