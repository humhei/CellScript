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


    type ICellValue =
        abstract member Convertible: IConvertible
    



    let internal readContentAsConvertible (content: obj) =
        match content with 
        | null -> null 
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
    
    type RangeGettingOptions =
        | RangeIndexer of string
        /// MaxColumnIndex: cfg(CellScript.Core.UserRangeMaxColumnIndex)
        | UserRange
        | UserRange_SkipRows of int

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
            let ending = 
                let addr = sheet.Dimension.End
                ExcelCellAddress(addr.Row, min userRange_maxColumnWidth.Value addr.Column)

            let start = sheet.Dimension.Start

            ExcelAddress(
                start.Row,
                start.Column,
                ending.Row,
                ending.Column
            )


    type VisibleExcelWorksheet = private VisibleExcelWorksheet of ExcelWorksheet
    with 
        member x.Value =
            let (VisibleExcelWorksheet v) = x
            v

        member x.Name = x.Value.Name

        member x.GetRange options = 
            let sheet = x.Value

            match options with
            | RangeIndexer indexer ->
                sheet.Cells.[indexer]

            | UserRange ->
                let dimension = sheet.FixedDimension

                let indexer = dimension.Start.Address + ":" + dimension.Address
                sheet.Cells.[indexer]

            | UserRange_SkipRows skipRows ->
                let dimension = sheet.FixedDimension
                let start = 
                    let start = dimension.Start
                    ExcelCellAddress(start.Row + skipRows, start.Column)

                let indexer = start.Address + ":" + dimension.End.Address
                sheet.Cells.[indexer]

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

      

        member x.LoadFromArraysAsTable(datas: IConvertible [, ], ?columnAutofitOptions, ?tableName: string, ?tableStyle: Table.TableStyles, ?addr, ?includingFormula, ?allowRerangeTable) =
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

                headers :: contents
                |> array2D

            let datas = fixHeaders datas

            let worksheet = x.Value
            let allowRerangeTable = defaultArg allowRerangeTable false

            let tableName = defaultArg tableName DefaultTableName

            let findedTable =
                worksheet.Tables
                |> Seq.tryFind(fun m -> StringIC m.Name = StringIC tableName)

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
                    let column_substract = columns - array2D_columns

                    match row_substract with 
                    | 0 -> ()
                    | BiggerThan 0 ->
                        findedTable.DeleteRowsBack(abs row_substract)
                        |> ignore

                    | SmallerThan 0 -> 
                        findedTable.AddRow(abs row_substract)
                        |> ignore
                    
                    | _ -> failwith "Invalid token"

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

                | Some tb, false ->
                    failwithf "Duplicate table name %s" tb.Name

                | None, false
                | None, true ->
                    let range = worksheet.Cells.[addr].LoadFromArray2D(datas)
                    worksheet.Tables.Add(ExcelAddress range.Address, tableName)

                
            match defaultArg includingFormula false with 
            | false -> ()
            | true ->
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
                                range.Offset(rowNum, colNum).Formula <- v
                            | _ -> ()
                        | _ -> ()
                    )
                )


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

        member sheet.ReadDatas(rangeGettingOptions) =
            let sheet = sheet.VisibleExcelWorksheet
            
            use range = sheet.GetRange(rangeGettingOptions)
    
            let rowStart = range.Start.Row
            let rowEnd = range.End.Row
            let columnStart = range.Start.Column
            let columnEnd = range.End.Column
    
            let content = range.ReadDatas()
                
    
            {|
                RowStart = rowStart
                RowEnd = rowEnd
                ColumnStart = columnStart
                ColumnEnd = columnEnd
                Content = content
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
            //SheetGettingOptions.SheetIndex 0
            SheetGettingOptions.SheetNameOrSheetIndex (StringIC SHEET1, 0)



    type ExcelPackage with

        member excelPackage.GetVisibleWorksheets() =
            excelPackage.Workbook.Worksheets
            |> Seq.filter(fun m -> m.Hidden = eWorkSheetHidden.Visible)
            |> Seq.map VisibleExcelWorksheet

        member excelPackage.GetVisibleSheetByIndex(index) =
            let worksheet =
                excelPackage.Workbook.Worksheets
                |> Seq.filter(fun m -> m.Hidden = eWorkSheetHidden.Visible)
                |> Seq.tryItem index

            match worksheet with 
            | Some worksheet -> worksheet
            | None -> failwithf "Cannot get visible worksheet %A from %s, please check xlsx file" index excelPackage.File.FullName


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
                    match Seq.tryFind (fun (worksheet: ExcelWorksheet) -> StringIC worksheet.Name = sheetName && worksheet.Hidden = eWorkSheetHidden.Visible) excelPackage.Workbook.Worksheets with 
                    | Some worksheet -> worksheet
                    | None -> excelPackage.GetVisibleSheetByIndex index


            VisibleExcelWorksheet.Create excelworksheet

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
    
