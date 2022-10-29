namespace CellScript.Core
#nowarn "0104"
open Deedle
open Extensions
open System
open OfficeOpenXml
open System.IO
open Shrimp.FSharp.Plus
open Constrants

[<AutoOpen>]
module Types = 

    let private fixTableName (tableName: string) =
        tableName
            .Replace(' ','_')

    type ValidTableName (v: string) =
        inherit POCOBaseV<StringIC>(StringIC(fixTableName v))

        member x.OriginName = v

        member x.StringIC = x.VV.Value

        member x.Name = x.StringIC.Value

        static member Convert(tableName: string) = ValidTableName(tableName).Name


    type ICellValue =
        abstract member Convertible: IConvertible
    

    let internal fixRawContent (content: obj) =
        match content with
        | :? ConvertibleUnion as v -> v.Value
        | null -> null
        | :? ExcelErrorValue -> null
        | :? string as text -> 
            match text.Trim() with 
            | "" -> null 
            | _ -> text :> IConvertible
        | :? IConvertible as v -> 
            match v with 
            | :? DBNull -> null
            | _ -> v
        | _ -> failwithf "Cannot convert to %A to iconvertible" (content.GetType())

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
        | UserRange

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


    type SheetContentsEmptyException(error: string) =
        inherit Exception(error)


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
                let indexer = sheet.Dimension.Start.Address + ":" + sheet.Dimension.End.Address
                sheet.Cells.[indexer]


        member x.LoadFromArraysAsTable(arrays: IConvertible [, ], ?columnAutofitOptions, ?tableName: string, ?tableStyle: Table.TableStyles, ?addr) =


            
            let worksheet = x.Value
            let addr = defaultArg addr "A1"

            let range = worksheet.Cells.[addr].LoadFromArray2D(arrays) 
                
            let tab = 
                let tableName =
                    match tableName with 
                    | Some tableName -> tableName
                    | None -> DefaultTableName

                worksheet.Tables.Add(ExcelAddress range.Address, tableName)

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
    
            let content =
                array2D
                    [ for i = rowStart to rowEnd do 
                        yield
                            [ for j = columnStart to columnEnd do yield fixRawContent range.[i,j].Value ]
                    ] 
                |> Array2D.map ConvertibleUnion.Convert
    
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
            ValidExcelWorksheet(visibleExcelWorksheet)

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

        member excelPackage.GetVisibleWorksheet (options) =
            let getVisibleSheetByIndex(index) =
                let worksheet =
                    excelPackage.Workbook.Worksheets
                    |> Seq.filter(fun m -> m.Hidden = eWorkSheetHidden.Visible)
                    |> Seq.tryItem index

                match worksheet with 
                | Some worksheet -> worksheet
                | None -> failwithf "Cannot get visible worksheet %A from %s, please check xlsx file" options excelPackage.File.FullName

            let excelworksheet = 
                match options with 
                | SheetGettingOptions.SheetName sheetName -> 
                    let sheet = excelPackage.Workbook.Worksheets.[sheetName.Value]
                    match sheet with 
                    | null -> failwithf "No sheet named %s was found in %A" sheetName.Value (excelPackage.File)
                    | _ -> sheet

                | SheetGettingOptions.SheetIndex index -> getVisibleSheetByIndex index
                | SheetGettingOptions.SheetNameOrSheetIndex (sheetName, index) ->
                    match Seq.tryFind (fun (worksheet: ExcelWorksheet) -> StringIC worksheet.Name = sheetName && worksheet.Hidden = eWorkSheetHidden.Visible) excelPackage.Workbook.Worksheets with 
                    | Some worksheet -> worksheet
                    | None -> getVisibleSheetByIndex index


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
    
