namespace CellScript.Core
open Deedle
open Extensions
open System
open OfficeOpenXml
open System.IO
open Shrimp.FSharp.Plus

[<AutoOpen>]
module Types = 

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
        static member DefaultValue = ColumnAutofitOptions.AutoFit(10., 50.)



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


        member x.LoadFromArraysAsTable(arrays: obj [, ], ?columnAutofitOptions, ?tableName: string, ?tableStyle: Table.TableStyles) =
            let worksheet = x.Value
            
            worksheet.Cells.["A1"].LoadFromArray2D(arrays) |> ignore
                
            let userRange = x.GetRange(RangeGettingOptions.UserRange)

            let tab = worksheet.Tables.Add(userRange, defaultArg tableName "Table1")

            tab.TableStyle <- 
                defaultArg tableStyle Table.TableStyles.Medium2

            match defaultArg columnAutofitOptions ColumnAutofitOptions.DefaultValue with 
            | ColumnAutofitOptions.AutoFit(minimumSize, maximumSize) ->
                userRange.AutoFitColumns(minimumSize, maximumSize)

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
    type ValidExcelWorksheet(visibleExcelWorksheet: VisibleExcelWorksheet) =
        do
            match visibleExcelWorksheet.Value.Dimension with
            | null -> raise(SheetContentsEmptyException (sprintf "Cannot create VisibleExcelWorksheet: Contents in sheet %s is empty" visibleExcelWorksheet.Value.Name))
            | _ -> ()


        member x.Value = visibleExcelWorksheet.Value

        member x.Name = x.Value.Name

        member x.VisibleExcelWorksheet = visibleExcelWorksheet

        new (excelworksheet: ExcelWorksheet) =
            ValidExcelWorksheet(VisibleExcelWorksheet.Create(excelworksheet))
        



    type SheetGettingOptions =
        | SheetName of string
        | SheetIndex of int
        | SheetNameOrSheetIndex of sheetName: string * index: int
    with 
        static member DefaultValue = 
            SheetGettingOptions.SheetNameOrSheetIndex (SHEET1, 0)

    type ExcelPackage with
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
                | SheetGettingOptions.SheetName sheetName -> excelPackage.Workbook.Worksheets.[sheetName]
                | SheetGettingOptions.SheetIndex index -> getVisibleSheetByIndex index
                | SheetGettingOptions.SheetNameOrSheetIndex (sheetName, index) ->
                    match Seq.tryFind (fun (worksheet: ExcelWorksheet) -> worksheet.Name = sheetName && worksheet.Hidden = eWorkSheetHidden.Visible) excelPackage.Workbook.Worksheets with 
                    | Some worksheet -> worksheet
                    | None -> getVisibleSheetByIndex index

            VisibleExcelWorksheet.Create excelworksheet

        member excelPackage.GetValidWorksheet (options) =
            excelPackage.GetVisibleWorksheet(options)
            |> ValidExcelWorksheet


    type ExcelRangeContactInfo =
        { ColumnFirst: int
          RowFirst: int
          ColumnLast: int
          RowLast: int
          XlsxFile: XlsxFile
          SheetName: string
          Content: obj[,] }
    with 
        member x.WorkbookPath = x.XlsxFile.Path

        member xlRef.CellAddress = ExcelAddress(xlRef.RowFirst, xlRef.ColumnFirst, xlRef.RowLast, xlRef.ColumnLast)
    
        override xlRef.ToString() = sprintf "%s %s (%d, %d, %d, %d)" xlRef.WorkbookPath xlRef.SheetName xlRef.RowFirst xlRef.ColumnFirst xlRef.RowLast xlRef.ColumnLast
    
    


    [<RequireQualifiedAccess>]
    module ExcelRangeContactInfo =
    
        let cellAddress (xlRef: ExcelRangeContactInfo) = xlRef.CellAddress
    
        let readFromFile (rangeGettingArg: RangeGettingOptions) (sheetGettingArgs: SheetGettingOptions) (xlsxFile: XlsxFile) =
    
            use excelPackage = new ExcelPackage(FileInfo(xlsxFile.Path))
            let sheet = excelPackage.GetValidWorksheet sheetGettingArgs

            let sheet = sheet.VisibleExcelWorksheet

            use range = sheet.GetRange(rangeGettingArg)
    
            let rowStart = range.Start.Row
            let rowEnd = range.End.Row
            let columnStart = range.Start.Column
            let columnEnd = range.End.Column
    
            let content =
                array2D
                    [ for i = rowStart to rowEnd do 
                        yield
                            [ for j = columnStart to columnEnd do yield range.[i,j].Value ]
                    ] 
    
            { ColumnFirst = columnStart
              RowFirst = rowStart
              ColumnLast = columnEnd
              RowLast = rowEnd
              XlsxFile = xlsxFile
              SheetName = sheet.Name
              Content = content }   
    

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
    
    

    /// both contents and headers must be fixed before created
    ///
    /// code base of excelArray and table
    ///
    ///
    type internal ExcelFrame<'TColumnKey when 'TColumnKey: equality> =
        private ExcelFrame of Frame<int, 'TColumnKey>
    with
        member x.AsFrame =
            let (ExcelFrame frame) = x
            frame
    
    [<RequireQualifiedAccess>]
    module internal ExcelFrame =

        //let private ensureIsNotEmptyWhenCreating(excelFrame: ExcelFrame<_, _>) =
        //    match excelFrame.AsFrame.IsEmpty with 
        //    | true -> failwith "Cannot create ExcelFrame when it is empty"
        //    | false -> excelFrame
            
        let private ensureDatasValid (frame: ExcelFrame<_>) =
            frame
            //match frame.AsFrame.RowCount with 
            //| 0 -> failwith "Frame is empty"
            //| _ -> frame


        let mapFrame mapping (ExcelFrame frame) =
            mapping frame
            |> ExcelFrame
            |> ensureDatasValid
    
        let mapValuesString mapping =
            mapFrame (Frame.mapValuesString mapping)
    
        let toArray2D (ExcelFrame frame) =
            frame.ToArray2D(null)
    
        let toArray2DWithHeader (ExcelFrame frame) =
    
            let header = frame.ColumnKeys |> Seq.map (fun m -> box(m.ToString()))
    
            let contents = frame.ToArray2D(null) |> Array2D.toSeqs
            let result =
                Seq.append [header] contents
                |> array2D
            result

        let private fixContent (content: obj) =
            match content with
            | :? ExcelErrorValue -> box null
            | _ -> 
                content


        let ofArray2D (array2D: obj[,]) =

            let fixContents contents =
                contents |> Array2D.map fixContent
            array2D
            |> Array2D.rebase
            |> fixContents
            |> Frame.ofArray2D
            |> ExcelFrame
            //|> ensureIsNotEmptyWhenCreating
    
        let ofArray2DWithHeader (array: obj[,]) =
            let fixHeaders headers =
                let headers = List.ofSeq headers
                headers |> List.mapi (fun i (header: obj) ->
                    match header with
                    | :? ExcelErrorValue 
                    | null -> 
                        (sprintf "Column%d" i)
                    | _ -> 
                        headers.[0 .. i - 1]
                        |> List.filter(fun preHeader -> preHeader = header)
                        |> function
                            | matchHeaders when matchHeaders.Length > 0 ->
                                let headerText = header.ToString()
                                sprintf "%s%d" headerText (matchHeaders.Length + 1) 
                            | [] -> header.ToString()
                            | _ -> failwith "invalid token"
                )

            let array = Array2D.rebase array
    
            let headers = array.[0,*] |> fixHeaders |> Seq.map (fun header -> header.ToString())
    
            let contents = ofArray2D array.[1..,*] 
    
            Frame.indexColsWith headers contents.AsFrame
            |> ExcelFrame
    
        let ofRecords (records: seq<'record>) =
            (Frame.ofRecords records)
            |> ExcelFrame


        let ofFrameWithHeaders (frame: Frame<_, _>) =
            frame
            |> Frame.mapValues fixContent
            |> ExcelFrame
            |> ensureDatasValid
