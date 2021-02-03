namespace CellScript.Core
open Deedle
open Extensions
open System
open OfficeOpenXml
open System.IO

[<AutoOpen>]
module Types = 

    type SheetReference =
        { WorkbookPath: string 
          SheetName: string }
    
    type RangeGettingOptions =
        | RangeIndexer of string
        | UserRange

    type ExcelWorksheet with
        member sheet.GetRange options = 
            match options with
            | RangeIndexer indexer ->
                sheet.Cells.[indexer]

            | UserRange ->
                let indexer = sheet.Dimension.Start.Address + ":" + sheet.Dimension.End.Address
                sheet.Cells.[indexer]

    type SheetGettingOptions =
        | SheetName of string
        | SheetIndex of int
        | SheetNameOrSheetIndex of sheetName: string * index: int
    

    type ExcelPackage with
        member excelPackage.GetWorkSheet (options) =
            match options with 
            | SheetGettingOptions.SheetName sheetName -> excelPackage.Workbook.Worksheets.[sheetName]
            | SheetGettingOptions.SheetIndex index -> excelPackage.Workbook.Worksheets.[index]
            | SheetGettingOptions.SheetNameOrSheetIndex (sheetName, index) ->
                match Seq.tryFind (fun (worksheet: ExcelWorksheet) -> worksheet.Name = sheetName) excelPackage.Workbook.Worksheets with 
                | Some worksheet -> worksheet
                | None -> excelPackage.Workbook.Worksheets.[index]


    type ExcelRangeContactInfo =
        { ColumnFirst: int
          RowFirst: int
          ColumnLast: int
          RowLast: int
          WorkbookPath: string
          SheetName: string
          Content: obj[,] }
    with 
        member xlRef.CellAddress = ExcelAddress(xlRef.RowFirst, xlRef.ColumnFirst, xlRef.RowLast, xlRef.ColumnLast)
    
        override xlRef.ToString() = sprintf "%s_%s_%d_%d_%d_%d" xlRef.WorkbookPath xlRef.SheetName xlRef.RowFirst xlRef.ColumnFirst xlRef.RowLast xlRef.ColumnLast
    
    
    type SheetContentsEmptyException(error: string) =
        inherit Exception(error)

    [<RequireQualifiedAccess>]
    module ExcelRangeContactInfo =
    
        let cellAddress (xlRef: ExcelRangeContactInfo) = xlRef.CellAddress
    
        let readFromFile (rangeGettingArg: RangeGettingOptions) (sheetGettingArg: SheetGettingOptions) workbookPath =
            if not <| File.Exists workbookPath then failwithf "File %s not exists" workbookPath
    
            use excelPackage = new ExcelPackage(FileInfo(workbookPath))
            let sheet = excelPackage.GetWorkSheet sheetGettingArg
            match sheet.Dimension  with 
            | null -> 
                raise(SheetContentsEmptyException (sprintf "%A Contents in sheet %s is empty" rangeGettingArg sheet.Name))
                
            | _ ->
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
                  WorkbookPath = workbookPath
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
          WorkbookPath: string
          SheetName: string }
    with 
        member xlRef.CellAddress = ExcelAddress(xlRef.RowFirst, xlRef.ColumnFirst, xlRef.RowLast, xlRef.ColumnLast)
    
    [<RequireQualifiedAccess>]
    module SerializableExcelReference =
        let ofExcelRangeContactInfo (xlRef: ExcelRangeContactInfo) =
            { ColumnFirst = xlRef.ColumnFirst
              RowFirst = xlRef.RowFirst
              ColumnLast = xlRef.ColumnLast
              RowLast = xlRef.RowLast
              WorkbookPath = xlRef.WorkbookPath
              SheetName = xlRef.SheetName }
    
        let toExcelRangeContactInfo content (xlRef: SerializableExcelReference) : ExcelRangeContactInfo =
            { ColumnFirst = xlRef.ColumnFirst
              RowFirst = xlRef.RowFirst
              ColumnLast = xlRef.ColumnLast
              RowLast = xlRef.RowLast
              WorkbookPath = xlRef.WorkbookPath
              SheetName = xlRef.SheetName
              Content = content }
    
        let cellAddress xlRef =
            ExcelAddress(xlRef.RowFirst, xlRef.ColumnFirst, xlRef.RowLast, xlRef.ColumnLast)
    
    

    /// both contents and headers must be fixed before created
    ///
    /// code base of excelArray and table
    type internal ExcelFrame<'TRowKey,'TColumnKey when 'TRowKey: equality and 'TColumnKey: equality> =
        private ExcelFrame of Frame<'TRowKey,'TColumnKey>
    with
        member x.AsFrame =
            let (ExcelFrame frame) = x
            frame
    
    [<RequireQualifiedAccess>]
    module internal ExcelFrame =

        let mapFrame mapping (ExcelFrame frame) =
            mapping frame
            |> ExcelFrame
    
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

        let ofArray2D (array2D: obj[,]) =
            let fixContents contents =
                contents |> Array2D.map (fun (value: obj) ->
                    match value with
                    | :? ExcelErrorValue -> box null
                    | _ -> value
                )
    
            array2D
            |> Array2D.rebase
            |> fixContents
            |> Frame.ofArray2D
            |> ExcelFrame
    
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
            Frame.ofRecords records
            |> ExcelFrame

    
    