namespace CellScript.Core
open Deedle
open Extensions
open System
open OfficeOpenXml
open System.IO

type CellScriptEvent<'EventArg> = CellScriptEvent of 'EventArg

type SheetReference =
    { WorkbookPath: string 
      SheetName: string }

[<AttributeUsage(AttributeTargets.Property)>]
type SubMsgAttribute() =
    inherit Attribute()


[<AttributeUsage(AttributeTargets.Property)>]
type NameInheritedSubMsgAttribute() =
    inherit Attribute()

type CommandSetting =
    { Shortcut: string option }

type ExcelErrorEnum =
    | Null = 0s
    | Div = 7s
    | Value = 15s
    | Ref = 23s
    | Name = 29s
    | Num = 36s
    | NA = 42s
    | GettingData = 43s


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


[<RequireQualifiedAccess>]
module ExcelRangeContactInfo =

    let cellAddress (xlRef: ExcelRangeContactInfo) = xlRef.CellAddress

    let readFromFile (rangeGettingArg: RangeGettingOptions) (sheetGettingArg: SheetGettingOptions) workbookPath =
        if not <| File.Exists workbookPath then failwithf "File %s not exists" workbookPath

        use excelPackage = new ExcelPackage(FileInfo(workbookPath))
        let sheet = excelPackage.GetWorkSheet sheetGettingArg
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

[<AutoOpen>]
module ExcelRangeContactInfoExtensions =
    type ExcelPackage with 
        member excelPackage.GetWorkSheet(xlRef: ExcelRangeContactInfo) =
            excelPackage.Workbook.Worksheets
            |> Seq.find (fun worksheet -> worksheet.Name = xlRef.SheetName)

        member excelPackage.GetExcelRange(xlRef: ExcelRangeContactInfo) =
            let workbookSheet = excelPackage.GetWorkSheet xlRef
            workbookSheet.Cells.[xlRef.RowFirst, xlRef.ColumnFirst, xlRef.RowLast, xlRef.ColumnLast]

    type ExcelWorksheet with
        member sheet.GetRange arg = 
            match arg with
            | RangeIndexer indexer ->
                sheet.Cells.[indexer]

            | UserRange ->
                let indexer = sheet.Dimension.Start.Address + ":" + sheet.Dimension.End.Address
                sheet.Cells.[indexer]

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



type CommandCaller = CommandCaller of ExcelRangeContactInfo

type FunctionSheetReference = FunctionSheetReference of SheetReference



[<RequireQualifiedAccess>]
module internal CellValue =

    let private isTextEmpty (v: obj) =
        match v with
        | :? string as v -> v = ""
        | _ -> false


    let private isEmpty (v: obj) =
        match v with 
        | null -> true
        | _ ->
            isTextEmpty v

    let isNotEmpty (v: obj) =
        isEmpty v |> not

    /// no missing or cell empty
    let (|HasSense|NoSense|) (v: obj option) =
        match v with
        | Some value -> if isEmpty v then NoSense else HasSense value
        | None -> NoSense

    let isMissingOrCellValueEmpty (v: obj option) =
        match v with
        | HasSense _ -> true
        | _ -> false

    let isNotMissingOrCellValueEmpty (v: obj option) =
        isMissingOrCellValueEmpty v |> not

/// both contents and headers must be fixed before created
type ExcelFrame<'TRowKey,'TColumnKey when 'TRowKey: equality and 'TColumnKey: equality> =
    private ExcelFrame of Frame<'TRowKey,'TColumnKey>
with
    member x.AsFrame =
        let (ExcelFrame frame) = x
        frame

[<RequireQualifiedAccess>]
module ExcelFrame =
    let mapFrame mapping (ExcelFrame frame) =
        mapping frame
        |> ExcelFrame

    let mapValuesString mapping =
        mapFrame (Frame.mapValuesString mapping)

    let removeEmptyCols excelFrame =
        mapFrame (Frame.filterColValues (fun column ->
            Seq.exists CellValue.isNotEmpty column.Values
        )) excelFrame

    let toArray2D (ExcelFrame frame) =
        frame.ToArray2D(null)

    let toArray2DWithHeader (ExcelFrame frame) =

        let header = frame.ColumnKeys |> Seq.map box

        let contents = frame.ToArray2D(null) |> Array2D.toSeqs
        let result =
            Seq.append [header] contents
            |> array2D
        result

    let private fixHeaders headers =
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

    let private fixContents contents =
        contents |> Array2D.map (fun (value: obj) ->
            match value with
            | :? ExcelErrorValue -> box null
            | _ -> value
        )

    let ofArray2D (array2D: obj[,]) =
        array2D
        |> fixContents
        |> Array2D.rebase
        |> Frame.ofArray2D
        |> ExcelFrame

    let ofArray2DWithHeader (array: obj[,]) =

        let array = Array2D.rebase array

        let headers = array.[0,*] |> fixHeaders |> Seq.map (fun header -> header.ToString())

        let contents = array.[1..,*] |> Array2D.map box |> fixContents |> Frame.ofArray2D

        Frame.indexColsWith headers contents
        |> ExcelFrame


