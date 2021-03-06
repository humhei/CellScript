namespace CellScript.Core
open Deedle
open Extensions
open System
open Newtonsoft.Json
open Akka.Util
open ExcelProcess
open OfficeOpenXml

module Types =

    type CellScriptEvent<'EventArg> = CellScriptEvent of 'EventArg

    type SheetActiveArg =
        { WorkbookPath: string 
          WorksheetName: string }


    [<AttributeUsage(AttributeTargets.Property)>]
    type SubMsgAttribute() =
        inherit Attribute()


    type CommandSetting =
        { Shortcut: string option }

    type SerializableExcelReference =
        { ColumnFirst: int
          RowFirst: int
          ColumnLast: int
          RowLast: int
          WorkbookPath: string
          SheetName: string
          Content: obj[,] }
    with 
        member xlRef.CellAddress = ExcelAddress(xlRef.RowFirst, xlRef.ColumnFirst, xlRef.RowLast, xlRef.ColumnLast)

    [<RequireQualifiedAccess>]
    module SerializableExcelReference =

        let cellAddress (xlRef: SerializableExcelReference) = 
            xlRef.CellAddress

        let createByFile (rangeIndexer: string) sheetName workbookPath =
            use sheet = Excel.getWorksheetByName sheetName workbookPath
            use range = sheet.Cells.[rangeIndexer]

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
              SheetName = sheetName
              Content = content }   


        let positionText (xlRef: SerializableExcelReference) = 
            sprintf "%s_%s_%d_%d_%d_%d" xlRef.WorkbookPath xlRef.SheetName xlRef.RowFirst xlRef.ColumnFirst xlRef.RowLast xlRef.ColumnLast

    type CommandCaller = CommandCaller of SerializableExcelReference

    type internal ExcelEmpty() = class end
    type internal ExcelError() = class end

    [<RequireQualifiedAccess>]
    module internal CellValue =
        let private isArrayEmpty (v: obj) =
            match v with
            | :? float as v -> v = 0.
            | _ -> false


        let private isTextEmpty (v: obj) =
            match v with
            | :? string as v -> v = ""
            | _ -> false


        let private isEmpty (v: obj) =
            isArrayEmpty v || isTextEmpty v

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


    type ExcelFrame<'TRowKey,'TColumnKey when 'TRowKey: equality and 'TColumnKey: equality> =
        ExcelFrame of Frame<'TRowKey,'TColumnKey>
    with
        member x.AsFrame =
            let (ExcelFrame frame) = x
            frame


    [<RequireQualifiedAccess>]
    module ExcelFrame =
        let map mapping (ExcelFrame frame) =
            mapping frame
            |> ExcelFrame

        let mapValuesString mapping =
            map (Frame.mapValuesString mapping)

        let removeEmptyCols excelFrame =
            map (Frame.filterColValues (fun column ->
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
            headers |> Seq.mapi (fun i (header: obj) ->
                match header with
                | :? ExcelEmpty | :? ExcelError -> box (sprintf "Column%d" i)
                | _ -> header
            )
        let private fixContents contents =
            contents |> Array2D.map (fun (value: obj) ->
                match value with
                | :? ExcelEmpty | :? ExcelError -> box null
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

