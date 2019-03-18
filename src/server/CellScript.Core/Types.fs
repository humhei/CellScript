namespace CellScript.Core
open Deedle
open Extensions
open System

module Types =

    type SerializableExcelReference =
        { ColumnFirst: int
          RowFirst: int
          ColumnLast: int
          RowLast: int }

    type ExcelEmpty() = class end
    type ExcelError() = class end


    [<RequireQualifiedAccess>]
    module CellValue =
        let isArrayEmpty (v: obj) =
            match v with
            | :? float as v -> v = 0.
            | _ -> false

        let isNotArrayEmpty (v: obj) =
            isArrayEmpty v |> not

        let isTextEmpty (v: obj) =
            match v with
            | :? string as v -> v = ""
            | _ -> false

        let isNotTextEmpty (v: obj) =
            isTextEmpty v |> not

        let isEmpty (v: obj) =
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

        let ofArray2DWithHeader (array:obj[,]) =
            let fixHeaders headers =
                headers |> Seq.mapi (fun i (header: obj) ->
                    match header with
                    | :? ExcelEmpty | :? ExcelError -> box (sprintf "Column%d" i)
                    | _ -> header
                )
            let fixContents contents =
                contents |> Array2D.map (fun (value: obj) ->
                    match value with
                    | :? ExcelEmpty | :? ExcelError -> box null
                    | _ -> value
                )

            let array = Array2D.rebase array

            let headers = array.[0,*] |> fixHeaders |> Seq.map (fun header -> header.ToString())

            let contents = array.[1..,*] |> Array2D.map box |> fixContents |> Frame.ofArray2D

            Frame.indexColsWith headers contents
            |> ExcelFrame
