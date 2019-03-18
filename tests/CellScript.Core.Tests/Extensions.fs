namespace CellScript.Core.Tests
open System.Text.RegularExpressions
open LiteDB.FSharp.Extensions
open LiteDB.FSharp.Linq
open CellScript.Core.Types
open CellScript.Core
open ExcelProcess.MatrixParsers
open FParsec
open OfficeOpenXml
open CellScript.Core.Extensions
open Deedle

module Extensions =
    [<RequireQualifiedAccess>]
    module String =
        let returnZH (text: string) =
            let m = Regex.Match(text,"[\\u4e00-\\u9fa5]+")
            m.Value

    [<RequireQualifiedAccess>]
    module LiteRepository =
        let addOrUpdate queryFactory initialValue updateFactory repository =
            LiteRepository.query repository
            |> LiteQueryable.tryFind (Expr.prop (fun a -> queryFactory a))
            |> function
                | Some value ->
                    let newValue = updateFactory value
                    LiteRepository.insertItem newValue repository
                | None ->
                    LiteRepository.insertItem initialValue repository

    [<AutoOpen>]
    module ExcelProcesser =
        let runMatrixParserForArray2D parser (array: obj[,]) =
            use p = new ExcelPackage()
            let range = ExcelRangeBase.ofArray2D p array
            runMatrixParserForRangeWith parser range

        let rec runMatrixParserForArray2DUntilExatlyOne parsers (array: obj[,]) =
            match parsers with
            | parser :: t ->
                let ms =
                    runMatrixParserForArray2D parser array

                let l = Seq.length ms.State

                match l with
                | 0 -> runMatrixParserForArray2DUntilExatlyOne t array
                | 1 -> ms
                | _ -> failwithf "get more than one results %A" (List.ofSeq ms.State)

            | [] -> failwithf "no results when parse %A with parsers %A" array parsers

        let runMatrixParserForSeq parser (seq: seq<string>) =
            array2D [Seq.map box seq]
            |> runMatrixParserForArray2D parser

        [<RequireQualifiedAccess>]
        module Seq =
            let partitionByMatrixParsers (parsers: MatrixParser<'a list> list) (seq: seq<string>) =
                let ms =
                    array2D [Seq.map box seq]
                    |> runMatrixParserForArray2DUntilExatlyOne parsers

                let state = Seq.exactlyOne ms.State

                let stateValues,others =
                    let range = Seq.exactlyOne ms.XLStream.userRange
                    let address = ExcelAddress(range.Address)
                    let column = address.Start.Column
                    let values = List.ofSeq seq
                    values.[column-1..column-2+state.Length] ,values.[0..column-2] @ values.[column-1+state.Length..]

                state,stateValues,others