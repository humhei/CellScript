namespace CellScript.Core
open Deedle
open Registration
open CellScript.Core.Extensions
open Types

[<CustomParamConversion>]
type ExcelSeries<'T> =
    | Column of seq<'T>
    | Row of seq<'T>

with
    member private x.ToExcelArray() =
        let asRow series =
            array2D [series]
            |> Frame.ofArray2D

        match x with
        | ExcelSeries.Row series -> asRow series
        | ExcelSeries.Column series ->
            asRow series
            |> Frame.transpose
        |> ExcelFrame
        |> ExcelArray

    static member Convert(cast) =
        fun (values: obj[,]) ->
            let l1 = values.GetLength(0)
            let l2 = values.GetLength(1)
            let result =
                if l2 > 0 then
                    values.[*,0] |> Array.map cast |> Seq.ofArray |> ExcelSeries.Column
                elif l1 > 0 then
                    values.[0,*] |> Array.map cast |> Seq.ofArray |> ExcelSeries.Row
                else
                    failwith "not implemented"
            result
        |> CustomParamConversion.array2D

    interface ICustomReturn with
        member x.ReturnValue() =
            (x.ToExcelArray() :> ICustomReturn).ReturnValue()

[<RequireQualifiedAccess>]
module ExcelSeries =
    let row values = ExcelSeries.Row values
    let column values = ExcelSeries.Column values

    let leftOf (pattern: string) (input: ExcelArray) =
        input
        |> ExcelArray.mapValuesString (String.leftOf pattern)
