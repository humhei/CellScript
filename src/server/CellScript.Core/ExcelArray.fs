namespace CellScript.Core
open Deedle
open Registration
open CellScript.Core.Extensions
open Types

[<CustomParamConversion>]
type ExcelArray = ExcelArray of ExcelFrame<int,int>
with
    member x.MapValues mapping =
        let (ExcelArray frame) = x
        mapping frame
        |> ExcelArray

    member x.RemoveEmptyCols() =
        x.MapValues ExcelFrame.removeEmptyCols

    static member convert =
        fun (table:obj[,]) ->
            Frame.ofArray2D (Array2D.rebase table)
            |> ExcelFrame
            |> ExcelArray

    static member Convert() =
        ExcelArray.convert
        |> CustomParamConversion.array2D

    interface ICustomReturn with
        member x.ReturnValue() =
            let (ExcelArray excelFrame) = x.RemoveEmptyCols()
            ExcelFrame.toArray2D excelFrame

[<RequireQualifiedAccess>]
module ExcelArray =
    let map mapping (ExcelArray array) =
        array
        |> mapping
        |> ExcelArray
        
    let mapValuesString mapping =
        map (ExcelFrame.mapValuesString mapping)
