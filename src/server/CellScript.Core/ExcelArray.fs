namespace CellScript.Core
open Deedle
open Registration
open CellScript.Core.Extensions
open Types
open Newtonsoft.Json

[<CustomParamConversion>]
type ExcelArray(frame: ExcelFrame<int,int>) = 
    [<System.NonSerialized>]
    let frame = frame

    [<JsonProperty>]
    let serializedData = ExcelFrame.toArray2D frame

    [<JsonConstructor>]
    new (array2D: obj[,]) =
        let frame = ExcelFrame.ofArray2D array2D
        new ExcelArray(frame)

    member x.AsFrame = frame.AsFrame

    member x.AsExcelFrame = frame



[<RequireQualifiedAccess>]
module ExcelArray =
    let map mapping (excelArray: ExcelArray) =
        excelArray.AsFrame
        |> mapping
        |> ExcelFrame
        |> ExcelArray
        
    let mapValuesString mapping =
        map (Frame.mapValuesString mapping)
