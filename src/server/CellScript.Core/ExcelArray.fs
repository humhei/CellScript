namespace CellScript.Core
open Deedle
open CellScript.Core.Extensions
open Types
open Newtonsoft.Json
open Akka.Util

type ExcelArray = ExcelArray of ExcelFrame<int,int>
with

    member x.AsExcelFrame = 
        let (ExcelArray frame) = x
        frame

    member x.AsFrame = x.AsExcelFrame.AsFrame

    static member Convert(array2D: obj[,]) =
        let frame = ExcelFrame.ofArray2D array2D
        ExcelArray(frame)

    interface IToArray2D with 
        member x.ToArray2D() =
            ExcelFrame.toArray2D x.AsExcelFrame

    interface ISurrogated with 
        member x.ToSurrogate(system) = 
            ExcelArraySurrogate ((x :> IToArray2D).ToArray2D()) :> ISurrogate

and private ExcelArraySurrogate = ExcelArraySurrogate of obj [,]
with 
    interface ISurrogate with 
        member x.FromSurrogate(system) = 
            let (ExcelArraySurrogate array2D) = x
            ExcelArray.Convert array2D :> ISurrogated


[<RequireQualifiedAccess>]
module ExcelArray =
    let map mapping (excelArray: ExcelArray) =
        excelArray.AsFrame
        |> mapping
        |> ExcelFrame
        |> ExcelArray
        
    let mapValuesString mapping =
        map (Frame.mapValuesString mapping)
