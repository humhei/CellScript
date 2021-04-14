namespace CellScript.Core
open Deedle
open CellScript.Core.Extensions
open Akka.Util

type ExcelArray = private ExcelArray of ExcelFrame<int>
with

    member internal x.AsExcelFrame = 
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
    let mapFrame mapping (excelArray: ExcelArray) =
        ExcelFrame.mapFrame mapping excelArray.AsExcelFrame
        |> ExcelArray
        
    let mapValuesString mapping =
        mapFrame (Frame.mapValuesString mapping)

