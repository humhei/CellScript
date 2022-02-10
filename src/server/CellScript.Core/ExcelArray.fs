namespace CellScript.Core
open Deedle
open CellScript.Core.Extensions
open Akka.Util
open System
open Shrimp.FSharp.Plus

type ExcelArray = private ExcelArray of ExcelFrame<int>
with

    member internal x.AsExcelFrame = 
        let (ExcelArray frame) = x
        frame

    member x.AsFrame = x.AsExcelFrame.AsFrame


    static member internal Convert(array2D: IConvertible[,]) =
        let frame = ExcelFrame.ofArray2DWithConvitable array2D
        ExcelArray(frame)

    static member Convert(array2D: ConvertibleUnion[,]) =
        array2D
        |> Array2D.map(fun m -> m.Value)
        |> ExcelArray.Convert

    static member Convert(array2D: obj[,]) =
        let frame = ExcelFrame.ofArray2D array2D
        ExcelArray(frame)


    member x.ToArray2D() =
        ExcelFrame.toArray2D x.AsExcelFrame
        

    interface IToArray2D with 
        member x.ToArray2D() =
            ExcelFrame.toArray2D x.AsExcelFrame
            |> Array2D.map ConvertibleUnion.Convert

    interface ISurrogated with 
        member x.ToSurrogate(system) = 
            ExcelArraySurrogate ((x :> IToArray2D).ToArray2D()) :> ISurrogate

and private ExcelArraySurrogate = ExcelArraySurrogate of ConvertibleUnion [,]
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
        

    let fillEmptyUp (excelArray: ExcelArray) =
        mapFrame Frame.fillEmptyUp excelArray