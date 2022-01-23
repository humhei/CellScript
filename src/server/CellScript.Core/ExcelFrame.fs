namespace CellScript.Core
open Deedle
open Extensions
open System
open OfficeOpenXml
open System.IO
open Shrimp.FSharp.Plus

[<AutoOpen>]
module ExcelFrame = 

    /// both contents and headers must be fixed before created
    ///
    /// code base of excelArray and table
    ///
    ///
    type internal ExcelFrame<'TColumnKey when 'TColumnKey: equality> =
        ExcelFrame of Frame<int, 'TColumnKey>
    with
        member x.AsFrame =
            let (ExcelFrame frame) = x
            frame
    
    [<RequireQualifiedAccess>]
    module internal ExcelFrame =
    
        let private ensureDatasValid (frame: ExcelFrame<_>) =
            frame
    
        let mapFrame mapping (ExcelFrame frame) =
            mapping frame
            |> ExcelFrame
            |> ensureDatasValid
    
        let toArray2D (ExcelFrame frame): IConvertible [,] =
            frame.ToArray2D(null)
            |> Array2D.map(fun (m: obj) ->
                match m with 
                | null -> null 
                | :? IConvertible as convertible -> convertible 
                | :? ICellValue as v -> v.Convertible 
                | _ -> failwithf "type of cell value %A should either be ICellValue or IConvertible" (m.GetType())
            )
            
    
        let ofArray2DWithConvitable (array2D: IConvertible[,]) =
            array2D
            |> Array2D.rebase
            |> Frame.ofArray2D
            |> ExcelFrame
    
        let ofArray2D (array2D: obj[,]) =
            array2D
            |> Array2D.map fixContent
            |> ofArray2DWithConvitable
    


    