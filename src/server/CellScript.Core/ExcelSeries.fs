namespace CellScript.Core
open Registration
open CellScript.Core.Extensions

[<CustomParamConversion>]
type ExcelVector<'T> =
    | Column of seq<'T>
    | Row of seq<'T>

[<RequireQualifiedAccess>]
module ExcelSeries =
    let row values = ExcelVector.Row values
    let column values = ExcelVector.Column values

