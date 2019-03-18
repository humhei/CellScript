namespace CellScript.Core
open CellScript.Core.Extensions

module UDF =
    [<RequireQualifiedAccess>]
    module ExcelArray =

        let leftOf (pattern: string) (input: ExcelArray) =
            input
            |> ExcelArray.mapValuesString (String.leftOf pattern)
