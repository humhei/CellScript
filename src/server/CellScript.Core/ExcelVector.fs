namespace CellScript.Core
open System
open Shrimp.FSharp.Plus

[<RequireQualifiedAccess>]
type ExcelVector =
    | Column of list<ConvertibleUnion>
    | Row of list<ConvertibleUnion>

with 
    member x.Value =
        match x with 
        | ExcelVector.Column values -> values
        | ExcelVector.Row values -> values

    member x.MapValue(fValue) =
        match x with 
        | ExcelVector.Column values -> fValue values |> ExcelVector.Column
        | ExcelVector.Row values -> fValue values |> ExcelVector.Row


    //member x.AutoNumberic() =
    //    x.MapValue(fun values ->
    //        match ListObj.toNumberic values with 
    //        | Some v -> v
    //        | None -> values) 

    member x.ToArray2D() =
        match x with
        | ExcelVector.Column values -> 
            array2D (Seq.map Seq.singleton values)
        | ExcelVector.Row values -> array2D [values]

    interface IToArray2D with 
        member x.ToArray2D() = 
            x.ToArray2D()

    static member Convert(array2D: ConvertibleUnion[,]) =
        let l1 = array2D.GetLength(0)
        let l2 = array2D.GetLength(1)
        let result =
            if l2 > 0 && l1 = 1 then
                array2D.[0,*] |> List.ofArray |> ExcelVector.Row
            elif l1 > 0 && l2 = 1 then
                array2D.[*,0] |> List.ofArray |> ExcelVector.Column
            elif l1 = 0 && l2 = 0 then
                failwith "empty inputs"
            else 
                failwithf "Multiple column or rows as inputs %A" array2D

        result
   
    static member Convert(array2D: obj[,]) =
        array2D
        |> Array2D.map (fixRawContent >> ConvertibleUnion.Convert)
        |> ExcelVector.Convert


