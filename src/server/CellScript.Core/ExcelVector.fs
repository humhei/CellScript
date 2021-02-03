namespace CellScript.Core


[<RequireQualifiedAccess>]
type ExcelVector =
    | Column of list<obj>
    | Row of list<obj>

with 
    member x.Value =
        match x with 
        | ExcelVector.Column values -> values
        | ExcelVector.Row values -> values

    interface IToArray2D with 
        member x.ToArray2D() =
            match x with
            | ExcelVector.Column values -> 
                array2D (Seq.map Seq.singleton values)
            | ExcelVector.Row values -> array2D [values]

    static member Convert(array2D: obj[,]) =
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
   

