namespace CellScript.Core
open Akkling
open Types


module Conversion =
    type BaseType =
        | Int of int
        | String of string
        | Float of float
    with 
        interface IToArray2D with 
            member x.ToArray2D() =
                match x with 
                | BaseType.Int x ->
                    array2D [[box x]]
                | BaseType.String x ->
                    array2D [[box x]]
                | BaseType.Float x ->
                    array2D [[box x]]

    type Record<'T> = Record of 'T
    with 
        interface IToArray2D with 
            member x.ToArray2D() =
                let result = 
                    let (Record recordValue) = x 
                    let tp = typeof<'T>
                    tp.GetProperties()
                    |> Array.map (fun prop ->
                        let value = prop.GetValue(recordValue)
                        [prop.Name ; value |> string]
                    )
                    |> array2D
                    |> Array2D.map box
                result


    let inline toArray2D (x : ^a when ^a :> IToArray2D) =
        x.ToArray2D()

