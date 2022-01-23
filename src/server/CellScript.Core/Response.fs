namespace CellScript.Core
open Extensions
open System
open Shrimp.FSharp.Plus

type IToArray2D =
    abstract member ToArray2D: unit -> ConvertibleUnion [,]
