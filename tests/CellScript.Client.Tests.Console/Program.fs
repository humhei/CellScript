﻿// Learn more about F# at http://fsharp.org

open System
open CellScript.Client.Tests.Registration

[<EntryPoint>]
let main argv =
    let f = excelFunctions()
    printfn "Hello World from F#!"
    Console.ReadLine() |> ignore
    0 // return an integer exit code
