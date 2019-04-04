// Learn more about F# at http://fsharp.org

open System
open CellScript.Core
open CellScript.Core.Tests
open Akkling

[<EntryPoint>]
let main argv =
    //let f = excelFunctions()
    let client: Client<OuterMsg> = 
        { ClientPort = 6003
          ServerPort = 6001 }
    let remote = Remote(RemoteKind.Client client)

    let yes = 
        remote.RemoteActor.Value <? OuterMsg.TestString "Hello world"
        |> Async.RunSynchronously

    printfn "Hello World from F#!"
    Console.ReadLine() |> ignore
    0 // return an integer exit code
