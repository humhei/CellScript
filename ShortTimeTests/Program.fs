// Learn more about F# at http://fsharp.org

open System
open CellScript.Core
open CellScript.Core.Extensions
open Deedle
open Shrimp.FSharp.Plus
open System.Collections.Generic
open System.IO
open Hyperion
open Akka.Util
open Akkling
open Newtonsoft.Json

[<EntryPoint>]
let main argv =
    let cc = JsonConvert.SerializeObject(stringIC "Yes")
    let pp = JsonConvert.DeserializeObject<stringIC>(cc)

    let c = 眼镜复杂外箱贴.Module_眼镜复杂外箱贴.storages()

    let yes = stringIC "Yes" = stringIC "yes"
    printfn "Hello World from F#!"
    0 // return an integer exit code
