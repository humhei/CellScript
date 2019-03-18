// Learn more about F# at http://fsharp.org

open System
open CellScript.Server.Main
open System.IO


[<EntryPoint>]
let main argv =
    runCellScriptServer {WorkingDir = Path.GetDirectoryName __SOURCE_DIRECTORY__}
    0 // return an integer exit code
