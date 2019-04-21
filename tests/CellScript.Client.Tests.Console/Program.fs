// Learn more about F# at http://fsharp.org

open System
open CellScript.Core
open CellScript.Core.Tests
open Akkling
open CellScript.Client
open CellScript.Core.Remote
open Expecto.Logging
open Expecto

let testConfig =  
    { Expecto.Tests.defaultConfig with 
         parallelWorkers = 1
         verbosity = LogLevel.Debug }

let tests = 
    testList "All tests" [ 
        Tests.udfTests           
        Tests.evalTests         
    ] 

[<EntryPoint>]
let main argv =
    runTests testConfig tests
