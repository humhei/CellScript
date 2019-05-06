// Learn more about F# at http://fsharp.org

open System
open CellScript.Core
open CellScript.Core.Tests
open Akkling
open CellScript.Client
open CellScript.Core.Remote
open Expecto.Logging
open Expecto
open Akkling.Persistence

let testConfig =  
    { Expecto.Tests.defaultConfig with 
         parallelWorkers = 1
         verbosity = LogLevel.Debug }

let tests = 
    testList "All tests" [ 
        Tests.udfTests    
        Tests.commandTests
        Tests.evalTests         
    ] 


let system = System.create "persisting-sys" <| Configuration.defaultConfig()

type CounterChanged =
    { Delta : int }

type CounterCommand =
    | Inc
    | Dec
    | GetState

type CounterMessage =
    | Command of CounterCommand
    | Event of CounterChanged




[<EntryPoint>]
let main argv =

    runTests testConfig tests
