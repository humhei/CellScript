// Learn more about F# at http://fsharp.org

open CellScript.Core
open Expecto.Logging
open Expecto
open Akka.Configuration
open Akka.Cluster
open Akkling
open Akka.Cluster.Tools.Singleton
open Akkling.Cluster.ClusterExtensions
open System


let testConfig =  
    { Expecto.Tests.defaultConfig with 
         parallelWorkers = 1
         verbosity = LogLevel.Debug }

let tests = 
    testList "All tests" [ 
        Tests.commandTests
        Tests.evalTests         
    ] 


[<EntryPoint>]
let main argv =
    runTests testConfig tests
