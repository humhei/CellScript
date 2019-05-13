module Tests
open CellScript.Core.Tests
open Akkling
open CellScript.Core.Cluster
open Expecto
open CellScript.Core.Types
open System.IO
open System.Reflection
open Fake.IO.FileSystemOperators
open NLog
open CellScript.Client.Core
open CellScript.Client.Desktop
open System.Threading
open CellScript.Core

let dir = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location)

let book1 = dir </> "datas/book1.xlsx"

let equal actual expected = Expect.equal expected actual "passed"

let equalByResponse expected response = 
    let array2D: obj[,] = unbox response 
    Expect.equal array2D.[0,0] expected "passed"

let logger = NLog.FSharp.Logger(LogManager.GetCurrentClassLogger())

let createClient<'Msg>() =
    let client = NetCoreClient.create<'Msg> 9050
    Thread.Sleep(2000)
    client

let client = createClient<OuterMsg>()
let remoteServer = client.RemoteServer


let udfTests =
    let test (expected: obj[,]) msg =

        let response: obj [,] =
            client.InvokeFunction msg
            |> Async.RunSynchronously

        response
        |> Array2D.iteri(fun i j v ->
            equal expected.[i,j] v
        )

    let testSingleton (expected: 'value) msg =
        test (array2D [[box expected]]) msg

    testList "udf tests" [
        testCase "get excel functions" (fun _ ->
            /// ensure no exceptions
            let apiLambdas = NetCoreClient.apiLambdas logger client
            let excelFunctions = Registration.excelFunctions apiLambdas
            Expect.hasLength excelFunctions 6 "pass"
        )

        testCase "string" (fun _ ->
            testSingleton "Hello world" (OuterMsg.TestString "Hello world")
        )
    ]


let commandTests =
    testList "command tests" [

        testCase "get commands" (fun _ ->
            /// ensure no exceptions
            let apiLambdas = NetCoreClient.apiLambdas logger client
            let excelCommands = Registration.excelCommands apiLambdas

            Expect.hasLength excelCommands 1 "pass"
            Expect.equal excelCommands.[0].CommandLambda.Name "TestCommand" "pass"

        )

        testCase "test command" (fun _ ->

            let commandCaller = 
                SerializableExcelReference.createByFile "B3" "Script" book1
                |> CommandCaller

            let response: Result<obj, string> =
                remoteServer <? OuterMsg.TestCommand commandCaller
                |> Async.RunSynchronously

            match response with 
            | Result.Ok _ -> ()
            | Result.Error error -> failwithf "%s" error 
        )

    ]

