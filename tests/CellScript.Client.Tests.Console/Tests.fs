module Tests
open CellScript.Core.Tests
open Akkling
open Expecto
open CellScript.Core.Types
open System.IO
open System.Reflection
open Fake.IO.FileSystemOperators
open NLog
open CellScript.Client.Core
open System.Threading
open CellScript.Client.Desktop
open CellScript.Core
open OfficeOpenXml
open CellScript.Core.Extensions

let pass() = Expect.isTrue true "pass"
let fail() = Expect.isTrue false "failed"

let dir = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location)

let book1 = dir </> "datas/book1.xlsx"

let equal actual expected = Expect.equal expected actual "passed"

let equalByResponse expected response = 
    let array2D: obj[,] = unbox response 
    Expect.equal array2D.[0,0] expected "passed"

let logger = NLog.FSharp.Logger(LogManager.GetCurrentClassLogger())

let createClient<'Msg>() =
    let client = NetCoreClient.create<'Msg> 9050 id
    Thread.Sleep(2000)
    client

let client = lazy createClient<OuterMsg>()
let remoteServer = 
    lazy client.Value.RemoteServer

let functionTests =
    testList "function tests" [
    ]


let udfTests =
    let test (expected: obj[,]) msg =

        let response: obj [,] =
            client.Value.InvokeFunction msg
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
            let apiLambdas = NetCoreClient.apiLambdas logger client.Value
            let excelFunctions = Registration.excelFunctions apiLambdas
            Expect.hasLength excelFunctions 9 "pass"
            let funcNames = 
                excelFunctions 
                |> List.ofSeq 
                |> List.map (fun func -> func.FunctionLambda.Name)

            Expect.contains funcNames "InnerMsg.TestString" "pass"
            Expect.contains funcNames "TestNameInherited.TestString" "pass"
        )

        testCase "string" (fun _ ->
            testSingleton "Hello world" (OuterMsg.TestString "Hello world")
        )
        testCase "records" (fun _ ->
            test (array2D [["Name"; "Id"]; ["MyName"; "0"]]) (OuterMsg.TestRecords (Records [{Name = "MyName"; Id = 0}]))
        )

        testCase "test table have duplicate keys" (fun _ ->
            let excelRange = 
                ExcelRangeContactInfo.readFromFile (RangeIndexer "A15:K76") (SheetName "Table") book1

            let table = Table.Convert excelRange.Content

            let msg = OuterMsg.TestTable table

            let response: obj [,] =
                client.Value.InvokeFunction msg
                |> Async.RunSynchronously

            let resultStart = response.[0,0].ToString()

            let textStart = "ORDER NO."
            Expect.equal resultStart textStart "pass"

            let resultEnd = response.[9,27].ToString().Substring(0,5)
            let textEnd = "12.53"

            Expect.equal resultEnd textEnd "pass"
        )
    ]


let commandTests =
    testList "command tests" [

        testCase "get commands" (fun _ ->
            /// ensure no exceptions
            let apiLambdas = NetCoreClient.apiLambdas logger client.Value
            let excelCommands = Registration.excelCommands apiLambdas

            Expect.hasLength excelCommands 1 "pass"
            Expect.equal excelCommands.[0].CommandLambda.Name "TestCommand" "pass"

        )

        testCase "test command" (fun _ ->

            let commandCaller = 
                ExcelRangeContactInfo.readFromFile (RangeIndexer "B3") (SheetName "Script") book1
                |> CommandCaller

            let response: Result<obj, string> =
                remoteServer.Value <? OuterMsg.TestCommand commandCaller
                |> Async.RunSynchronously

            match response with 
            | Result.Ok _ -> ()
            | Result.Error error -> failwithf "%s" error 
        )

    ]

