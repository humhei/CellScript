module Tests
open CellScript.Core.Tests
open Akkling
open CellScript.Core.Remote
open Expecto
open CellScript.Core.Types
open System.IO
open System.Reflection
open Fake.IO.FileSystemOperators
open NLog
open CellScript.Client.Core
open CellScript.Client.Desktop



let dir = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location)

let book1 = dir </> "datas/book1.xlsx"

let equal actual expected = Expect.equal expected actual "passed"

let equalByResponse expected response = 
    let array2D: obj[,] = unbox response 
    Expect.equal array2D.[0,0] expected "passed"

let logger = NLog.FSharp.Logger(LogManager.GetCurrentClassLogger())

let createClient<'Msg>() =
    Client.create<'Msg>()


let udfTests =
    let test (expected: obj[,]) (msg: 'Msg) =
        let client = createClient<'Msg>()
        let actor = client.RemoteServer
        let response: obj[,] =
            actor <? msg
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
            let client = createClient<OuterMsg>()
            let apiLambdas = Client.apiLambdas logger client
            let excelFunctions = Registration.excelFunctions apiLambdas
            Expect.hasLength excelFunctions 6 "pass"
        )

        testCase "get fcs functions" (fun _ ->
            /// ensure no exceptions
            let client = createClient<FcsMsg>()
            let apiLambdas = Client.apiLambdas logger client
            let excelFunctions = Registration.excelFunctions apiLambdas
            Expect.hasLength excelFunctions 1 "pass"
        )

        testCase "string" (fun _ ->
            testSingleton "Hello world" (OuterMsg.TestString "Hello world")
        )
    ]


let commandTests =
    testList "command tests" [
        testCase "get command" (fun _ ->
            /// ensure no exceptions
            let client = createClient<FcsMsg>()
            let apiLambdas = Client.apiLambdas logger client
            let excelCommands = Registration.excelCommands apiLambdas

            Expect.hasLength excelCommands 1 "pass"
            Expect.equal excelCommands.[0].CommandLambda.Name "EditCode" "pass"

        )

        ftestCase "invoke edit command" (fun _ ->

            let client = createClient<FcsMsg>()
            let actor = client.RemoteServer
            let commandCaller = 
                SerializableExcelReference.createByFile "B3" "Script" book1
                |> CommandCaller

            let response: Result<unit, string> =
                actor <? FcsMsg.EditCode commandCaller
                |> Async.RunSynchronously

            match response with 
            | Result.Ok _ -> ()
            | Result.Error error -> failwithf "%s" error 
        )

    ]



let evalTests =
    let test (buildExpected: SerializableExcelReference -> obj[,]) inputIndexer codeIndexer =
        let client = createClient<FcsMsg>()
        let actor = client.RemoteServer

        let input = SerializableExcelReference.createByFile inputIndexer "Script" book1
        let expected = buildExpected input
        let code = SerializableExcelReference.createByFile codeIndexer "Script" book1

        let response =
            actor <? FcsMsg.Eval (input, string code.Content.[0,0])
            |> Async.RunSynchronously

        response
        |> Array2D.iteri(fun i j v ->
            equal expected.[i,j] v
        )

    let testSingleton (buildExpected: obj -> 'b) inputIndexer codeIndexer =
        let buildExpected (xlRef: SerializableExcelReference) =
            let input = xlRef.Content.[0,0]
            let result = buildExpected input
            array2D [[box result]]

        test buildExpected inputIndexer codeIndexer

    testList "eval tests" [
    
        testCase "string test" (fun _ ->
            testSingleton (fun a -> string a + "8") "B2" "B3"
        )

        testCase "table" (fun _ ->
            test (fun xlRef ->
                xlRef.Content
            ) "B8:D10" "B12"
        )
    
    ]