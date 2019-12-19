module Tests
open Akkling
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
open CellScript.Core.Extensions

let dir = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location)

let book1 = dir </> "datas/book1.xlsx"

let equal actual expected = Expect.equal expected actual "passed"

let equalByResponse expected response = 
    let array2D: obj[,] = unbox response 
    Expect.equal array2D.[0,0] expected "passed"

let logger = NLog.FSharp.Logger(LogManager.GetCurrentClassLogger())

let createClient<'Msg>() =
    let client = NetCoreClient.create<'Msg> Route.port id
    Thread.Sleep(2000)
    client

let client = createClient<FcsMsg>()

let remoteServer = client.RemoteServer


let commandTests =
    testList "command tests" [

        testCase "get commands" (fun _ ->
            /// ensure no exceptions
            let apiLambdas = NetCoreClient.apiLambdas logger client
            let excelCommands = Registration.excelCommands apiLambdas

            Expect.hasLength excelCommands 1 "pass"
            Expect.equal excelCommands.[0].CommandLambda.Name "EditCode" "pass"

        )

        testCase "invoker edit code" (fun _ ->

            let commandCaller = 
                ExcelRangeContactInfo.readFromFile (RangeGettingArg.RangeIndexer "B3") (SheetGettingArg.SheetName "Script") book1
                |> CommandCaller

            let response: Result<obj, string> =
                remoteServer <? FcsMsg.EditCode commandCaller
                |> Async.RunSynchronously

            match response with 
            | Result.Ok _ -> ()
            | Result.Error error -> failwithf "%s" error 
        )

    ]


let evalTests =
    let test (buildExpected: ExcelRangeContactInfo -> obj[,]) inputIndexer codeIndexer =
        let actor = client.RemoteServer

        let input = ExcelRangeContactInfo.readFromFile (RangeGettingArg.RangeIndexer inputIndexer) (SheetGettingArg.SheetName "Script") book1
        let expected = buildExpected input
        let code = ExcelRangeContactInfo.readFromFile (RangeGettingArg.RangeIndexer codeIndexer) (SheetGettingArg.SheetName "Script") book1

        let response =
            client.InvokeFunction (FcsMsg.Eval (input, string code.Content.[0,0]))
            |> Async.RunSynchronously

        response
        |> Array2D.iteri(fun i j v ->
            equal expected.[i,j] v
        )

    let testSingleton (buildExpected: obj -> 'b) inputIndexer codeIndexer =
        let buildExpected (xlRef: ExcelRangeContactInfo) =
            let input = xlRef.Content.[0,0]
            let result = buildExpected input
            array2D [[box result]]

        test buildExpected inputIndexer codeIndexer

    testList "eval tests" [

        testCase "get fcs functions" (fun _ ->
            /// ensure no exceptions
            let apiLambdas = NetCoreClient.apiLambdas logger client
            let excelFunctions = Registration.excelFunctions apiLambdas
            Expect.hasLength excelFunctions 1 "pass"
        )

        testCase "string test" (fun _ ->
            testSingleton (fun a -> string a + "8") "B2" "B3"
        )

        testCase "table" (fun _ ->
            test (fun xlRef ->
                xlRef.Content
            ) "B8:D10" "B12"
        )
    ]