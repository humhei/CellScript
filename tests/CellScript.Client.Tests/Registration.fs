module CellScript.Client.Tests.Registration
open ExcelDna.Integration
open ExcelDna.Registration
open ExcelDna.Registration.FSharp
open CellScript.Client.Core
open CellScript.Client.Desktop
open CellScript.Core.Tests
open NLog

let logger = NLog.FSharp.Logger(LogManager.GetCurrentClassLogger())
let client = NetCoreClient.create<OuterMsg> 9050 (fun msg -> msg)
let apiLambdas = NetCoreClient.apiLambdas logger client
let excelFunctions() =
    Registration.excelFunctions apiLambdas
    |> FsAsyncRegistration.ProcessFsAsyncRegistrations
    |> AsyncRegistration.ProcessAsyncRegistrations
    |> ParamsRegistration.ProcessParamsRegistrations
    |> MapArrayFunctionRegistration.ProcessMapArrayFunctions

let installExcelFunctions() =
    let logger = NLog.FSharp.Logger(LogManager.GetCurrentClassLogger())

    logger.Info "Begin install excel functions"
    ExcelIntegration.RegisterUnhandledExceptionHandler(fun ex ->
        let exText = ex.ToString()
        "!!! ERROR: " + exText
        |> box
    )


    let excelFunctions = excelFunctions()

    ExcelRegistration.RegisterFunctions excelFunctions

    ExcelRegistration.GetExcelCommands().RegisterCommands()
    logger.Info "End install excel functions"

type FsAsyncAddIn () =
    interface IExcelAddIn with
        member this.AutoOpen () =
            // appStation.Initial()
            installExcelFunctions()

        member this.AutoClose () =
            ()
