module CellScript.Client.Tests.Registration
open ExcelDna.Integration
open ExcelDna.Registration
open CellScript.Core
open ExcelDna.Registration.FSharp
open CellScript.Client
open CellScript.Core.Tests

[<AutoOpen>]
module internal Global =
    let logger = Logger.create Logger.Level.Normal


let client: Client<OuterMsg> = 
    { ClientPort = 6000 
      ServerPort = 6001 }

let excelFunctions() =
    Registration.excelFunctions client
    |> FsAsyncRegistration.ProcessFsAsyncRegistrations
    |> AsyncRegistration.ProcessAsyncRegistrations
    |> ParamsRegistration.ProcessParamsRegistrations
    |> MapArrayFunctionRegistration.ProcessMapArrayFunctions

let installExcelFunctions() =
    logger.Diagnostics "Begin install excel functions"
    ExcelIntegration.RegisterUnhandledExceptionHandler(fun ex ->
        let exText = ex.ToString()
        "!!! ERROR: " + exText
        |> box
    )
    let excelFunctions = excelFunctions()

    ExcelRegistration.RegisterFunctions excelFunctions

    ExcelRegistration.GetExcelCommands().RegisterCommands()
    logger.Diagnostics "End install excel functions"

type FsAsyncAddIn () =

    interface IExcelAddIn with
        member this.AutoOpen () =
            // appStation.Initial()
            installExcelFunctions()
        member this.AutoClose () =
            ()
