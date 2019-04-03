module CellScript.Client.Tests.Registration
open ExcelDna.Integration
open ExcelDna.Registration
open CellScript.Core
open ExcelDna.Registration.FSharp
open CellScript.Client

let excelFunctions() =
    Registration.excelFunctions()
    |> List.ofSeq
    |> fun fns -> 
        let s = ParameterConversionRegistration.ProcessParameterConversions (fns, Registration.paramConvertConfig)
        s
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
