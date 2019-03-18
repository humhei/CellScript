module CellScript.Client.Tests.Registration
open ExcelDna.Integration
open ExcelDna.Registration
open CellScript.Core
open CellScript.Client.UDF
open ExcelDna.Registration.FSharp
open CellScript.Client

let installExcelFunctions() =
    logger.Diagnostics "Begin install excel functions"
    ExcelIntegration.RegisterUnhandledExceptionHandler(fun ex ->
        let exText = ex.ToString()
        "!!! ERROR: " + exText
        |> box
    )
    Registration.excelFunctions()
    |> Seq.append (ExcelRegistration.GetExcelFunctions())
    |> List.ofSeq
    |> fun fns -> ParameterConversionRegistration.ProcessParameterConversions (fns, Registration.paramConvertConfig)
    |> FsAsyncRegistration.ProcessFsAsyncRegistrations
    |> AsyncRegistration.ProcessAsyncRegistrations
    |> ParamsRegistration.ProcessParamsRegistrations
    |> MapArrayFunctionRegistration.ProcessMapArrayFunctions
    |> ExcelRegistration.RegisterFunctions

    ExcelRegistration.GetExcelCommands().RegisterCommands()
    logger.Diagnostics "End install excel functions"

type FsAsyncAddIn () =

    interface IExcelAddIn with
        member this.AutoOpen () =
            // appStation.Initial()
            installExcelFunctions()
        member this.AutoClose () =
            ()
