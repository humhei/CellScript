module CellScript.Client.Fcs.Registration
open ExcelDna.Integration
open ExcelDna.Registration
open ExcelDna.Registration.FSharp
open CellScript.Client.Core
open NLog
open CellScript.Core.Cluster
open CellScript.Core.Types
open System.Reflection
open Microsoft.Office.Interop.Excel
open Akkling
open CellScript.Client.Desktop
open CellScript.Core

let logger = NLog.FSharp.Logger(LogManager.GetCurrentClassLogger())

let client = NetCoreClient.create<FcsMsg> 9060
let apiLambdas = NetCoreClient.apiLambdas logger client

let excelFunctions() =
    Registration.excelFunctions apiLambdas
    |> FsAsyncRegistration.ProcessFsAsyncRegistrations
    |> AsyncRegistration.ProcessAsyncRegistrations
    |> ParamsRegistration.ProcessParamsRegistrations
    |> MapArrayFunctionRegistration.ProcessMapArrayFunctions

let installPlugin() =
    let currentAssemblyLocation = Assembly.GetExecutingAssembly().Location

    logger.Info "Begin install plugin %s" currentAssemblyLocation
    ExcelIntegration.RegisterUnhandledExceptionHandler(fun ex ->
        let exText = "!!! ERROR: " + ex.ToString()
        
        logger.Error "%s" exText

        exText
        |> box
    )

    let excelFunctions = excelFunctions()

    ExcelRegistration.RegisterFunctions excelFunctions
    (Registration.excelCommands apiLambdas).RegisterCommands()
    logger.Info "End install plugin %s" currentAssemblyLocation

    let app = ExcelDnaUtil.Application :?> Application

    let sheetEventOfWorkbook (workbook: Workbook) =
        let sheet = workbook.ActiveSheet :?> Worksheet

        { WorkbookPath = workbook.Path 
          WorksheetName = sheet.Name }
        |> CellScriptEvent

    app.add_WorkbookOpen(fun workbook ->
        client.RemoteServer <! FcsMsg.Sheet_Active (sheetEventOfWorkbook workbook)
    )


    app.add_WorkbookBeforeClose(fun workbook _ ->
        let event = CellScriptEvent workbook.Path
        client.RemoteServer <! FcsMsg.Workbook_BeforeClose event
    )

    app.SheetActivate.Add(fun _ ->
        let app = ExcelDnaUtil.Application :?> Application
        client.RemoteServer <! FcsMsg.Sheet_Active (sheetEventOfWorkbook app.ActiveWorkbook)
    )


type FsAsyncAddIn () =
    interface IExcelAddIn with
        member this.AutoOpen () =
            // appStation.Initial()
            installPlugin()

        member this.AutoClose () =
            ()
