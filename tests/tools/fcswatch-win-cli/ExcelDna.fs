module Fcswatch.Win.Cli.ExcelDna

open Fake.IO
open System.IO
open Fake.IO.FileSystemOperators

open System.Diagnostics
open System.Runtime.InteropServices
open Microsoft.Office.Interop.Excel
open Fake.Core
open FcsWatch.Binary
open ExcelDna.Integration
open System.Xml
open Utils
open FcsWatch.AutoReload.Core.Remote
open Akkling

[<RequireQualifiedAccess>]
module User32 =

    [<DllImport("user32")>]
    extern int GetWindowThreadProcessId(int hwnd, int& lpdwprocessid )

    let getPidFromHwnd(hwnd:int) :int =
        let mutable pid = 0
        GetWindowThreadProcessId(hwnd, &pid) |> ignore
        pid


[<RequireQualifiedAccess>]
type ReloadBy =
    | Xll of project: Project * client: Client<Msg>
    | AddIn

[<RequireQualifiedAccess>]
module ReloadBy =

    let private xllPath (project: Project) =
        let outputPath = List.exactlyOne project.OutputPaths
        let outputDir = Path.getDirectory outputPath
        let xllName = 
            let name = Path.GetFileNameWithoutExtension project.ProjPath
            name + "-AddIn64.xll"
        outputDir </> xllName

    let addInNameStarter projectPath =
        Path.GetFileNameWithoutExtension projectPath


    let uninstall projectPath (app: Application) watcherPostion =
        try 
            match watcherPostion with 
            | ReloadBy.Xll (project, client) ->
                let xll = xllPath project
                let response = 
                    client.RemoteServer <? (Msg.UnregistraterXll xll)
                    |> Async.RunSynchronously

                printfn "unRegistrater xll with result %A" response

            | ReloadBy.AddIn ->
                let addInNameStarter = addInNameStarter projectPath
                let addIn =
                    seq {
                        for addIn in app.AddIns do
                            yield addIn
                    } |> Seq.find (fun addIn -> addIn.Name.StartsWith addInNameStarter)
                addIn.Installed <- false
        with ex ->
            printfn "%A" ex
        Trace.trace "unInstalled plugin"

    let install projectPath (app: Application) watcherPostion =
        try 
            match watcherPostion with
            | ReloadBy.Xll (project, client) ->
                let xll = xllPath project
                let response = 
                    client.RemoteServer <? (Msg.RegistraterXll xll)
                    |> Async.RunSynchronously

                printfn "Registrater xll with result %A" response

            | ReloadBy.AddIn ->
                let addInNameStarter = addInNameStarter projectPath
                let addIn =
                    seq {
                        for addIn in app.AddIns do
                            yield addIn
                    } |> Seq.find (fun addIn -> addIn.Name.StartsWith addInNameStarter)
                addIn.Installed <- true
        with ex -> 
            printfn "%A" ex

        Trace.trace "installed plugin"

let plugin reloadBy developmentTarget (projectPath: string) =
    if not (File.exists projectPath) then failwithf "Cannot find file %s" projectPath

    let projectName = Path.GetFileNameWithoutExtension projectPath
    let projectDir = Path.getDirectory projectPath

    let app =
        let appInbox =
            Process.GetProcesses()
            |> Array.tryFind (fun proc -> proc.ProcessName = "EXCEL")
            |> function
                | Some proc -> Marshal.GetActiveObject("Excel.Application")
                | None ->
                    failwithf "Please manual open excel, and add plugin %s" projectName
            :?> Application


        // let app = new ApplicationClass()
        let workbooks = appInbox.Workbooks
        

        let workBookOp = 
            let xlPath = projectDir </> "datas/book1.xlsx"
            if File.exists xlPath 
            then 
                let names =
                    [ for workbook in workbooks do yield workbook.Name ]

                if not (List.contains "book1.xlsx" names) then
                    workbooks.Open(Filename = xlPath, Editable = true)
                    |> Some
                else None
            else 
                printfn "It's possible place file %s to test" xlPath
                None

        appInbox


    // let app = new ApplicationClass()
    let procId = User32.getPidFromHwnd app.Hwnd


    app.ScreenUpdating <- true

    let install() = ReloadBy.install projectPath app reloadBy

    let unInstall() = ReloadBy.uninstall projectPath app reloadBy

    let calculate() =
        let ws = app.ActiveSheet :?> Worksheet
        try ws.Calculate() 
        with ex -> printfn "%A" ex
        Trace.trace "calculate worksheet"

    let pluginDebugInfo: PluginDebugInfo = 
        { DebuggerAttachTimeDelay = 2000
          Pid = procId 
          VscodeLaunchConfigurationName = "Attach Excel" } 

    match developmentTarget with 
    | DevelopmentTarget.AutoReload _ ->
        let plugin : AutoReload.Plugin = 
            { Load = install
              Unload = unInstall
              Calculate = calculate
              TimeDelayAfterUninstallPlugin = 500
              PluginDebugInfo = Some pluginDebugInfo }

        DevelopmentTarget.autoReloadPlugin plugin
    | DevelopmentTarget.Debuggable _ ->
        let plugin : DebuggingServer.Plugin = 
            { Load = install
              Unload = unInstall
              Calculate = calculate
              TimeDelayAfterUninstallPlugin = 500
              PluginDebugInfo = pluginDebugInfo } 
        DevelopmentTarget.debuggablePlugin plugin
       


