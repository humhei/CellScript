// Learn more about F# at http://fsharp.org

open System
open System.Diagnostics
open Fake.Core
open Fake.IO
open Microsoft.FSharp.Compiler.SourceCodeServices
open FcsWatch.FakeHelper
open ExcelDna.Integration
open Microsoft.Office.Interop.Excel
open System.Runtime.InteropServices
open Newtonsoft.Json
open Fake.IO.FileSystemOperators
open System.IO
open ExcelDna.Integration
open System.Threading
open FcsWatch.Types
open FcsWatch.FcsWatcher
open Fake.DotNet

let root =
    __SOURCE_DIRECTORY__ </> "../../../"
    |> Path.getFullName


let projectPath = root </> "tests" </> "CellScript.Client.Tests"

let projectDir = Path.getDirectory projectPath

let projectName = Path.GetFileNameWithoutExtension projectDir

module VscodeHelper =

    type Configuration =
        { name: string
          ``type``: string
          request: string
          preLaunchTask: obj
          program: obj
          args: obj
          processId: obj
          justMyCode: obj
          cwd: obj
          stopAtEntry: obj
          console: obj }

    type Launch =
        { version: string
          configurations: Configuration list }

    module Launch =
        let file = root </> ".vscode" </> "launch.json"

        let read() =
            let jsonTest = File.readAsString file
            JsonConvert.DeserializeObject<Launch> jsonTest

        let write (launch: Launch) =
            let settings = JsonSerializerSettings(NullValueHandling = NullValueHandling.Ignore)
            let jsonText = JsonConvert.SerializeObject(launch,Formatting.Indented,settings)
            File.writeString false file jsonText


module ExcelLoader =
    open VscodeHelper
    [<DllImport("user32")>]
    extern int GetWindowThreadProcessId(int hwnd, int& lpdwprocessid )

    let getPidFromHwnd(hwnd:int) :int =
        let mutable pid = 0
        GetWindowThreadProcessId(hwnd, &pid) |> ignore
        pid

    module Launch =
        let writePid pid (launch: Launch) =
            { launch with
                configurations =
                    launch.configurations
                    |> List.map (fun configuration ->
                        if configuration.request = "attach" && configuration.name = "Attach Excel" then
                            {configuration with processId = pid}
                        else configuration
                    ) }

open VscodeHelper
open ExcelLoader
open FcsWatch
open FcsWatch.DebuggingServer


[<EntryPoint>]
let main argv =
    // let addInName = "CellScript.Client.Tests-AddIn64"
    let addInNameStarter = projectName
    let app =
        let appInbox =
            Process.GetProcesses()
            |> Seq.tryFind (fun proc -> proc.ProcessName = "EXCEL")
            |> function
                | Some proc -> Marshal.GetActiveObject("Excel.Application")
                | None ->
                    let proc = Process.Start("Excel.exe")
                    Thread.Sleep(3000)
                    Marshal.GetActiveObject("Excel.Application")
            :?> Application

        // let app = new ApplicationClass()
        let procId = getPidFromHwnd appInbox.Hwnd
        let xlPath = projectDir </> "datas/book1.xlsx"
        // let app = new ApplicationClass()
        let workbooks = appInbox.Workbooks
        let workbook = workbooks.Open(Filename = xlPath, Editable = true)

        Launch.read()
        |> Launch.writePid procId
        |> Launch.write

        appInbox


    let addIn =
        seq {
            for addIn in app.AddIns do
                yield addIn
        } |> Seq.find (fun addIn -> addIn.Name.StartsWith addInNameStarter)
    addIn.Installed <- false
    app.Visible <- true

    let projectFile = projectDir </> sprintf "%s.fsproj" projectName

    //DotNet.build (fun ops -> {ops with Configuration = DotNet.BuildConfiguration.Debug }) projectFile

    addIn.Installed <- true

    let checker = FSharpChecker.Create()

    let installPlugin() =
        //let r = DotNet.exec (fun ops -> { ops with WorkingDirectory = projectDir}) "msbuild" "/t:ExcelDnaBuild;ExcelDnaPack"
        //assert (r.OK)
        try 
            addIn.Installed <- true
        with ex -> printfn "%A" ex
        Trace.trace "installed plugin"

    let unInstall() =
        try 
            addIn.Installed <- false
        with ex -> printfn "%A" ex
        Trace.trace "unInstalled plugin"

    let calculate() =
        Trace.trace "calculate worksheet"
        let ws = app.ActiveSheet :?> Worksheet
        ws.Calculate()

    let plugin =
        { Load = installPlugin
          Unload = unInstall
          Calculate = calculate
          DebuggerAttachTimeDelay = 2000 }

    let watcher =
        runFcsWatcherWith (fun config ->
            { config with
                WorkingDir = root
                LoggerLevel = Logger.Level.Normal
                DevelopmentTarget = DevelopmentTarget.Plugin plugin
            }
        ) checker projectFile

    printfn "Hello World from F#!"
    0 // return an integer exit code
