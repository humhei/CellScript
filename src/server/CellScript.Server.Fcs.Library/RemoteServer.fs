module CellScript.Server.Fcs.RemoteServer
open Akkling
open Shrimp.Akkling.Cluster.Intergraction
open CellScript.Core.Cluster
open CellScript.Core.Types
open Fake.IO
open System.IO
open System
open System.Text
open Watcher
open CellScript.Server.Fcs
open CellScript.Core
open Fake.IO.FileSystemOperators

#nowarn "0044"


type Config = 
    { ServerPort: int }

let ok: Result<unit, string> = Result.Ok ()

let [<Literal>] defaultTemplate = """#load "..\\Packages.fsx"                                         
open CellScript.Core
fun (table: Table) ->
    table"""


//let private snapshotConfig = 
//    Configuration.parse """
//akka.persistence.snapshot-store.plugin = "akka.persistence.snapshot-store.litedb.fsharp"
//akka.persistence.snapshot-store.litedb.fsharp {
//    class = "Akka.Persistence.LiteDB.FSharp.LiteDBSnapshotStore, Akka.Persistence.LiteDB.FSharp"
//    plugin-dispatcher = "akka.actor.default-dispatcher"
//    connection-string = "hello.db"
//}
//        """


let run (logger: NLog.FSharp.Logger) =

    let scriptsDir = Path.Combine(Directory.GetCurrentDirectory(), "Scripts")
    Directory.ensure scriptsDir

    let system = 
        let nlogConfig = 
            Configuration.parse """
                    akka {
                        loglevel = DEBUG
                        loggers=["Akka.Logger.NLog.NLogLogger, Akka.Logger.NLog"]
                    }
                """ 

        System.create "localSystem" nlogConfig

    let compilerAgent = Compiler.createAgent system

    let cellScriptWatcher = createCellScriptWatcher scriptsDir logger system

    let clusterSystem =
        Server.createAgent Route.port Map.empty (fun ctx msg activeCells ->
            let logger = ctx.Log.Value
            logger.Info (sprintf "cellscript fcs server receive %A" msg)
            match msg with 

            | FcsMsg.Eval (xlRef, code) ->

                let result = 
                    compilerAgent <? (CompilerMsg.Eval (scriptsDir, xlRef, code))
                    |> Async.RunSynchronously

                let sender = ctx.Sender()
                sender <! result
                activeCells

            | FcsMsg.EditCode (CommandCaller xlRef) ->
                let editFsxFileIncode xlRef =

                    let fsxFile = SerializableExcelReference.getFsxPath scriptsDir (SerializableExcelReference.ofExcelRangeContactInfo xlRef)
                    let code = 
                        let content = xlRef.Content.[0,0]
                        let contentText =
                            match content with 
                            | :? ExcelErrorEnum as errorEnum ->
                                match errorEnum with 
                                | ExcelErrorEnum.Null -> ""
                                | _ -> errorEnum.ToString()
                            | _ -> content.ToString()

                        if contentText.Trim() = "" then 
                            let template = scriptsDir </> "template.fsx"
                            if File.exists template then 
                                File.readAsStringWithEncoding Encoding.UTF8 template
                            else defaultTemplate
                        else contentText

                    File.writeStringWithEncoding Encoding.UTF8 false fsxFile code
                    let sender = ctx.Sender()
                    
                    //let editor = Configuration.load().GetString("akka.editor", @"C:\Program Files (x86)\Microsoft Visual Studio\2019\Community\Common7\IDE\devenv.exe")
                    let editor = Configuration.load().GetString("akka.editor", @"code")
                    
                    let toolName = Path.GetFileNameWithoutExtension editor
                    if String.Compare(toolName, "code", true) = 0 then
                        sender <! ClientCallbackMsg.Exec (editor, [scriptsDir.TrimEnd('\\').TrimEnd('/'); fsxFile], scriptsDir )
                    elif String.Compare(toolName, "devenv", true) = 0 then
                        sender <! ClientCallbackMsg.Exec (editor, ["/Edit"; fsxFile], scriptsDir)
                    else failwithf "Unkonwn editor %s" editor

                let result =
                    try 
                        editFsxFileIncode xlRef
                        ok
                    with ex ->
                        logger.Error ex.Message
                        Result.Error ex.Message

                ctx.Sender() <! result

                let newActiveCells =
                    let withoutContent = SerializableExcelReference.ofExcelRangeContactInfo xlRef
                    activeCells.Add (withoutContent, (SerializableExcelReference.getFsxPath scriptsDir withoutContent))

                cellScriptWatcher <! CellScriptWatcherMsg.UpdateActiveCells newActiveCells
                cellScriptWatcher <! CellScriptWatcherMsg.EditCode xlRef
                newActiveCells

            | FcsMsg.Sheet_Active (CellScriptEvent arg) ->
                cellScriptWatcher <! CellScriptWatcherMsg.Sheet_Active arg
                activeCells

            | FcsMsg.Workbook_BeforeClose (CellScriptEvent workbook) ->
                let newActiveCells = 
                    activeCells
                    |> Map.filter (fun xlRef _ -> xlRef.WorkbookPath <> workbook)

                cellScriptWatcher <! CellScriptWatcherMsg.UpdateActiveCells newActiveCells

                newActiveCells
        )

    clusterSystem.EventStream.Subscribe((cellScriptWatcher :> IInternalTypedActorRef).Underlying, typeof<UpdateCallbackClientsEvent<ClientCallbackMsg>>)
    |> ignore