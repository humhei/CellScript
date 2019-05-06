module CellScript.Server.Fcs.RemoteServer
open Akkling
open CellScript.Core.Remote
open CellScript.Core.Types
open Fake.IO
open System.IO
open System
open System.Text
open Watcher
open CellScript.Server.Fcs

#nowarn "0044"





type Config = 
    { ServerPort: int }

let ok: Result<unit, string> = Result.Ok ()

//let private nlogConfig = 
//    Configuration.parse """
//            akka {
//                loglevel = DEBUG
//                loggers=["Akka.Logger.NLog.NLogLogger, Akka.Logger.NLog"]
//            }
//        """ 

let private snapshotConfig = 
    Configuration.parse """
akka.persistence.snapshot-store.plugin = "akka.persistence.snapshot-store.litedb.fsharp"
akka.persistence.snapshot-store.litedb.fsharp {
    class = "Akka.Persistence.LiteDB.FSharp.LiteDBSnapshotStore, Akka.Persistence.LiteDB.FSharp"
    plugin-dispatcher = "akka.actor.default-dispatcher"
    connectString = "hello.db"
}
        """


let run (logger: NLog.FSharp.Logger) =
    #if Release
    let port =  9049
    #else 
    //let port = 9049
    let port = 9050
    #endif


    let remoteConfig = 
        snapshotConfig.WithFallback(Server.fetchConfig port)

    let remoteSystem = 
        Server.createSystem remoteConfig

    //let localSystem = 
    //    System.create "cellScriptServerLocalHost" (nlogConfig.WithFallback snapshotConfig)


    let scriptsDir = Path.Combine(Directory.GetCurrentDirectory(), "Scripts")

    Directory.CreateDirectory(scriptsDir)
    |> ignore

    let cellScriptWatcher = createCellScriptWatcher scriptsDir logger remoteSystem

    let actionAfterUpdateClients clients =
        cellScriptWatcher <! CellScriptWatcherMsg.UpdateClients clients

    let compilerAgent = Compiler.createAgent remoteSystem

    Server.createRemoteActor remoteConfig remoteSystem Map.empty actionAfterUpdateClients (fun ctx msg clients activeCells ->
    
        logger.Info "cellscript fcs server receive %A" msg
        match msg with 

        | FcsMsg.Eval (xlRef, code) ->
            let result = 
                compilerAgent <? (CompilerMsg.Eval (scriptsDir, xlRef, code))
                |> Async.RunSynchronously
            ctx.Sender() <! result
            activeCells

        | FcsMsg.EditCode (CommandCaller xlRef) ->
            let editFsxFileIncode xlRef =

                let fsxFile = SerializableExcelReferenceWithoutContent.getFsxPath scriptsDir (SerializableExcelReferenceWithoutContent.ofSerializableExcelReference xlRef)
                let code = string xlRef.Content.[0,0]

                File.writeStringWithEncoding Encoding.UTF8 false fsxFile code
                let sender = ctx.Sender()

                clients |> List.iter (fun client ->
                    client <! ClientMsg.Exec ("code", [scriptsDir.TrimEnd('\\').TrimEnd('/'); fsxFile], scriptsDir )
                )

            let result =
                try 
                    editFsxFileIncode xlRef
                    ok
                with ex ->
                    Result.Error ex.Message

            ctx.Sender() <! result

            let newActiveCells =
                let withoutContent = SerializableExcelReferenceWithoutContent.ofSerializableExcelReference xlRef
                activeCells.Add (withoutContent, (SerializableExcelReferenceWithoutContent.getFsxPath scriptsDir withoutContent))

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
    |> ignore
