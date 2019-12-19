module CellScript.Server.Fcs.Watcher
open Akkling
open System
open CellScript.Core.Cluster
open CellScript.Core.Types
open Fake.IO
open System.IO
open System.Threading
open System.Timers
open Fake.IO.Globbing
open System.Text
open CellScript.Core
open Shrimp.Akkling.Cluster.Intergraction
open Akka.Actor


[<RequireQualifiedAccess>]
module private IntervalAccumMailBoxProcessor =
    type State<'accum> =
        { Timer: Timer option
          Accums: 'accum list }

    type Msg<'accum> =
        | Accum of 'accum
        | IntervalUp

    type IntervalAccumMailBoxProcessor<'accum> =
        { Agent: MailboxProcessor<'accum>
          CancellationTokenSource: CancellationTokenSource }

    let create (onChange : 'accum seq -> unit) =

        let cancellationTokenSource = new CancellationTokenSource()
        let agent =
            MailboxProcessor<Msg<'accum>>.Start ((fun inbox ->

                let createTimer() =
                    let timer = new Timer(100.)
                    timer.Start()
                    timer.Elapsed.Add (fun _ ->
                        timer.Stop()
                        timer.Dispose()
                        inbox.Post IntervalUp
                    )
                    timer

                let rec loop state = async {
                    let! msg = inbox.Receive()
                    match msg with
                    | Accum accum ->
                        match state.Timer with
                        | Some timer ->
                            timer.Stop()
                            timer.Dispose()
                        | None -> ()

                        return!
                            loop
                                { state with
                                    Timer = Some (createTimer())
                                    Accums = accum :: state.Accums }

                    | IntervalUp ->
                        onChange state.Accums
                        return! loop { state with Timer = None; Accums = [] }
                }

                loop { Timer = None; Accums = [] }
            ),cancellationTokenSource.Token)

        { CancellationTokenSource = cancellationTokenSource
          Agent = agent }

    let accum value = Msg.Accum value

[<RequireQualifiedAccess>]
module private ChangeWatcher =

    type Options = ChangeWatcher.Options

    let private handleWatcherEvents (status : FileStatus) (onChange : FileChange -> unit) (e : FileSystemEventArgs) =
        onChange ({ FullPath = e.FullPath
                    Name = e.Name
                    Status = status })

    let runWithAgent (foptions:Options -> Options) (onChange : FileChange seq -> unit) (fileIncludes : IGlobbingPattern) =
        let options = foptions { IncludeSubdirectories = true }
        let dirsToWatch = fileIncludes |> GlobbingPattern.getBaseDirectoryIncludes

        let intervalAccumAgent = IntervalAccumMailBoxProcessor.create onChange

        let postFileChange (fileChange: FileChange) =
            if fileIncludes.IsMatch fileChange.FullPath
            then intervalAccumAgent.Agent.Post (IntervalAccumMailBoxProcessor.accum fileChange)

        let watchers =
            dirsToWatch |> List.map (fun dir ->
                               //tracefn "watching dir: %s" dir

                               let watcher = new FileSystemWatcher(Path.getFullName dir, "*.*")
                               watcher.EnableRaisingEvents <- true
                               watcher.IncludeSubdirectories <- options.IncludeSubdirectories
                               watcher.Changed.Add(handleWatcherEvents Changed postFileChange)
                               watcher.Created.Add(handleWatcherEvents Created postFileChange)
                               watcher.Deleted.Add(handleWatcherEvents Deleted postFileChange)
                               watcher.Renamed.Add(fun (e : RenamedEventArgs) ->
                                   postFileChange { FullPath = e.OldFullPath
                                                    Name = e.OldName
                                                    Status = Deleted }
                                   postFileChange { FullPath = e.FullPath
                                                    Name = e.Name
                                                    Status = Created })
                               watcher)

        { new IDisposable with
              member this.Dispose() =
                  for watcher in watchers do
                      watcher.EnableRaisingEvents <- false
                      watcher.Dispose()
                  // only dispose the timer if it has been constructed
                  intervalAccumAgent.CancellationTokenSource.Cancel()
                  intervalAccumAgent.CancellationTokenSource.Dispose() }

    let run (onChange : FileChange seq -> unit) (fileIncludes : IGlobbingPattern) = runWithAgent id onChange fileIncludes


[<RequireQualifiedAccess>]
type CellScriptWatcherMsg =
    | EditCode of ExcelRangeContactInfo
    | Sheet_Active of SheetReference
    | DetectSourceFileChanges of FileChange seq
    | UpdateActiveCells of Map<SerializableExcelReference, string>



type CellScriptWatcherModel =
    { ChangeWatcher: IDisposable option
      ActivedCells: Map<SerializableExcelReference, string>
      Clients: Map<Address, ClusterActor<ClientCallbackMsg>> }


let createCellScriptWatcher scriptsDir (logger: NLog.FSharp.Logger) system : IActorRef<CellScriptWatcherMsg> =
    spawnAnonymous system (props(fun ctx ->

        let tryDisposeChangeWatcher model =
            match model.ChangeWatcher with
            | Some changeWatcher -> changeWatcher.Dispose()
            | None -> ()

            { model with ChangeWatcher = None }



        let rec loop model  = actor {
            let updateWatcher (workbookPath, sheetName) =
                let newModel = tryDisposeChangeWatcher model
                let files =
                    model.ActivedCells
                    |> Map.filter(fun xlRef _ ->
                        xlRef.WorkbookPath = workbookPath
                        && xlRef.SheetName = sheetName
                    )
                    |> Seq.map (fun pair -> pair.Value)

                let pattern =
                    { BaseDirectory = scriptsDir
                      Includes = List.ofSeq files
                      Excludes = [] }

                let newChangeWatcher =
                    ChangeWatcher.run
                        (fun fileChanges -> ctx.Self <! box (CellScriptWatcherMsg.DetectSourceFileChanges fileChanges))
                        pattern
                { newModel with ChangeWatcher = Some newChangeWatcher }

            let! msg = ctx.Receive() : IO<obj>
            logger.Info "cellscript watcher receive %A" msg

            match msg with
            | :? UpdateCallbackClientsEvent<ClientCallbackMsg> as updateClientEvent ->
                let (UpdateCallbackClientsEvent clients) = updateClientEvent
                return! loop { model with Clients = clients }

            | :? CellScriptWatcherMsg as cellScriptWatcherMsg ->
                match cellScriptWatcherMsg with
                | CellScriptWatcherMsg.EditCode xlRef ->

                    if model.Clients.Count <> 1 then
                        logger.Warn "sprintf callback clients's length %d is not euqal to 1" model.Clients.Count

                    model.Clients
                    |> Map.iter (fun add client ->
                        client <! ClientCallbackMsg.SetRowHeight (xlRef,14.4)
                    )

                    return! loop (updateWatcher(xlRef.WorkbookPath, xlRef.SheetName))

                | CellScriptWatcherMsg.Sheet_Active arg ->
                    return! loop (updateWatcher(arg.WorkbookPath, arg.SheetName))

                | CellScriptWatcherMsg.UpdateActiveCells cells ->
                    return! loop { model with ActivedCells = cells }

                | CellScriptWatcherMsg.DetectSourceFileChanges fileChanges ->
                    let xlRefs =
                        fileChanges
                        |> List.ofSeq
                        |> List.distinctBy (fun fileChange -> fileChange.FullPath)
                        |> List.map (fun fileChange ->
                            let contents = File.readAsStringWithEncoding Encoding.UTF8 fileChange.FullPath
                            let xlRef =
                                model.ActivedCells
                                |> Map.findKey(fun _ fsxFile -> fsxFile = fileChange.FullPath)
                            SerializableExcelReference.toExcelRangeContactInfo (array2D [[contents]]) xlRef
                        )

                    if model.Clients.Count <> 1 then
                        logger.Warn "sprintf callback clients's length %d is not euqal to 1" model.Clients.Count

                    model.Clients
                    |> Map.iter (fun add client ->
                        client <! ClientCallbackMsg.UpdateCellValues xlRefs
                    )

                    return! loop model
            | _ -> return Unhandled
        }

        loop { ChangeWatcher = None; ActivedCells = Map.empty; Clients = Map.empty }
    )) |> retype
