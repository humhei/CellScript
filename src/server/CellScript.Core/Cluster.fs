namespace CellScript.Core
open Akkling
open Akka.Event
open CellScript.Core.Types
open Shrimp.Akkling.Cluster.Intergraction
open Akka.Configuration
open Shrimp.Akkling.Cluster.Intergraction.Configuration


module Cluster = 
    type private AssemblyFinder = AssemblyFinder

    let internal referenceConfig = 
        lazy 
             ConfigurationFactory.FromResource<AssemblyFinder>("CellScript.Core.reference.conf")
            |> Configuration.fallBackByApplicationConf


    type ExecArguments =
        { ToolName: string 
          Args: string list 
          WorkingDir: string }

    [<RequireQualifiedAccess>]
    module Major = 

        let [<Literal>] private SERVER = "server"
        let [<Literal>] private CELL_SCRIPT_CLUSETER = "cellScriptCoreCluster"

        let [<Literal>] private CLIENT = "client"

        [<RequireQualifiedAccess>]
        type CallbackMsg =
            | UpdateCellValues of xlRefs: ExcelRangeContactInfo list
            | SetRowHeight of xlRef: ExcelRangeContactInfo * height: float
            | Exec of ExecArguments

        [<RequireQualifiedAccess>]
        module private Routed =
            let port = 
                lazy
                    referenceConfig.Value.GetInt("CellScript.Core.Cluster.Major.port")

        [<RequireQualifiedAccess>]
        module Client =
            let create callback: Client<CallbackMsg, 'ServerMsg> =
                Client(CELL_SCRIPT_CLUSETER, CLIENT, SERVER, 0, Routed.port.Value, callback, ( fun args -> {args with ``akka.loggers`` = Loggers (Set.ofList [Logger.NLog])}))

        [<RequireQualifiedAccess>]
        module Server =
            let create receive =
                Server(CELL_SCRIPT_CLUSETER, SERVER, CLIENT, Routed.port.Value, Routed.port.Value, ( fun args -> {args with ``akka.loggers`` = Loggers (Set.ofList [Logger.NLog])}), receive)
            

    [<RequireQualifiedAccess>]
    module COM = 
        let [<Literal>] private CELL_SCRIPT_COM = "cellScriptCoreCOM"
        let [<Literal>] private COM_CLIENT = "client"
        let [<Literal>] private COM_SERVER = "server"

        [<RequireQualifiedAccess>]
        module private Routed =
            let port = 
                lazy
                    referenceConfig.Value.GetInt("CellScript.Core.Cluster.COM.port")

        type ServerMsg =
            | SaveXlsToXlsx of filePaths: string list

        [<RequireQualifiedAccess>]
        module Client =
            let create(): Client<unit, ServerMsg> =
                Client(CELL_SCRIPT_COM, COM_CLIENT, COM_SERVER, 0, Routed.port.Value, Behaviors.ignore, ( fun args -> {args with ``akka.loggers`` = Loggers (Set.ofList [Logger.NLog])}))

        [<RequireQualifiedAccess>]
        module Server =
            let create receive: Server<unit, ServerMsg> =
                Server(CELL_SCRIPT_COM, COM_SERVER, COM_CLIENT, Routed.port.Value, Routed.port.Value, ( fun args -> {args with ``akka.loggers`` = Loggers (Set.ofList [Logger.NLog])}), receive)

