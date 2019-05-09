namespace CellScript.Core
open Akkling
open Types
open System
open Microsoft.FSharp.Quotations.Patterns
open Akka.Remote
open Akka.Event
open Akka.Cluster.Tools.Singleton
open Akka.Cluster

module Cluster = 
    open AkkaExtensions

    [<AutoOpen>]
    module Types =
        type Command =
            { Shortcut: string option }

        type CellScriptEvent<'EventArg> = CellScriptEvent of 'EventArg

        type SheetActiveArg =
            { WorkbookPath: string 
              WorksheetName: string }


        [<RequireQualifiedAccess>]
        type FcsMsg =
            | EditCode of CommandCaller
            | Eval of input: SerializableExcelReference * code: string
            | Sheet_Active of CellScriptEvent<SheetActiveArg>
            | Workbook_BeforeClose of CellScriptEvent<string>


    let private withAppconfigOrWebconfig baseConfig =
        Configuration.load().WithFallback baseConfig

    let [<Literal>] private SERVER = "server"
    let [<Literal>] private CELL_SCRIPT_CLUSETER = "cellScriptCluster"
    //let [<Literal>] private CELL_SCRIPT_CLUSETER_NODE = "cellScriptClusterClient"

    [<RequireQualifiedAccess>]
    module private Hopac =
        let listText (lists: string list) =
            lists
            |> List.map (sprintf "\"%s\"")
            |> String.concat ","

    module ClusterOperators =

        let clusterConfig (roles: string list) remotePort seedPort = 
            /// ["crawler", "logger"]

            let config = 
                sprintf 
                    """
                    akka {
                        actor {
                          provider = cluster
                          serializers {
                            hyperion = "Akka.Serialization.HyperionSerializer, Akka.Serialization.Hyperion"
                          }
                          serialization-bindings {
                            "System.Object" = hyperion
                          }
                        }
                      remote {
                        dot-netty.tcp {
                          public-hostname = "localhost"
                          hostname = "localhost"
                          port = %d
                        }
                      }
                      cluster {
                        auto-down-unreachable-after = 5s
                        seed-nodes = [ "akka.tcp://%s@localhost:%d/" ]
                        roles = [%s]
                      }
                      persistence {
                        journal.plugin = "akka.persistence.journal.inmem"
                        snapshot-store.plugin = "akka.persistence.snapshot-store.local"
                      }
                    loglevel = DEBUG
                    loggers=["Akka.Logger.NLog.NLogLogger, Akka.Logger.NLog"]	
                    }
                    """ remotePort CELL_SCRIPT_CLUSETER seedPort (Hopac.listText roles)
                |> Configuration.parse

            config.WithFallback(ClusterSingletonManager.DefaultConfig())
            |> withAppconfigOrWebconfig


        let createSystem roles remotePort seedPort = System.create CELL_SCRIPT_CLUSETER <| clusterConfig roles remotePort seedPort


    open ClusterOperators

    [<AutoOpen>]
    module Client =

        let rec asCommandUci expr =
            match expr with 
            | Lambda(_, expr) ->
                asCommandUci expr

            | NewUnionCase (uci, exprs) ->
                if exprs.Length <> 1 then failwithf "command %A msg's paramters length should be equal to 1" uci
                else uci
            | _ -> failwithf "command expr %A should be a union case type" expr

        let commandMapping =
            lazy 
                [ asCommandUci <@ FcsMsg.EditCode @>, { Shortcut = Some "+^{F11}" }]
                |> dict


        [<RequireQualifiedAccess>]
        type ServerMsg<'ServerCustomMsg> =
            | CustomMsg of 'ServerCustomMsg

        [<RequireQualifiedAccess>]
        type ClientMsg =
            | UpdateCellValues of xlRefs: SerializableExcelReference list
            | Exec of toolName: string * args: string list * workingDir: string


        type Client<'ServerCustomMsg> =
            { RemoteServer: ICanTell<'ServerCustomMsg>
              Logger: ILoggingAdapter }

        let [<Literal>] CLIENT = "client"

        [<RequireQualifiedAccess>]
        module Client =

            let create seedPort (handleClientMsg: ILoggingAdapter -> ClientMsg -> ActorEffect<ClientMsg>): Client<'ServerCustomMsg> =
                
                let system = createSystem [CLIENT] 0 seedPort

                let log = system.Log

                let callbackActor = spawnAnonymous system (props (actorOf(handleClientMsg log)))

                let remoteServer: ICanTell<'ServerCustomMsg> = 

                    ClusterClient.create system CLIENT SERVER (Some callbackActor)
                    |> retypeToDUChild ServerMsg.CustomMsg

                { RemoteServer = remoteServer
                  Logger = log }

    [<AutoOpen>]
    module Server =


        type ServerModel<'CustomModel> =
            { CustomModel: 'CustomModel }

        type ContantClientMsg =
            | AddServer

        [<RequireQualifiedAccess>]
        module Server =

            type HandleMsg<'ServerCustomMsg, 'ServerCustomModel> = 
                Actor<ServerMsg<'ServerCustomMsg>> 
                    -> 'ServerCustomMsg 
                    -> 'ServerCustomModel
                    -> 'ServerCustomModel

            let createAgent seedport initialCustomModel (handleMsg: HandleMsg<'ServerCustomMsg, 'ServerCustomModel>) =
                
                let system = createSystem [SERVER] seedport seedport
                let server = spawn system SERVER (props (fun ctx ->
                    let rec loop model = actor {
                        let! msg = ctx.Receive()
                        match msg with 
                        | ServerMsg.CustomMsg customMsg ->
                            let newCustomModel = handleMsg ctx customMsg model.CustomModel
                            return! loop { model with CustomModel = newCustomModel }
                    }

                    loop { CustomModel = initialCustomModel }
                ))

                ClusterServerStatusListener.listenServer server system CLIENT

                ()
            







