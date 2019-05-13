namespace CellScript.Core
open Akkling
open System
open Microsoft.FSharp.Quotations.Patterns
open Akka.Remote
open Akka.Event
open Akka.Cluster.Tools.Singleton
open Akka.Cluster
open CellScript.Core.Types
open Akkling.Cluster.ServerSideDevelopment
open Akkling.Cluster.ServerSideDevelopment.Client

module Cluster = 



    let [<Literal>] private SERVER = "server"
    let [<Literal>] private CELL_SCRIPT_CLUSETER = "cellScriptCluster"


    [<RequireQualifiedAccess>]
    module Config = 

        [<RequireQualifiedAccess>]
        module private Hopac =
            let listText (lists: string list) =
                lists
                |> List.map (sprintf "\"%s\"")
                |> String.concat ","


        let internal withAppconfigOrWebconfig baseConfig =
            Configuration.load().WithFallback baseConfig

        let createClusterConfig (roles: string list) remotePort seedPort = 
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


    let private createClusterSystem roles remotePort seedPort = 
        System.create CELL_SCRIPT_CLUSETER <| Config.createClusterConfig roles remotePort seedPort

    let [<Literal>] private CLIENT = "client"

    [<RequireQualifiedAccess>]
    type ClientCallbackMsg =
        | UpdateCellValues of xlRefs: SerializableExcelReference list
        | Exec of toolName: string * args: string list * workingDir: string

    type Client<'ServerMsg> =
        { RemoteServer: ICanTell<'ServerMsg>
          Logger: ILoggingAdapter }

    [<RequireQualifiedAccess>]
    module Client =

        let createAgent seedPort (handleClientCallback: ILoggingAdapter -> ClientCallbackMsg -> ActorEffect<ClientCallbackMsg>): Client<'ServerMsg> =
            let system = createClusterSystem [CLIENT] 0 seedPort

            let log = system.Log

            let callbackActor = spawnAnonymous system (props (actorOf(handleClientCallback log)))

            let remoteServer: ICanTell<'ServerMsg> = 

                ClientAgent.create system CLIENT SERVER (Some callbackActor)

            { RemoteServer = remoteServer
              Logger = log }

    [<RequireQualifiedAccess>]
    module Server =

        let createAgent seedport initialCustomModel (handleMsg) =
            let system = createClusterSystem [SERVER] seedport seedport
            Server.createAgent (fun (c: ClientCallbackMsg) -> ()) system SERVER CLIENT initialCustomModel handleMsg
            system
            







