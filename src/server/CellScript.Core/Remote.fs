namespace CellScript.Core
open Akkling
open Types
open System.Reflection
open System
open System.IO
open Akka.Configuration

module Remote = 
    type Client<'Msg> =
        { ClientPort: int 
          ServerPort: int
          RemoteServer: TypedActorSelection<'Msg>}

    let private withFallBack baseConfig =
        Configuration.load().WithFallback baseConfig

    [<RequireQualifiedAccess>]
    module Client =

        let create() =
            let config = 
                sprintf 
                    """
                        akka {
                            remote.port = 10031
                            actor.name = CellScript.Core.Client
                            remote.actor.name = CellScript.Core.Server
                            actor.provider = remote
                            remote.dot-netty.tcp {
                                hostname = localhost
                                port = 0
                            }
                            loglevel = DEBUG
                            loggers=["Akka.Logger.NLog.NLogLogger, Akka.Logger.NLog"]
                        }
                    """ 
                |> Configuration.parse
                |> withFallBack

            let remoteProps = 
                let actor = actorOf ignored
                { props actor with Deploy = None }


            let client = System.create "client" <| config

            let clientPort =
                config.GetInt("akka.remote.dot-netty.tcp.port")

            let clientAgent = 
                let actorName = config.GetString "akka.actor.name"
                spawn client actorName remoteProps

            let remotePort = config.GetInt "akka.remote.port"

            let remotePath = 
                let remoteAddr = sprintf "akka.tcp://server@localhost:%d" remotePort
                let remoteActorName = config.GetString "akka.remote.actor.name"
                sprintf "%s/user/%s" remoteAddr remoteActorName

            let remoteServer = select client remotePath


            { ClientPort = clientPort
              ServerPort = remotePort
              RemoteServer = remoteServer }
        


    [<RequireQualifiedAccess>]
    module Server =
        let createActor handleMsg =
            let config = 
                sprintf 
                    """
                        akka {
                            actor.name = CellScript.Core.Server
                            actor.provider = remote
                            remote.dot-netty.tcp {
                                hostname = localhost
                                port = 10031
                            }
                            loglevel = DEBUG
                            loggers=["Akka.Logger.NLog.NLogLogger, Akka.Logger.NLog"]
                        }
                    """
                |> Configuration.parse
                |> withFallBack
            
            let systemPathName = config.GetString "akka.actor.name"

            let system = System.create "server" <| config
            let server = spawn system systemPathName (props (actorOf2 handleMsg))
            ()


    [<RequireQualifiedAccess>]
    type FcsMsg =
        | Eval of input: SerializableExcelReference * code: string