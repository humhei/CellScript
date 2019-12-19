namespace FcsWatch.AutoReload.ExcelDna.Core

open Akkling

module Remote = 
    type Client<'Msg> =
        { ClientPort: int 
          ServerPort: int
          RemoteServer: TypedActorSelection<'Msg>}

    let [<Literal>] private serverSystemName = "FcsWatchAutoReloadExcelDnaCoreServer"
    let [<Literal>] private clientSystemName = "FcsWatchAutoReloadExcelDnaCoreClient"
    let [<Literal>] private SERVER = "server"
    let [<Literal>] private CLIENT = "client"

    [<RequireQualifiedAccess>]
    module Client =

        let create clientPort serverPort =
            let config = 
                sprintf 
                    """
                        akka {
                            actor.provider = remote
                            remote.dot-netty.tcp {
                                hostname = localhost
                                port = %d
                            loglevel = DEBUG
                            loggers=["Akka.Logger.NLog.NLogLogger, Akka.Logger.NLog"]	
                            }
                        }
                    """ clientPort

            let serverAddr = sprintf "akka.tcp://%s@localhost:%d" serverSystemName serverPort

            let remoteProps = 
                let actor = actorOf ignored
                { props actor with Deploy = None }


            let client = System.create clientSystemName <| Configuration.parse config

            let clientAgent = spawn client CLIENT remoteProps

            let serverPath = sprintf "%s/user/%s" serverAddr SERVER

            let remoteServer = select client serverPath


            { ClientPort = clientPort
              ServerPort = serverPort
              RemoteServer = remoteServer }
        


    [<RequireQualifiedAccess>]
    module Server =
        let createActor port handleMsg =
                let serverConfig = 
                    sprintf 
                        """
                            akka {
                                actor.provider = remote
                                remote.dot-netty.tcp {
                                    hostname = localhost
                                    port = %d
                                }
                            loglevel = DEBUG
                            loggers=["Akka.Logger.NLog.NLogLogger, Akka.Logger.NLog"]
                            }
                        """ port


                let system = System.create serverSystemName <| Configuration.parse serverConfig
                let server = spawn system SERVER (props (actorOf2 handleMsg))
                ()


    [<RequireQualifiedAccess>]
    type Msg =
        | RegistraterXll of xll: string
        | UnregistraterXll of xll: string

