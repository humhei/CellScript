namespace FcsWatch.AutoReload.Core

open Akkling

module Remote = 
    type Client<'Msg> =
        { ClientPort: int 
          ServerPort: int
          RemoteServer: TypedActorSelection<'Msg>}

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
                            }
                        }
                    """ clientPort

            let serverAddr = sprintf "akka.tcp://server@localhost:%d" serverPort

            let remoteProps = 
                let actor = actorOf ignored
                { props actor with Deploy = None }


            let client = System.create "client" <| Configuration.parse config

            let clientAgent = spawn client "CellScript.Core.Client" remoteProps

            let serverPath = sprintf "%s/user/CellScript.Core.Server" serverAddr

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
                            }
                        """ port


                let system = System.create "server" <| Configuration.parse serverConfig
                let server = spawn system "CellScript.Core.Server" (props (actorOf2 handleMsg))
                ()


    [<RequireQualifiedAccess>]
    type Msg =
        | RegistraterXll of xll: string
        | UnregistraterXll of xll: string

