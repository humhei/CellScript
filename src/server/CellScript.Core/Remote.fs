namespace CellScript.Core
open Akkling
open Types
open System
open Microsoft.FSharp.Quotations.Patterns
open Akka.Actor
open System.Reflection
open System.IO
open Akka.Remote.Transport
open Akka.Remote

module Remote = 

    [<RequireQualifiedAccess>]
    type ClientMsg<'CustomMsg> =
        | UpdateCellValues of xlRefs: SerializableExcelReference list
        | Exec of toolName: string * args: string list * workingDir: string
        | CustomMsg of 'CustomMsg

    [<RequireQualifiedAccess>]
    type ServerMsg<'ServerCustomMsg,'ClientCustomMsg> =
        | RemoveClient of DisassociatedEvent
        | AddClient of IActorRef<ClientMsg<'ClientCustomMsg>>
        | CustomMsg of 'ServerCustomMsg

    type Client<'ServerCustomMsg> =
        { ClientPort: int 
          ServerPort: int
          RemoteServer: ICanTell<'ServerCustomMsg> }


    let private withAppconfigOrWebconfig baseConfig =
        Configuration.load().WithFallback baseConfig


    [<RequireQualifiedAccess>]
    module Client =

        let create remotePort (handleClientMsg: ClientMsg<'ClientCustomMsg> -> Effect<ClientMsg<'ClientCustomMsg>>): Client<'ServerCustomMsg> =
            let remoteActorName = typeof<'ServerCustomMsg>.FullName
            //let remoteActorName = "CellScript.Core.Server"
            let config = 
                sprintf 
                    """
                        akka {
                            remote.port = %d
                            remote.actor.name = %s
                            actor.provider = remote
                            remote.dot-netty.tcp {
                                hostname = localhost
                                port = 0
                            }
                            loglevel = DEBUG
                            loggers=["Akka.Logger.NLog.NLogLogger, Akka.Logger.NLog"]
                        }
                    """ remotePort remoteActorName
                |> Configuration.parse
                |> withAppconfigOrWebconfig

            let system = System.create "cellScriptClient" <| config

            let clientPort =
                config.GetInt("akka.remote.dot-netty.tcp.port")

            let client: IActorRef<ClientMsg<'ClientCustomMsg>> = 
                let actorName = config.GetString("akka.actor.name","cellScriptClient")
                spawn system actorName (props (actorOf handleClientMsg))

            let remotePort = config.GetInt "akka.remote.port"

            let remotePath = 
                let remoteAddr = sprintf "akka.tcp://cellScriptServer@localhost:%d" remotePort
                let remoteActorName = config.GetString "akka.remote.actor.name"
                sprintf "%s/user/%s" remoteAddr remoteActorName

            let remoteServer: TypedActorSelection<ServerMsg<'ServerCustomMsg,'ClientCustomMsg>> = select system remotePath
            remoteServer <! ServerMsg.AddClient client


            { ClientPort = clientPort
              ServerPort = remotePort
              RemoteServer = 
                let remoteServerICanTell = remoteServer :> ICanTell<ServerMsg<'ServerCustomMsg,'ClientCustomMsg>> 

                { new ICanTell<'ServerCustomMsg> with 
                    member x.Ask (msg, ?timeSpan) = 
                        remoteServerICanTell.Ask(ServerMsg.CustomMsg msg, timeSpan)

                    member x.Tell(msg, actorRef) =
                        remoteServerICanTell.Tell(ServerMsg.CustomMsg msg, actorRef)

                    member x.Underlying = remoteServer :> Akka.Actor.ICanTell }
                }


    type ServerModel<'CustomModel, 'ClientCustomMsg> =
        { Clients: IActorRef<ClientMsg<'ClientCustomMsg>> list
          CustomModel: 'CustomModel }

    [<RequireQualifiedAccess>]
    module Server =

        let createSystem config = System.create "cellScriptServer" config

        let fetchConfig port =
            let config = 
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
                |> Configuration.parse
                |> withAppconfigOrWebconfig
            
            config

        type HandleMsg<'ServerCustomMsg,'ClientCustomMsg, 'ServerCustomModel> = 
            Actor<ServerMsg<'ServerCustomMsg,'ClientCustomMsg>> 
                -> 'ServerCustomMsg 
                -> IActorRef<ClientMsg<'ClientCustomMsg>> list
                -> 'ServerCustomModel
                -> 'ServerCustomModel

        let createRemoteActor (config: Akka.Configuration.Config) (system: ActorSystem) initialCustomModel actionAfterUpdatedClient (handleMsg: HandleMsg<'ServerCustomMsg,'ClientCustomMsg, 'ServerCustomModel>) =
            
            let systemPathName = config.GetString ("akka.actor.name", typeof<'ServerCustomMsg>.FullName)
            let server = spawn system systemPathName (props (fun ctx ->
                let rec loop model = actor {
                    let! msg = ctx.Receive()
                    match msg with 
                    | ServerMsg.AddClient client ->
                        let newClients = client :: model.Clients
                        let newState =
                            { model with Clients = newClients }
                        actionAfterUpdatedClient newClients
                        return! loop newState

                    | ServerMsg.RemoveClient disassociatedEvent ->
                        let newState =
                            { model with Clients = model.Clients |> List.filter (fun client -> client.Path.Address <> disassociatedEvent.RemoteAddress)}
                        return! loop newState

                    | ServerMsg.CustomMsg customMsg ->
                        let newCustomModel = handleMsg ctx customMsg model.Clients model.CustomModel
                        return! loop { model with CustomModel = newCustomModel }
                }

                loop { Clients = []; CustomModel = initialCustomModel }
            
            ))

            let disassociated = spawn system "CellScriptServerDisassociated" (props (fun ctx ->
                let rec loop () = actor {
                    let! msg = ctx.Receive() : IO<obj>
                    match msg with 
                    | :? DisassociatedEvent as disassociated ->
                        server <! ServerMsg.RemoveClient disassociated
                        return! loop ()

                    | _ -> return Unhandled
                }
                loop()
            ))

            let b = system.EventStream.Subscribe(untyped disassociated, typeof<DisassociatedEvent>)
            assert (b = true)
            
            ()





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

    let commandMapping =
        lazy 
            let rec asCommandUci expr =
                match expr with 
                | Lambda(_, expr) ->
                    asCommandUci expr

                | NewUnionCase (uci, exprs) ->
                    if exprs.Length <> 1 then failwithf "command %A msg's paramters length should be equal to 1" uci
                    else uci
                | _ -> failwithf "command expr %A should be a union case type" expr

            [ asCommandUci <@ FcsMsg.EditCode @>, { Shortcut = Some "+^{F11}" }]
            |> dict