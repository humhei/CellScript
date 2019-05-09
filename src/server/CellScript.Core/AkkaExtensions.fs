namespace CellScript.Core
open Akkling
open System
open Akka.Actor
open Akka.Cluster
open Akkling.Cluster
open System.Threading
open Akka.Remote
open Akka.Remote.Transport

module AkkaExtensions =

    [<RequireQualifiedAccess; Struct>]
    type RemoteResponse = 
        | Ok of ok: obj
        | Error of errorText: string


    type EndpointMsg<'ServerMsg> =
        | AddServer of IActorRef<'ServerMsg>
        | ClientMsg of obj

    [<RequireQualifiedAccess>]
    module CancelableAskAgent =
        type Msg<'AskMsg> =
            | Ask of remoteServer: IActorRef<'AskMsg> * Address * 'AskMsg * TimeSpan option

        type CancelableAsk =
            { ClusterMemberIdentity: Address
              Sender: IActorRef<Result<RemoteResponse, Address>> }

        let create serverRoleName (system: ActorSystem): IActorRef<Msg<'AskMsg>> =

            let actor = 
                spawnAnonymous system (props (fun ctx ->
                    let cluster = Cluster.Get(ctx.System)

                    let log = ctx.Log.Value
                    let rec loop (state: Map<Address, CancelableAsk>) = actor {
                        let! msg = ctx.Receive() : IO<obj>

                        match msg with
                        | LifecycleEvent e ->
                            match e with
                            | PreStart ->
                                cluster.Subscribe(untyped ctx.Self, ClusterEvent.InitialStateAsEvents,
                                    [| typedefof<ClusterEvent.IMemberEvent> |])
                                log.Info (sprintf "Actor subscribed to Cluster status updates: %A" ctx.Self)
                            | PostStop ->
                                log.Info (sprintf "Actor unsubscribed from Cluster status updates: %A" ctx.Self)

                            | _ -> return Unhandled
                            return! loop state

                        | IMemberEvent e ->
                            match e with
                            | MemberJoined m ->
                                return! loop state

                            | MemberUp m ->
                                return! loop state

                            | MemberLeft m ->
                                return! loop state

                            | MemberExited m ->
                                log.Info (sprintf "Node removed: %A" m)
                                return! loop state

                            | MemberRemoved m ->
                                log.Info (sprintf "Node removed: %A" m)
                                match Map.tryFind m.Address state, m.Roles.Contains serverRoleName with 
                                | Some cancelableAsk, true ->
                                    cancelableAsk.Sender <! Result.Error cancelableAsk.ClusterMemberIdentity
                                    return! loop (state.Remove cancelableAsk.ClusterMemberIdentity )
                                | _ ->
                                    return! loop state

                        | :? RemoteResponse as response ->
                            let remoteSender = ctx.Sender()
                            let memberIdentity = remoteSender.Path.Address

                            let cancelableAsk = state.[memberIdentity]
                            cancelableAsk.Sender <! Result.Ok response

                            return! loop (state.Remove memberIdentity)

                        | :? Msg<'AskMsg> as msg ->
                            match msg with 
                            | Msg.Ask (remoteServer, memberIdentity, msg, timeSpan) ->
                                remoteServer <! msg

                                let cancelableAsk =
                                    { ClusterMemberIdentity = memberIdentity
                                      Sender = ctx.Sender() }

                                return! loop (state.Add(memberIdentity, cancelableAsk))

                        | _ -> return Unhandled

                        return! loop state
                    }
                    loop Map.empty
                ))

            system.EventStream.Subscribe((actor :> IInternalTypedActorRef).Underlying, typeof<DisassociatedEvent>)
            |> ignore

            actor
            |> retype



    let retypeToDUChild (mapping: 'ChildMsg -> 'Msg) (actor: ICanTell<'Msg>) =
        { new ICanTell<'ChildMsg> with
            member this.Ask(arg1: 'ChildMsg, ?arg2: TimeSpan): Async<'Response> = 
                actor.Ask(mapping arg1, arg2)

            member this.Tell(arg1: 'ChildMsg, arg2: IActorRef): unit = 
                actor <! (mapping arg1)

            member this.Underlying: ICanTell = 
                actor.Underlying }

    [<RequireQualifiedAccess>]
    module ClusterClient =

        type private Model<'ServerMsg> =
            { Servers: Set<IActorRef<'ServerMsg>> 
              Jobs: int }

        [<RequireQualifiedAccess>]
        type private ProcessJob<'ServerMsg> = 
            | Tell of 'ServerMsg
            | Ask of 'ServerMsg * timespan: TimeSpan option


        [<RequireQualifiedAccess>]
        type private Msg<'ServerMsg, 'CallbackMsg> =
            | ProcessJob of ProcessJob<'ServerMsg>
            | Callback of 'CallbackMsg


        let create system name serverRoleName (callbackActor: IActorRef<'CallbackMsg> option) : ICanTell<'ServerMsg> = 
            let cancelableAskAgent = CancelableAskAgent.create serverRoleName system

            let actor : ICanTell<ProcessJob<'ServerMsg>> = 
                spawn system name (props (fun ctx ->
                    let log = ctx.Log.Value

                    let rec loop (model:Model<'ServerMsg>)  = actor {
                        let! recievedMsg = ctx.Receive() : IO<EndpointMsg<'ServerMsg>>
                        match recievedMsg with 
                        | EndpointMsg.AddServer server ->
                            return! loop { model with Servers = model.Servers.Add server }
                        | EndpointMsg.ClientMsg clientMsg ->
                            match unbox clientMsg with 
                            | Msg.Callback callback ->
                                match callbackActor with 
                                | Some callbackActor ->
                                    callbackActor <! callback
                                | None -> log.Error (sprintf "Cannot find a callback actor to process %A" callback)

                                return! loop model

                            | Msg.ProcessJob (processJob) ->
                                if model.Servers.Count = 0 then 
                                    match processJob with 
                                    | ProcessJob.Tell _  ->
                                        log.Warning(sprintf "Service unavailable, try again later. %A" recievedMsg)

                                    | ProcessJob.Ask _ ->
                                        log.Warning(sprintf "Service unavailable, try again later. %A" recievedMsg)
                                        ctx.Sender() <! Result.Error (sprintf "Service unavailable, try again later. %A" recievedMsg)

                                    return Unhandled
                                else 
                                    let remoteServer =  
                                        model.Servers
                                        |> Seq.item (model.Jobs % model.Servers.Count)  

                                    match processJob with 
                                    | ProcessJob.Tell msg -> 
                                        remoteServer <! msg
                                        return! loop { model with Jobs = model.Jobs + 1}

                                    | ProcessJob.Ask (msg, timeSpan) ->
                                        let! (result: Result<RemoteResponse, Address>) = 
                                            let memberIdentity = remoteServer.Path.Address
                                            cancelableAskAgent <? (CancelableAskAgent.Msg.Ask (remoteServer, memberIdentity , msg, timeSpan))
                                        
                                        match result with 
                                        | Result.Ok result ->
                                            let result =
                                                match result with 
                                                | RemoteResponse.Ok ok -> Result.Ok ok
                                                | RemoteResponse.Error errorText -> Result.Error errorText

                                            ctx.Sender() <! result

                                            return! loop { model with Jobs = model.Jobs + 1}

                                        | Result.Error memberIdentity ->
                                            let newModel = 
                                                { model with 
                                                    Servers = 
                                                        model.Servers
                                                        |> Set.filter(fun server -> 
                                                            server.Path.Address <> memberIdentity) } 

                                            let sender = ctx.Sender()
                                            async {
                                                let! (result: Result<obj, string>) = 
                                                    ctx.Self <? recievedMsg

                                                sender <! result
                                            } |> Async.Start

                                            return! loop newModel
                    }
                    loop { Servers = Set.empty; Jobs = 0 }
                ))
                |> retypeToDUChild (fun processJob ->
                    let msg : Msg<'ServerMsg, 'CallbackMsg> = Msg.ProcessJob processJob
                    msg
                    |> box
                    |> EndpointMsg.ClientMsg
                )


            { new ICanTell<'ServerMsg> with
                member this.Ask(arg1: 'ServerMsg, ?arg2: TimeSpan): Async<'Response> = 
                    actor <? (ProcessJob.Ask (arg1, arg2))
                member this.Tell(arg1: 'ServerMsg, arg2: IActorRef): unit = 
                    actor <! (ProcessJob.Tell arg1)
                member this.Underlying: ICanTell = 
                    actor.Underlying }




    [<RequireQualifiedAccess>]
    module ClusterServerStatusListener =

        type private Model<'Msg> = 
            { Nodes: Map<string * string, MemberStatus * ICanTell<'Msg>> }


        let private start server (system: ActorSystem) clientRoleName = 
            spawnAnonymous system (props (fun ctx ->
                let log = ctx.Log.Value
                let cluster = Cluster.Get(ctx.System)

                let rec loop (model: Model<EndpointMsg<'ServerMsg>>) = actor {
                    let! msg = ctx.Receive()  : IO<obj>
                    match msg with
                    | LifecycleEvent e ->
                        match e with
                        | PreStart ->
                            cluster.Subscribe(untyped ctx.Self, ClusterEvent.InitialStateAsEvents,
                                [| typedefof<ClusterEvent.IMemberEvent> |])
                            log.Info (sprintf "Actor subscribed to Cluster status updates: %A" ctx.Self)
                        | PostStop ->
                            log.Info (sprintf "Actor unsubscribed from Cluster status updates: %A" ctx.Self)

                        | _ -> return Unhandled
                        return! loop model

                    | IMemberEvent e ->
                        match e with
                        | MemberJoined m ->
                            log.Info (sprintf "Node joined: %A" m)
                            return! loop model
                        | MemberUp m ->
                            log.Info (sprintf "Node up: %A" m)

                            if m.HasRole(clientRoleName) then
                                let remoteAddr = sprintf "%O/user/%s" m.Address clientRoleName
                                let remoteClient = select system remoteAddr :> ICanTell<EndpointMsg<'ServerMsg>>
                                log.Info (sprintf "Remote Node up: %A" m)

                                let newModel = 
                                    { model with 
                                        Nodes = 
                                            model.Nodes 
                                            |> Map.add (m.Address.ToString(), m.UniqueAddress.ToString()) (m.Status, remoteClient)
                                    } 

                                remoteClient <! EndpointMsg.AddServer server

                                return! loop newModel

                            else return! loop model

                        | MemberLeft m ->
                            log.Info (sprintf "Node left: %A" m)
                            return! loop model
                        | MemberExited m ->
                            log.Info (sprintf "Node exited: %A" m)
                            if m.HasRole(clientRoleName) then
                                log.Info (sprintf "Remote Node exited: %A" m)
                                let newModel = 
                                    { model with 
                                        Nodes = 
                                            model.Nodes 
                                            |> Map.remove (m.Address.ToString(), m.UniqueAddress.ToString())
                                    }

                                return! loop newModel
                            else return! loop model

                        | MemberRemoved m ->
                            log.Info (sprintf "Node removed: %A" m)
                            return! loop model

                        return! loop model

                    | _ -> return Unhandled
                }

                loop { Nodes = Map.empty }
            ))
            |> ignore



        let listenServer server system clientRoleName =
            start server system clientRoleName


