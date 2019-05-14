namespace Akkling.Cluster.ServerSideDevelopment
open Akkling
open System
open Akka.Actor
open Akka.Cluster
open Akkling.Cluster
open System.Threading
open System.Diagnostics
open System.Timers
open System.Collections.Generic

module Client =

    type SeedNodes = SeedNodes of IList<string>

    [<RequireQualifiedAccess>]
    module SeedNodes =
        let containAddress (address: Address) (SeedNodes seedNodes) =
            seedNodes
            |> Seq.exists (fun node -> String.Compare(node.TrimEnd('/'), address.ToString(),true) = 0)


    [<RequireQualifiedAccess>]
    type ProcessJob<'ServerMsg> = 
        | Tell of 'ServerMsg
        | Ask of 'ServerMsg * timespan: TimeSpan option

    [<RequireQualifiedAccess>]
    module MemberManagement =

        type private Model =
            { Servers: Map<Address, ClusterActor<obj>> }

        let createAgent (seedNodes: SeedNodes) system serverRoleName: IActorRef<EndpointMsg> = 
            spawnAnonymous system (props (fun ctx ->
                let log = ctx.Log.Value
                let rec loop (model: Model)  = actor {
                    let! recievedMsg = ctx.Receive()
                    let sender = ctx.Sender()

                    match recievedMsg with 
                    | EndpointMsg.AddServerFromEvent server ->
                        return! loop { model with Servers = model.Servers.Add (server.Address, ClusterActor.ofPersistent system server) }

                    | EndpointMsg.AddServer server ->
                        let server = { server with Address = sender.Path.Address }
                        if sender.Path.Name = serverRoleName then 
                            log.Info (sprintf "[CLIENT] Added remote server %A" server.Address)
                            return! loop { model with Servers = model.Servers.Add (server.Address, ClusterActor.ofPersistent system server) }
                        else 
                            log.Error (sprintf "[CLIENT] Server path %A doesn't contain role %s" server.Address serverRoleName)
                            return Unhandled

                    | EndpointMsg.RemoveServer addr ->
                        if SeedNodes.containAddress addr seedNodes then 
                            log.Info(sprintf "Didn't remove seed node %A" addr)
                            return! loop model
                        else
                            return! loop { model with Servers = model.Servers.Remove addr }

                    | EndpointMsg.AddClient _ ->

                        log.Error (sprintf "[CLIENT] Cannot accept AddClient msg in client side")
                        return Unhandled

                    | EndpointMsg.GetEndpoints _ ->
                        let sender = ctx.Sender()
                        sender <! model.Servers
                        return! loop model
                }
                loop { Servers = Map.empty }
            ))

    [<RequireQualifiedAccess>]
    type CancelableAskResponse =
        | Timeout of interval: float
        | MemberRemoved of Address
        | Success of obj

    type CancelableAsk =
        { ClusterMemberIdentity: Address
          Sender: IActorRef<CancelableAskResponse>
          Timer: Timer option }

    [<RequireQualifiedAccess>]
    type CancelableAskMsg<'ServerMsg> =
        | Ask of remoteServer: ClusterActor<'ServerMsg> * Address * 'ServerMsg * TimeSpan option



    [<RequireQualifiedAccess>]
    module CancelableAsk =

        type private Timeout = Timeout of Address * float

        let createAgent 
            callbackActor 
            name 
            serverRoleName 
            (memberManagementAgent: IActorRef<EndpointMsg>)
            (system: ActorSystem): IActorRef<CancelableAskMsg<'ServerMsg>> =

            let actor = 
                spawn system name (props (fun ctx ->
                    let cluster = Cluster.Get(ctx.System)

                    let log = ctx.Log.Value
                    let rec loop (state: list<CancelableAsk>) = actor {
                        let! msg = ctx.Receive() : IO<obj>
                        let sender = ctx.Sender()
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
                                return! loop state

                            | MemberRemoved m ->
                                log.Info (sprintf "[CLIENT] Remote Node removed: %A" m)
                                match state |> List.tryFind (fun ask -> ask.ClusterMemberIdentity = m.Address) , m.HasRole serverRoleName with 
                                | Some cancelableAsk, true ->

                                    match cancelableAsk.Timer with 
                                    | Some timer -> 
                                        timer.Stop()
                                        timer.Dispose()
                                    | None -> ()

                                    cancelableAsk.Sender <! CancelableAskResponse.MemberRemoved cancelableAsk.ClusterMemberIdentity
                                    return! loop (state |> List.filter (fun ask -> ask.ClusterMemberIdentity <> cancelableAsk.ClusterMemberIdentity))
                                | _ ->
                                    return! loop state

                        | :? 'CallbackMsg as callback ->
                            match callbackActor with 
                            | Some callbackActor ->
                                callbackActor <! callback
                            | None -> log.Error (sprintf "Cannot find a callback actor to process %A" callback)

                            return! loop state

                        | :? EndpointMsg as endpointMsg ->
                            memberManagementAgent <<! endpointMsg
                            return! loop state

                        | :? Timeout as timeout ->
                            log.Info (sprintf "[CLIENT] Ask task timeout: %A" timeout)
                            let (Timeout (addr, interval)) = timeout

                            match state |> List.tryFindIndex (fun ask -> ask.ClusterMemberIdentity = addr) with 
                            | Some index ->
                                let cancelableAsk = state.[index]
                                cancelableAsk.Sender <! CancelableAskResponse.Timeout interval
                                return! loop (state.[0..index - 1] @ state.[index + 1..state.Length - 1])
                            | _ ->
                                log.Error "Timeout, but the ask task has been already removed"
                                return! loop state

                        | :? CancelableAskMsg<'ServerMsg> as msg ->
                            match msg with 
                            | CancelableAskMsg.Ask (remoteServer, memberIdentity, msg, timeSpan) ->
                                let timer =
                                    match timeSpan with
                                    | Some timeSpan ->
                                        let timer = new Timer(timeSpan.TotalMilliseconds)
                                        timer.Start()
                                        timer.Elapsed.Add(fun _ ->
                                            timer.Stop()
                                            ctx.Self <! box (Timeout (memberIdentity, timeSpan.TotalMilliseconds))
                                        )
                                        Some timer
                                    | None -> None
                                
                                remoteServer <! msg

                                let cancelableAsk =
                                    { ClusterMemberIdentity = memberIdentity
                                      Sender = sender
                                      Timer = timer }

                                return! loop (cancelableAsk :: state)

                        | value -> 
                            match state |> List.tryFindIndex (fun ask -> ask.ClusterMemberIdentity = sender.Path.Address), sender.Path.Name = serverRoleName with 
                            | Some index, true ->
                                let cancelableAsk = state.[index]
                                match cancelableAsk.Timer with 
                                | Some timer -> 
                                    timer.Stop()
                                    timer.Dispose()
                                | None -> ()

                                log.Info (sprintf "[CLIENT] Recieve msg %A from remote server %A" value sender.Path.Address)
                                cancelableAsk.Sender <! CancelableAskResponse.Success value
                                return! loop (state.[0..index - 1] @ state.[index + 1..state.Length - 1])

                            | _ ->
                                log.Error (sprintf "[CLIENT] Unexcepted message %A" msg)
                                return Unhandled

                    }
                    loop []
                ))
                

            retype actor
            

    [<RequireQualifiedAccess>]
    module ClientAgent =

        type private Model =
            { Jobs: int }

        let create seedNodes system name serverRoleName (callbackActor: IActorRef<'CallbackMsg> option) : ICanTell<'ServerMsg> = 
            let memberManagementAgent = MemberManagement.createAgent seedNodes system serverRoleName

            let (cancelableAskAgent: IActorRef<CancelableAskMsg<'ServerMsg>>) = 
                CancelableAsk.createAgent callbackActor name serverRoleName memberManagementAgent system

            let actor : IActorRef<ProcessJob<'ServerMsg>> = 
                spawnAnonymous system (props (fun ctx ->
                    let log = ctx.Log.Value
                    let cluster = Cluster.Get(system)
                    let rec loop (model: Model)  = actor {
                        let! recievedMsg = ctx.Receive() : IO<obj>
                        match recievedMsg with 
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
                            | MemberJoined m | MemberUp m when isRemoteJoined serverRoleName e ->
                                log.Info (sprintf "[CLIENT] Remote Node joined: %A" m)
                                let server = 
                                    { Address = m.Address 
                                      Role = serverRoleName }

                                memberManagementAgent <! EndpointMsg.AddServerFromEvent server

                                let endpointMsg: EndpointMsg = 
                                    let client = 
                                        { Address = cancelableAskAgent.Path.Address 
                                          Role = name }

                                    EndpointMsg.AddClient client

                                (ClusterActor.ofPersistent system server :> ICanTell<_>).Underlying.Tell(endpointMsg, (cancelableAskAgent :> IInternalTypedActorRef).Underlying)

                                return! loop model

                            | MemberJoined m ->
                                log.Info (sprintf "[CLIENT] Node joined: %A" m)
                                return! loop model

                            | MemberUp m ->
                                log.Info (sprintf "[CLIENT] Node up: %A" m)
                                return! loop model

                            | MemberLeft m ->
                                return! loop model

                            | MemberExited m ->
                                return! loop model

                            | MemberRemoved m ->
                                return! loop model

                        | :? ProcessJob<'ServerMsg> as processJob ->
                            let! (servers: Map<Address, ClusterActor<obj>>) = 
                                memberManagementAgent <? EndpointMsg.GetEndpoints

                            let servers = servers |> Map.map(fun k v -> ClusterActor.retype v)

                            if servers.Count = 0 then 
                                match processJob with 
                                | ProcessJob.Tell _  ->
                                    log.Warning(sprintf "Service unavailable, try again later. %A" recievedMsg)

                                | ProcessJob.Ask _ ->
                                    log.Warning(sprintf "Service unavailable, try again later. %A" recievedMsg)
                                    let error = Result.Error (sprintf "Service unavailable, try again later. %A" recievedMsg)
                                    ctx.Sender() <! error

                                return Unhandled
                            else 
                                let remoteServer =  
                                    let pair =
                                        servers
                                        |> Seq.item (model.Jobs % servers.Count)  

                                    pair.Value

                                match processJob with 
                                | ProcessJob.Tell msg -> 
                                    remoteServer <! msg
                                    return! loop { model with Jobs = model.Jobs + 1}

                                | ProcessJob.Ask (msg, timeSpan) ->
                                    let result = 
                                        let addr = remoteServer.Address
                                        let task = cancelableAskAgent.Underlying.Ask(CancelableAskMsg.Ask (remoteServer, addr , msg, timeSpan))
                                        task.Result

                                    let (result: CancelableAskResponse) = unbox result
                                        

                                    match result with 
                                    | CancelableAskResponse.Success result ->
                                        let result: Result<obj, string> = Result.Ok result
                                        ctx.Sender() <! result
                                        return! loop { model with Jobs = model.Jobs + 1}

                                    | CancelableAskResponse.MemberRemoved memberIdentity ->
                                        memberManagementAgent <! EndpointMsg.RemoveServer memberIdentity
                                        ctx.Self <<! recievedMsg
                                        return! loop { model with Jobs = model.Jobs + 1}
                                    | CancelableAskResponse.Timeout timeout ->
                                        let result: Result<obj, string> = Result.Error (sprintf "Time exceed %f" timeout)
                                        ctx.Sender() <! result
                                        return! loop { model with Jobs = model.Jobs + 1}


                        | _ -> 
                            log.Error (sprintf "[CLIENT] unexcepted msg %A" recievedMsg)
                            return Unhandled
                    }
                    loop { Jobs = 0 }
                ))
                |> retype


            { new ICanTell<'ServerMsg> with
                member this.Ask(arg1: 'ServerMsg, ?arg2: TimeSpan): Async<'Response> = 
                    actor <? (ProcessJob.Ask (arg1, arg2))
                member this.Tell(arg1: 'ServerMsg, arg2: IActorRef): unit = 
                    actor <! (ProcessJob.Tell arg1)
                member this.Underlying: ICanTell = 
                    actor.Underlying }
