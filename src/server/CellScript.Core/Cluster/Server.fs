namespace Akkling.Cluster.ServerSideDevelopment
open Akkling
open Akka.Actor
open Akka.Cluster
open Akkling.Cluster
open System

[<RequireQualifiedAccess>]
module Server =

    type private Model<'Msg, 'ServerModel> = 
        { Nodes: Map<Address, ClusterActor<'Msg>>
          CustomModel: 'ServerModel }

    let createAgent
        (callbackGeneric: 'CallbackMsg -> unit)
        (system: ActorSystem) 
        name 
        clientRoleName 
        (initialServerModel: 'ServerModel) 
        (handleMsg: Actor<obj> -> 'ServerMsg -> 'ServerModel -> 'ServerModel) = 
            let notify clients = 
                system.EventStream.Publish(UpdateCallbackClientsEvent clients)

            spawn system name (props (fun ctx ->
                let log = ctx.Log.Value
                let cluster = Cluster.Get(ctx.System)

                let self: ClusterActor<'ServerMsg> = ClusterActor.retype (ClusterActor.ofIActorRef ctx.Self)

                let actorSelectOfAddress addr =
                    let remoteAddr = sprintf "%O/user/%s" addr clientRoleName
                    let remoteClient = TypedActorSelectionPersistent.create system remoteAddr
                    { Actor = Choice2Of2 remoteClient
                      Address = addr }

                let rec loop (model: Model<'CallbackMsg, 'ServerModel>) = actor {

                    let! msg = ctx.Receive() : IO<obj>
                    let sender = ctx.Sender()

                    match msg with
                    | :? EndpointMsg<'CallbackMsg, 'ServerMsg> as msg ->
                        match msg with 
                        | EndpointMsg.AddClient (client: ClusterActor<'CallbackMsg>) ->
                            let sender = ctx.Sender()
                            let client = { client with Address = sender.Path.Address }
                            log.Info (sprintf "[SERVER] Add remote client: %A" client.Address)
                            let newModel = 
                                { model with 
                                    Nodes = 
                                        model.Nodes.Add (client.Address, (client))
                                } 
                            notify newModel.Nodes
                            return! loop newModel

                        | _ ->
                            log.Error (sprintf "[SERVER] Unexcepted msg %A" msg)
                            return Unhandled

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
                        | MemberJoined m | MemberUp m when isRemoteJoined clientRoleName e ->
                            let remoteClient = actorSelectOfAddress m.Address

                            log.Info (sprintf "[SERVER] Remote Node joined: %A" m)

                            let newModel = 
                                { model with 
                                    Nodes = model.Nodes.Add (remoteClient.Address, remoteClient) } 

                            notify newModel.Nodes

                            let addServer: EndpointMsg<'CallbackMsg, 'ServerMsg> = EndpointMsg.AddServer self

                            ClusterActor.retype remoteClient <! addServer

                            return! loop newModel

                        | MemberJoined m ->
                            log.Info (sprintf "[SERVER] Node joined: %A" m)
                            return! loop model

                        | MemberUp m ->
                            log.Info (sprintf "[SERVER] Node up: %A" m)
                            return! loop model

                        | MemberLeft m ->
                            log.Info (sprintf "[SERVER] Node left: %A" m)
                            return! loop model

                        | MemberExited m ->
                            log.Info (sprintf "[SERVER] Node exited: %A" m)
                            return! loop model

                        | MemberRemoved m ->
                            log.Info (sprintf "[SERVER] Node removed: %A" m)
                            if m.HasRole(clientRoleName) then
                                log.Info (sprintf "[SERVER] Remote Node removed: %A" m)
                                let newModel = 
                                    { model with 
                                        Nodes = 
                                            model.Nodes
                                            |> Map.filter (fun addr _ -> addr <> m.Address)
                                    }

                                notify newModel.Nodes

                                return! loop newModel
                            else return! loop model

                        return! loop model

                    | :? 'ServerMsg as serverMsg ->
                        log.Info (sprintf "[SERVER] recieve msg %A" serverMsg)
                        let nodes = 

                            match Map.tryFind sender.Path.Address model.Nodes, sender.Path.Name = clientRoleName with 
                            | Some _, true ->
                                model.Nodes
                            | None, true ->
                                log.Info (sprintf "[SERVER] Remote Node joined %A" sender.Path.Address)
                                let remoteClient = actorSelectOfAddress sender.Path.Address
                                model.Nodes.Add(sender.Path.Address, remoteClient)
                            | _ -> 
                                log.Error (sprintf "[SERVER] Unknown remtoe client %A" sender.Path.Address )
                                model.Nodes

                        let newModel = 
                            { model with 
                                CustomModel = handleMsg ctx serverMsg model.CustomModel
                                Nodes = nodes }
                                 
                        return! loop newModel

                    | _ -> 
                        log.Error (sprintf "[SERVER] Unexpected msg %A" msg)
                        return Unhandled
                }

                loop { Nodes = Map.empty; CustomModel = initialServerModel }
            ))
            |> ignore



