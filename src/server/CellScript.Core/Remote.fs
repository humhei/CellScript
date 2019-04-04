namespace CellScript.Core

open System.Collections.Concurrent

#nowarn "0064"
open Akkling
open Akka.Actor




[<RequireQualifiedAccess>]
type RemoteKind<'Msg> =
    | Client of Client<'Msg>
    | Server of serverPort: int * processMsg: (Actor<'Msg> -> 'Msg -> Effect<'Msg>)

[<RequireQualifiedAccess>]
module private Cache =
    let serverPortProcessMsgMapping = ConcurrentDictionary<int (*server port*), obj>()


type Remote<'Msg>(remoteKind: RemoteKind<'Msg>) =

    let mutable processMsgMutable: option<(Actor<'Msg> -> 'Msg -> Effect<'Msg>)> = None

    do
        match remoteKind with 
        | RemoteKind.Client _ -> ()
        | RemoteKind.Server (port, processMsg) ->
            match Cache.serverPortProcessMsgMapping.TryAdd(port, (box processMsg)) with 
            | true -> ()
            | false -> failwithf "add server %d processMsg failed" port


    let remoteActor = 
        match remoteKind with 
        | RemoteKind.Client client ->
            let processMsg ctx msg = 
                match processMsgMutable with 
                | Some processMsg -> processMsg ctx msg
                | None ->
                    match Cache.serverPortProcessMsgMapping.TryGetValue client.ServerPort with 
                    | true, v -> 
                        printfn "connect process msg from server %d" client.ServerPort
                        processMsgMutable <- Some (unbox v)
                        processMsgMutable.Value ctx msg
                    | false, _ -> 
                        printfn "cannot process msg from server %d" client.ServerPort
                        ignored()



            let toRemoteProps addr actor = { props actor with Deploy = Some (Deploy(RemoteScope(Address.Parse addr))) }
            let remoteProps = toRemoteProps (sprintf "akka.tcp://server@localhost:%d" client.ServerPort) (actorOf2 processMsg)

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
                    """ client.ClientPort

            let client = System.create "client" <| Configuration.parse config
            let clientAgent = spawn client "remote-actor" remoteProps
            Some clientAgent

        | RemoteKind.Server (serverPort,processMsg) ->
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
                    """ serverPort


            let serverAgent = System.create "server" <| Configuration.parse serverConfig
            None

    member x.RemoteActor = remoteActor


[<AutoOpen>]
module Remote =
    type BaseType =
        | Int of int
        | String of string
        | Float of float
    with 
        member x.ToArray2D() =
            match x with 
            | BaseType.Int x ->
                array2D [[box x]]
            | BaseType.String x ->
                array2D [[box x]]
            | BaseType.Float x ->
                array2D [[box x]]


    let inline toArray2D (x : ^a) =
        (^a : (member ToArray2D : unit -> obj [,]) (x))

