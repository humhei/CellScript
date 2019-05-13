namespace Akkling.Cluster.ServerSideDevelopment
open Akkling
open System
open Akka.Actor
open Akka.Cluster
open Akkling.Cluster
open System.Threading
open Akka.Remote
open Akka.Remote.Transport

[<AutoOpen>]
module private Utils =
    let isRemoteJoined remoteRoleName (e: ClusterEvent.IMemberEvent) =
        match e with 
        | MemberJoined m ->
            m.HasRole remoteRoleName && m.Status = MemberStatus.Up
        | MemberUp m ->
            m.HasRole remoteRoleName
        | _ -> false



//[<RequireQualifiedAccess>]
//module TypedActorSelectionPersistent =
//    let create system path =
//        { Path = path 
//          System = system }
        

[<CustomComparison; CustomEquality>]
type ClusterActor<'Msg> =
    { Actor: TypedActorSelectionPersistent<'Msg>
      Address: Address
      System: ActorSystem
      Role: string }
with 
    override x.Equals(yobj) =
        match yobj with
        | :? ClusterActor<'Msg> as y -> (x.Address = y.Address)
        | _ -> false

    override x.GetHashCode() = hash x.Address

    member x.RemotePath = sprintf "%O/user/%s" x.Address x.Role

    member x.AsICanTell =
        match x.Actor with 
        | Choice1Of2 actor -> actor :> ICanTell<_>
        | Choice2Of2 actor -> actor.Value :> ICanTell<_>

    interface ICanTell<'Msg> with 
        member x.Ask(msg, ?timeSpan) = x.AsICanTell.Ask(msg, timeSpan)
        
        member x.Tell (msg, actorRef) = x.AsICanTell.Tell(msg, actorRef)

        member x.Underlying = x.AsICanTell.Underlying

    interface System.IComparable with

          member x.CompareTo yobj =
              match yobj with

              | :? ClusterActor<'Msg> as y -> compare x.Address y.Address

              | _ -> invalidArg "yobj" "cannot compare values of different types"

[<RequireQualifiedAccess>]
module ClusterActor =
    let ofIActorRef (actor: IActorRef<_>) = 
        { Actor = Choice1Of2 actor
          Address = actor.Path.Address }


    let retype (clusterActor: ClusterActor<_>) =

        { Address =  clusterActor.Address
          Actor = 
            match clusterActor.Actor with
            | Choice1Of2 actor -> 
                retype actor
                |> Choice1Of2
            | Choice2Of2 actor ->
                TypedActorSelectionPersistent.create actor.System actor.Path
                |> Choice2Of2
        }

type UpdateCallbackClientsEvent<'ClientCallbackMsg> = 
    UpdateCallbackClientsEvent of Map<Address, ClusterActor<'ClientCallbackMsg>>


[<RequireQualifiedAccess>]
type EndpointMsg<'CallbackMsg,'ServerMsg> =
    | AddServer of ClusterActor<'ServerMsg>
    | AddServerFromEvent of ClusterActor<'ServerMsg>
    | RemoveServer of Address
    | AddClient of ClusterActor<'CallbackMsg>
    | GetEndpoints