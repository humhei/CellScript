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
        

type ClusterActorPersistent =
    { Address: Address 
      Role: string }



[<CustomComparison; CustomEquality>]
type ClusterActor<'Msg> =
    { Address: Address
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
        select x.System x.RemotePath :> ICanTell<'Msg>

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
    let ofIActorRef system role (actor: IActorRef<_>) = 
        { System = system
          Role = role
          Address = actor.Path.Address }

    let ofPersistent system (clusterActorPersistent: ClusterActorPersistent) =
        { System = system
          Role = clusterActorPersistent.Role
          Address = clusterActorPersistent.Address }

    let retype (clusterActor: ClusterActor<_>) =

        { Address =  clusterActor.Address
          System = clusterActor.System
          Role = clusterActor.Role  }

type UpdateCallbackClientsEvent<'ClientCallbackMsg> = 
    UpdateCallbackClientsEvent of Map<Address, ClusterActor<'ClientCallbackMsg>>


[<RequireQualifiedAccess>]
type EndpointMsg =
    | AddServer of ClusterActorPersistent
    | AddServerFromEvent of ClusterActorPersistent
    | RemoveServer of Address
    | AddClient of ClusterActorPersistent
    | GetEndpoints