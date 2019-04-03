namespace CellScript.Core

open Akkling
open Akka.Actor


[<AutoOpen>]
module Msg =

    [<RequireQualifiedAccess>]
    type InnerMsg =
        | TestString of string


    [<RequireQualifiedAccess>]
    type OuterMsg = // That's the one the update expects
        | TestTable of Table
        | TestInnerMsg of InnerMsg
        | TestString of string
        | TestString2Params of text1: string * text2: string
        | TestArray of string []


[<RequireQualifiedAccess>]
type Remote<'msg> =
    | Client 
    | Server of (Actor<'msg> -> 'msg -> Effect<'msg>)

[<RequireQualifiedAccess>]
module Routed =
    let private clientPort = 3523
    let private serverPort = 9001

    let mutable remote: Remote<OuterMsg> = Remote.Client

    let private processMsg ctx msg = 
        match remote with 
        | Remote.Client -> ignored ()
        | Remote.Server processMsg ->
            processMsg ctx msg

    let private toRemoteProps addr actor = { props actor with Deploy = Some (Deploy(RemoteScope(Address.Parse addr))) }


    let clientConfig = 
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


    let remoteProps() = toRemoteProps (sprintf "akka.tcp://server@localhost:%d" serverPort) (actorOf2 processMsg)