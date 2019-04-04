open System.Runtime.Serialization
open Newtonsoft.Json
open System
open Akka.Util
open CellScript.Core.Tests

// Learn more about F# at http://fsharp.org
#nowarn "0064"

open CellScript.Core
open Akkling
open System.Threading

let inline (<!) (actorRef: IActorRef<_>) (msg) =
    let array2D = toArray2D msg
    actorRef <! array2D

[<EntryPoint>]
let main argv =
    let processMsg (ctx: Actor<_>) = function
        | OuterMsg.TestString text ->
            ctx.Sender() <! String text
            ignored ()
        | OuterMsg.TestString2Params (text1, text2) ->
            ctx.Sender() <! String text1
            ignored ()
        | OuterMsg.TestTable table ->
            ctx.Sender() <! table
            ignored ()
        | OuterMsg.TestInnerMsg innerMsg ->
            ctx.Sender() <! String "TestInnerMsg"
            ignored ()


    let remote = Remote(RemoteKind.Server (6001, processMsg))

    let manualSet = new ManualResetEventSlim()
    manualSet.Wait()
    0 // return an integer exit code
