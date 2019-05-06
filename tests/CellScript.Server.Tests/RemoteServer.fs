module internal RemoteServer
open Akkling
open CellScript.Core.Tests
open CellScript.Core
open CellScript.Core.Remote
open CellScript.Core.Conversion

let inline (<!) (actorRef: IActorRef<_>) (msg) =
    let array2D = toArray2D msg
    actorRef <! array2D

type Config = 
    { ServerPort: int }

let run (logger: NLog.FSharp.Logger) =
    let handleMsg (ctx: Actor<_>) msg clients customModel =
        logger.Info "server receive %A" msg
        match msg with   
        | OuterMsg.TestString text ->
            ctx.Sender() <! String text

        | OuterMsg.TestString2Params (text1, text2) ->
            ctx.Sender() <! String text1

        | OuterMsg.TestTable table ->
            ctx.Sender() <! table

        | OuterMsg.TestInnerMsg innerMsg ->
            match innerMsg with 

            | InnerMsg.TestExcelReference xlRef -> 
                ctx.Sender() <! Record xlRef

            | InnerMsg.TestString text ->
                ctx.Sender() <! String text

        | OuterMsg.TestExcelVector vector ->
            ctx.Sender() <! vector


    let config = Server.fetchConfig 9050

    let system = Server.createSystem config

    Server.createRemoteActor config system () ignore handleMsg
