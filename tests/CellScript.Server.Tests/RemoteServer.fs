module internal RemoteServer
open Akkling
open CellScript.Core.Tests
open CellScript.Core
open CellScript.Core.Cluster
open CellScript.Core.Conversion
open CellScript.Core.AkkaExtensions

let inline (<!) (actorRef: IActorRef<_>) (msg) = 
    let array2D = toArray2D msg
    actorRef <! RemoteResponse.Ok array2D

type Config = 
    { ServerPort: int }

let run (logger: NLog.FSharp.Logger) =
    let handleMsg (ctx: Actor<_>) msg customModel =
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


    Server.createAgent 9050 () handleMsg
