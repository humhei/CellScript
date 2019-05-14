module internal RemoteServer
open Akkling
open CellScript.Core.Tests
open CellScript.Core
open CellScript.Core.Cluster


type Config = 
    { ServerPort: int }

let run (logger: NLog.FSharp.Logger) =
    let handleMsg (ctx: Actor<_>) msg customModel =

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

        | OuterMsg.TestCommand commandCaller ->
            ctx.Sender() <! CommandResponse.Ok

    Server.createAgent 9050 () handleMsg
    |> ignore
