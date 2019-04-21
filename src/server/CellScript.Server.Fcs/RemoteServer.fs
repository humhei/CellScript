module internal CellScript.Server.Fcs.RemoteServer
open Akkling
open CellScript.Core
open CellScript.Core.Remote



type Config = 
    { ServerPort: int }

let run (logger: NLog.FSharp.Logger) =
    let processMsg (ctx: Actor<_>) msg =
        logger.Info "fcs server receive %A" msg
        match msg with 
        | FcsMsg.Eval (xlRef, code) ->
            let result = Fcs.eval xlRef code 
            ctx.Sender() <! result
            ignored()

    Server.createActor processMsg
