module FcsWatch.AutoReload.ExcelDna.Registration

open ExcelDna.Integration
open Akkling
open FcsWatch.AutoReload.ExcelDna.Core
open Remote
open System.Reflection
open System.IO
open NLog

type Config = 
    { ServerPort: int }

let logger = LogManager.GetCurrentClassLogger()


type FsAsyncAddIn () =
    interface IExcelAddIn with
        member this.AutoOpen () =

            ExcelIntegration.RegisterUnhandledExceptionHandler(fun ex ->
                let exText = "!!! ERROR: " + ex.ToString()
        
                logger.Error (sprintf "%s" exText)

                exText
                |> box
            )

            async {
                let config = 
                    let dir = 
                        let codeBaseDir = Path.GetDirectoryName (Assembly.GetExecutingAssembly().CodeBase)
                        codeBaseDir.Replace("file:\\", "")
                        //codeBase.Replace()
                    let file = Path.Combine(dir,"appsettings.json")
                    if File.Exists file then
                        let text = File.ReadAllText(file)
                        Newtonsoft.Json.JsonConvert.DeserializeObject<Config>(text)
                    else 
                        { ServerPort = 8090 }


                Server.createActor config.ServerPort (fun ctx msg ->
                    let log = ctx.Log.Value
                    match msg with 
                    | Msg.RegistraterXll xll -> 
                        try 
                            ExcelAsyncUtil.QueueAsMacro(fun _ ->
                                ExcelIntegration.RegisterXLL xll |> ignore
                            )
                            log.Info(sprintf "register xll %s successfully" xll)
                            ctx.Sender() <! "SUCCESS"
                        with ex ->
                            log.Error(sprintf "%A" ex)
                            ctx.Sender() <! ex.Message

                        ignored ()

                    | Msg.UnregistraterXll xll -> 
                        try 
                            ExcelAsyncUtil.QueueAsMacro(fun _ ->
                                ExcelIntegration.UnregisterXLL xll |> ignore
                            )             
                            log.Info(sprintf "unRegister xll %s successfully" xll)
                            ctx.Sender() <! "SUCCESS"
                        with ex ->
                            log.Error(sprintf "%A" ex)

                            ctx.Sender() <! ex.Message
                        ignored ()
                ) 
            } |> Async.Start
        member this.AutoClose () =
            ()
