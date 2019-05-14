module FcsWatch.AutoReload.Server.Registration

open ExcelDna.Integration
open Akkling
open FcsWatch.AutoReload.Core
open Remote
open System.Reflection
open System.IO

type Config = 
    { ServerPort: int }

type FsAsyncAddIn () =
    interface IExcelAddIn with
        member this.AutoOpen () =
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
                        { ServerPort = 7010 }


                Server.createActor config.ServerPort (fun ctx msg ->
                    match msg with 
                    | Msg.RegistraterXll xll -> 
                        try 
                            ExcelAsyncUtil.QueueAsMacro(fun _ ->
                                ExcelIntegration.RegisterXLL xll |> ignore
                            )
                            ctx.Sender() <! "SUCCESS"
                        with ex ->
                            ctx.Sender() <! ex.Message

                        ignored ()

                    | Msg.UnregistraterXll xll -> 
                        try 
                            ExcelAsyncUtil.QueueAsMacro(fun _ ->
                                ExcelIntegration.UnregisterXLL xll |> ignore
                            )                        
                            ctx.Sender() <! "SUCCESS"
                        with ex ->
                            ctx.Sender() <! ex.Message
                        ignored ()
                ) 
            } |> Async.Start
        member this.AutoClose () =
            ()
