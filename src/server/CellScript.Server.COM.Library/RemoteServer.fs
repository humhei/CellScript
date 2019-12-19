[<RequireQualifiedAccess>]
module CellScript.Server.COM.Library.RemoteServer
open Akkling
open System.IO
open CellScript.Core.COM
open Microsoft.Office.Interop.Excel

let private saveXlsToXlsx paths =
    let paths = 
        paths 
        |> List.filter(fun path ->
            Path.GetExtension(path) = ".xls" &&
                let xlsxFile = Path.ChangeExtension (path,".xlsx")
                not (File.Exists xlsxFile)
        )

    match paths with 
    | [] -> Result.Ok "Paths length is empty"
    | _ ->
        let app = new ApplicationClass()
        app.Visible <- false
        try 
            let result = 
                paths
                |> List.map (fun path ->
                    let workbook = app.Workbooks.Open(path)
                    let xlsxFile = Path.ChangeExtension(path,".xlsx")
                    workbook.SaveAs(xlsxFile,51)
                    sprintf "Successfully convert %s to %s" path xlsxFile
                )
                |> String.concat "\n"
            app.Quit()
            Result.Ok result
        with ex ->
            app.Quit()
            Result.Error ex.Message

let run (logger: NLog.FSharp.Logger) =
    let handleMsg (ctx: Actor<_>) msg customModel =
        logger.Info "[CellScriptComServer] Recieve msg %A" msg
        try 
            match msg with   
            | COMMsg.SaveXlsToXlsx paths ->
                let result = saveXlsToXlsx paths
                ctx.Sender() <! result

        with ex ->
            ctx.Sender() <! Result.Error (ex.Message)
    COMServer.createAgent (COMRouted.port) () handleMsg
    |> ignore
