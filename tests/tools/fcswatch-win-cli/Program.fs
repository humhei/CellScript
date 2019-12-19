// Learn more about F# at http://fsharp.org
module Program
open Argu
open FcsWatch.Core
open FcsWatch.Binary
open Fcswatch.Win.Cli.Utils
open FcsWatch.AutoReload.ExcelDna.Core
open Fcswatch.Win.Cli
open Fake.Core
open Fcswatch.Win.Cli.ExcelDna
open FcsWatch.AutoReload.ExcelDna.Core.Remote
open System.Threading


type CoreArguments = FcsWatch.Cli.Share.Arguments


[<RequireQualifiedAccess>]
type ReloadByArguments =
    | Xll
    | AddIn

type Arguments =
    | Working_Dir of string
    | Project_File of string
    | Debuggable
    | Logger_Level of Logger.Level
    | No_Build
    | XLL_ServerPort of int
    | XLL_ClientPort of int
    | ExcelDna of ReloadByArguments
with 
    member x.AsCore =
        match x with 
        | Working_Dir arg  -> CoreArguments.Working_Dir arg |> Some
        | Project_File arg -> CoreArguments.Project_File arg |> Some
        | Debuggable -> CoreArguments.Debuggable |> Some
        | Logger_Level arg -> CoreArguments.Logger_Level arg |> Some
        | No_Build -> Some CoreArguments.No_Build
        | _ -> None

    interface IArgParserTemplate with
        member x.Usage =
            match x with 
            | ExcelDna _ -> "develop excel dna plugin, reload and load plugin by <xll|addin>, default is xll"
            | XLL_ServerPort _ -> "xll server port, default is 8090"
            | XLL_ClientPort _ -> "xll client port, default is 0"
            | _ -> ((x.AsCore.Value) :> IArgParserTemplate).Usage




[<RequireQualifiedAccess>]
module ReloadByArguments =
    let toReloadBy (results: ParseResults<Arguments>) projectPath = function
        | ReloadByArguments.Xll ->
            let project = Project.create projectPath
            let client = 
                let clientPort = results.GetResult(XLL_ClientPort,0)
                let serverPort = results.GetResult(XLL_ServerPort,8090)
                Client.create clientPort serverPort
            ExcelDna.ReloadBy.Xll (project, client)

        | ReloadByArguments.AddIn -> ExcelDna.ReloadBy.AddIn

[<EntryPoint>]
let main argv =
    let coreParser = FcsWatch.Cli.Share.parser


    let parser = ArgumentParser.Create<Arguments>(programName = "fcswatch-win.exe")
    let results = parser.Parse argv

    let processResult = 
        results.GetAllResults()
        |> List.choose (fun (result: Arguments) -> result.AsCore)
        |> coreParser.ToParseResults
        |> FcsWatch.Cli.Share.processParseResults [||]

    let developmentTarget = 
        match results.TryGetResult ExcelDna with 
        | Some reloadByArg -> 
            let reloadBy = ReloadByArguments.toReloadBy results processResult.ProjectFile reloadByArg
            ExcelDna.plugin reloadBy processResult.Config.DevelopmentTarget processResult.ProjectFile
        | None ->
            processResult.Config.DevelopmentTarget

    let config =
        { processResult.Config with 
            DevelopmentTarget = developmentTarget }

    runFcsWatcher config processResult.ProjectFile

    0 // return an integer exit code
