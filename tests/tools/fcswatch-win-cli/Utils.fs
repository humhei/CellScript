module Fcswatch.Win.Cli.Utils

open Fake.IO

open System.IO
open Fake.IO.FileSystemOperators

open System.Xml


[<RequireQualifiedAccess>]
module Path =
    let nomarlizeToUnixCompitiable path =
        let path = (Path.getFullName path).Replace('\\','/')

        let dir = Path.getDirectory path

        let segaments =
            let fileName = Path.GetFileName path
            fileName.Split([|'\\'; '/'|])

        let folder dir segament =
            dir </> segament
            |> Path.getFullName

        segaments
        |> Array.fold folder dir

[<RequireQualifiedAccess>]
type Framework =
    | MultipleTarget of string list
    | SingleTarget of string

[<RequireQualifiedAccess>]
module Framework =
    let (|CoreApp|FullFramework|NetStandard|) (framework: string) =
        if framework.StartsWith "netcoreapp" 
        then CoreApp 
        elif framework.StartsWith "netstandard"
        then NetStandard
        else FullFramework

    let asList = function
        | Framework.MultipleTarget targets -> targets
        | Framework.SingleTarget target -> [target]

    let ofProjPath (projectFile: string) =
        let projectFile = projectFile.Replace('\\','/')
        let doc = new XmlDocument()
        doc.Load(projectFile)
        match doc.GetElementsByTagName "TargetFramework" with
        | frameworkNodes when frameworkNodes.Count = 0 ->
            let frameworksNodes = doc.GetElementsByTagName "TargetFrameworks"
            let frameworksNode = [ for node in frameworksNodes do yield node ] |> List.exactlyOne
            frameworksNode.InnerText.Split(';')
            |> Array.map (fun text -> text.Trim())
            |> List.ofSeq
            |> Framework.MultipleTarget

        | frameWorkNodes ->
            let frameworkNode = [ for node in frameWorkNodes do yield node ] |> List.exactlyOne
            Framework.SingleTarget frameworkNode.InnerText

[<RequireQualifiedAccess>]
type OutputType =
    | Exe
    | Library

[<RequireQualifiedAccess>]
module OutputType =
    let ofProjPath (projPath: string) =
        let doc = new XmlDocument()

        doc.Load(projPath)

        match doc.GetElementsByTagName "OutputType" with
        | nodes when nodes.Count = 0 -> OutputType.Library
        | nodes -> 
            let nodes = 
                [ for node in nodes -> node ]
            nodes |> List.tryFind (fun node ->
                node.InnerText = "Exe"
            )
            |> function 
                | Some _ -> OutputType.Exe
                | None -> OutputType.Library
            

    let outputExt framework = function
        | OutputType.Exe -> 
            match framework with 
            | Framework.FullFramework ->
                ".exe"
            | _ -> ".dll"
        | OutputType.Library -> ".dll"

type Project =
    { ProjPath: string
      OutputType: OutputType
      TargetFramework: Framework }
with
    member x.Name = Path.GetFileNameWithoutExtension x.ProjPath

    member x.Projdir = Path.getDirectory x.ProjPath

    member x.OutputPaths =
        Framework.asList x.TargetFramework
        |> List.map (fun framework ->
            x.Projdir </> "bin/Debug" </> framework </> x.Name + OutputType.outputExt framework x.OutputType
            |> Path.nomarlizeToUnixCompitiable
        )




[<RequireQualifiedAccess>]
module Project =
    let create projPath =
        { OutputType = OutputType.ofProjPath projPath
          ProjPath = projPath
          TargetFramework = Framework.ofProjPath projPath }