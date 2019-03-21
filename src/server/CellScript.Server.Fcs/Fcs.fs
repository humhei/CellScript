namespace CellScript.Server
open FSharp.Compiler.SourceCodeServices
open System.Reflection
open System.Runtime.Loader
open System.IO
open System.Collections.Concurrent
open System
open Newtonsoft.Json
open CellScript.Core.Registration
open CellScript.Core.Extensions
open System.Runtime.CompilerServices

type Config =
    { WorkingDir: string }

[<RequireQualifiedAccess>]
module Fcs =

    type private HostAssemblyLoadContext(pluginPath: string) =
        inherit AssemblyLoadContext(isCollectible = true)
        let resolver = new AssemblyDependencyResolver(pluginPath)

        override x.Load(name: AssemblyName) =
            let assemblyPath = resolver.ResolveAssemblyToPath(name)
            match assemblyPath with
            | null -> null
            | assemblyPath ->
                printfn "Loading assembly %s into the HostAssemblyLoadContext" assemblyPath
                x.LoadFromAssemblyPath assemblyPath


    let private compilerArgvWithDebuggerInfo dll script =
        [| "fsc.exe"; "-o"; dll; "-a"; script;"--optimize-";"--debug"; |]

    let private codeBinaryTable = ConcurrentDictionary<string,string option>()

    let private checker =
        lazy FSharpChecker.Create()

    [<MethodImpl(MethodImplOptions.NoInlining)>]
    let eval (config: Config) (input: string) (code: string) =
        let scriptsDir = Path.Combine(config.WorkingDir, "Scripts")

        let dll =
            codeBinaryTable.GetOrAdd(code, fun code ->
                let script =
                    let tmp = Path.GetRandomFileName()
                    let tmpFsx = Path.ChangeExtension(tmp,".fsx")
                    Path.Combine(scriptsDir, tmpFsx)

                File.WriteAllText(script,code,Text.Encoding.UTF8)

                let dll = Path.ChangeExtension(script, "dll")

                let errors,exitCode = checker.Value.Compile(compilerArgvWithDebuggerInfo dll script) |> Async.RunSynchronously
                if exitCode <> 0 then
                    printfn "ERROS: %A" errors
                    None
                else
                    Some dll
            )

        match dll with
        | Some dll ->

            let loadContext = HostAssemblyLoadContext(dll)

            let assembly = loadContext.LoadFromAssemblyPath(dll)

            let userCodeModule =
                assembly.GetExportedTypes()
                |> Array.find (fun tp ->
                    let assemblyName = Path.GetFileNameWithoutExtension assembly.CodeBase
                    System.String.Compare(tp.Name, assemblyName, true) = 0
                )

            let scriptLambda =
                userCodeModule.GetMembers(BindingFlags.NonPublic)
                |> Array.tryLast
                |> function
                    | Some memberInfo ->
                        LambdaExpression.ofRuntimeMemberInfo memberInfo
                    | None -> failwithf "Cannot find function in module %s" userCodeModule.FullName


            let castedInput =
                let paramType =
                    let param = Seq.exactlyOne scriptLambda.Parameters
                    param.Type


                let array2D = JsonConvert.DeserializeObject<obj[,]>(input)
                let correspondCastedInput =
                    match array2D.Length with
                    | 0 -> failwithf "empty input %A" array2D
                    | 1 ->
                        let v = array2D.[0,0]
                        let castLambda = CustomParamConversion.register paramType ignore
                        castLambda.Compile().DynamicInvoke(v)
                    | _ ->
                        let castLambda = CustomParamConversion.register paramType ignore
                        castLambda.Compile().DynamicInvoke(array2D)

                correspondCastedInput


            loadContext.Unload()

            let result = scriptLambda.Compile().DynamicInvoke(castedInput)

            let array2DResult =
                let returnLambda = ICustomReturn.register scriptLambda.ReturnType
                returnLambda.Compile().DynamicInvoke result

            JsonConvert.SerializeObject(array2DResult)

        | None ->
            let result =
                [["Error"]]

            JsonConvert.SerializeObject(result)
