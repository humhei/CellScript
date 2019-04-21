namespace CellScript.Server.Fcs
open FSharp.Compiler.SourceCodeServices
open System.Reflection
open System.Runtime.Loader
open System.IO
open System.Collections.Concurrent
open System
open Newtonsoft.Json
open System.Runtime.CompilerServices
open CellScript.Core.Types
open System.Linq.Expressions
open CellScript.Core.Conversion

type private BoxedFunction = private BoxedFunction of obj
with 
    member x.Value = 
        let (BoxedFunction value) = x
        value

[<RequireQualifiedAccess>]
module private BoxedFunction = 
    open System.Reflection

    /// retrive anonymous in module
    let ofRuntimeMemberInfo(memberInfo: MemberInfo) =
        let memberInfoType = memberInfo.GetType()
        if memberInfoType.FullName <> "System.RuntimeType" then failwithf "member info named %s is not runtime type" memberInfoType.Name
        let getValue propertyName = memberInfo.GetType().GetProperty(propertyName).GetValue(memberInfo)
        let baseType = getValue "BaseType" :?> Type
        assert (baseType.Name.StartsWith "FSharpFunc`")
        let constructor = getValue "DeclaredConstructors" :?> seq<ConstructorInfo> |> Seq.exactlyOne
        constructor.Invoke(null) |> BoxedFunction

[<RequireQualifiedAccess>]
module private MethodInfo =
    let ofBoxedFunction (BoxedFunction v) =
        v.GetType().GetMethods() |> Seq.find (fun m -> 
            m.Name = "Invoke" && m.GetParameters().Length = 1
        )

[<RequireQualifiedAccess>]
module private LambdaExpression =

    let ofBoxedFunction boxedFunction =
        let methodInfo = MethodInfo.ofBoxedFunction boxedFunction
        let (BoxedFunction value) =  boxedFunction
        let instance = Expression.Constant value
        let param = 
            methodInfo.GetParameters()
            |> Array.map (fun pi -> Expression.Parameter(pi.ParameterType,pi.Name))
        let paramExprs = param |> Seq.cast<Expression>
        Expression.Lambda(Expression.Call(instance,methodInfo,paramExprs),param)


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

    let private codeBinaryTable = ConcurrentDictionary<string,Result<string, FSharpErrorInfo []>>()

    let private checker =
        lazy FSharpChecker.Create()

    let private useContext dll f =
        let loadContext = HostAssemblyLoadContext(dll)

        let assembly = loadContext.LoadFromAssemblyPath(dll)

        let result = f assembly

        loadContext.Unload()

        result

    let private boxedFunctionAndMethodInfo (assembly: Assembly) =
        let userCodeModule =
            assembly.GetExportedTypes()
            |> Array.find (fun tp ->
                let assemblyName = Path.GetFileNameWithoutExtension assembly.CodeBase
                System.String.Compare(tp.Name, assemblyName, true) = 0
            )

        match userCodeModule.GetMembers(BindingFlags.DeclaredOnly ||| BindingFlags.NonPublic) with 
        | [| memberInfo |] ->
            let boxedFunction = BoxedFunction.ofRuntimeMemberInfo memberInfo
            let methodInfo = MethodInfo.ofBoxedFunction boxedFunction
            boxedFunction, methodInfo
        | _ -> failwith "only one exported member info is supported currently"

    let private parameterConversionMapping = new ConcurrentDictionary<Type, MethodInfo>()

    let private convertToParam (paramType: Type) (xlRef: SerializableExcelReference) =
        let asSingleBaseTypeInsance convert =
            convert xlRef.Content.[0,0] 
            |> box

        match paramType.FullName with 
        | "System.String" ->
            asSingleBaseTypeInsance string
        | "System.Int32" ->
            asSingleBaseTypeInsance (string >> Int32.Parse)
        | "System.Double" ->
            asSingleBaseTypeInsance (string >> Double.Parse)
        | _ ->
            let convertMethodInfo = 
                parameterConversionMapping.GetOrAdd(paramType, fun _ ->
                    paramType.GetMethod("Convert", BindingFlags.Static ||| BindingFlags.Public)
                )

            let paramValue = 
                match convertMethodInfo.GetParameters() with 
                | [| param |] ->
                    match param.ParameterType.FullName with 
                    | "System.Object[,]" ->
                        box xlRef.Content
                    | _ when param.ParameterType = typeof<SerializableExcelReference> ->
                        box xlRef
                    | fullName -> 
                        failwithf "cannot convert serializableExcelReference to %s, Please implement static Convert method first" fullName
                | _ ->
                    failwithf "only one parameter in %s is supported now" convertMethodInfo.Name

            convertMethodInfo.Invoke(null,[|paramValue|])

    let private convertToReturn (value: obj) =
        let iToArray2D = 
            match value with 
            | :? string as value  ->
                String value
                :> IToArray2D
            | :? float as value ->
                Float value
                :> IToArray2D
            | :? int as value ->
                Int value
                :> IToArray2D
            | :? IToArray2D as v ->
                v
            | _ -> failwithf "cannot convert %s to obj[,], please implement IToArray2D first" (value.GetType().FullName)
        iToArray2D.ToArray2D()

    [<MethodImpl(MethodImplOptions.NoInlining)>]
    let eval (xlRef: SerializableExcelReference) (code: string): obj[,] =
        let scriptsDir = Path.Combine(Directory.GetCurrentDirectory(), "Scripts")

        let dll =
            codeBinaryTable.GetOrAdd(code, fun code ->
                let script =
                    let tmp = Path.GetRandomFileName()
                    let tmpFsx = Path.ChangeExtension(tmp,".fsx")
                    Directory.CreateDirectory(scriptsDir)
                    |> ignore
                    Path.Combine(scriptsDir, tmpFsx)

                File.WriteAllText(script,code,Text.Encoding.UTF8)

                let dll = Path.ChangeExtension(script, "dll")

                let errors, exitCode = checker.Value.Compile(compilerArgvWithDebuggerInfo dll script) |> Async.RunSynchronously
                if exitCode <> 0 then
                    printfn "ERROS: %A" errors
                    Result.Error errors
                else
                    Result.Ok dll
            )

        match dll with
        | Result.Ok dll ->
            useContext dll (fun assembly ->
                
                let boxedFunction, methodInfo = boxedFunctionAndMethodInfo assembly
                let paramValue = 
                    match methodInfo.GetParameters() with 
                    | [| param |] -> convertToParam param.ParameterType xlRef
                    | _ -> failwithf "only one parameter in %s is supported now" methodInfo.Name
                
                let returnValue = methodInfo.Invoke(boxedFunction.Value, [| paramValue |])

                convertToReturn returnValue
            )

        | Result.Error errors ->
            let errorText =
                errors
                |> Array.map (fun error -> error.Message)
                |> String.concat "\n"
            
            array2D [[errorText]]
