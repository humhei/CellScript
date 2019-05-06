namespace CellScript.Server.Fcs
open FSharp.Compiler.SourceCodeServices
open System.Reflection
open System.Runtime.Loader
open System.IO
open System.Collections.Concurrent
open System
open CellScript.Core.Types
open CellScript.Core.Conversion
open System.Security.Cryptography
open System.Text
open Akkling
open Akkling.Persistence
open Akka.Persistence.LiteDB.FSharp
open OfficeOpenXml
open Fake.IO.FileSystemOperators

#nowarn "0044"


type SerializableExcelReferenceWithoutContent = 
    { ColumnFirst: int
      RowFirst: int
      ColumnLast: int
      RowLast: int
      WorkbookPath: string
      SheetName: string }


[<RequireQualifiedAccess>]
module SerializableExcelReferenceWithoutContent =
    let ofSerializableExcelReference (xlRef: SerializableExcelReference) =
        { ColumnFirst = xlRef.ColumnFirst
          RowFirst = xlRef.RowFirst
          ColumnLast = xlRef.ColumnLast
          RowLast = xlRef.RowLast
          WorkbookPath = xlRef.WorkbookPath
          SheetName = xlRef.SheetName }

    let toSerializableExcelReference content (xlRef: SerializableExcelReferenceWithoutContent) : SerializableExcelReference =
        { ColumnFirst = xlRef.ColumnFirst
          RowFirst = xlRef.RowFirst
          ColumnLast = xlRef.ColumnLast
          RowLast = xlRef.RowLast
          WorkbookPath = xlRef.WorkbookPath
          SheetName = xlRef.SheetName
          Content = content }

    let cellAddress xlRef =
        ExcelAddress(xlRef.RowFirst, xlRef.ColumnFirst, xlRef.RowLast, xlRef.ColumnLast)


    let getFsxPath (scriptsDir: string) (xlRef: SerializableExcelReferenceWithoutContent) =
        let addr = cellAddress xlRef

        let validNamingPath = xlRef.WorkbookPath.Replace('/','_').Replace('\\','_').Replace(':','_')

        let fileName = 
            [validNamingPath; xlRef.SheetName; addr.Address]
            |> String.concat "__"
        scriptsDir </> fileName + ".fsx"

type private BoxedFunction = private BoxedFunction of obj
with 
    member x.Value = 
        let (BoxedFunction value) = x
        value

[<RequireQualifiedAccess>]
module private BoxedFunction = 

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
type CompilerMsg =
    | Eval of scriptsDir: string * xlRef: SerializableExcelReference * code: string

[<RequireQualifiedAccess>]
module Compiler =
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



    let private checker =
        lazy FSharpChecker.Create()

    let private useContext dll f =
        let loadContext = HostAssemblyLoadContext(dll)

        let assembly = loadContext.LoadFromAssemblyPath(dll)

        let result = f assembly

        loadContext.Unload()

        result



    let private parameterConversionMapping = new ConcurrentDictionary<Type, MethodInfo>()

    let private md5 (data : byte array) : string =
        use md5 = MD5.Create()
        (StringBuilder(), md5.ComputeHash(data))
        ||> Array.fold (fun sb b -> sb.Append(b.ToString("x2")))
        |> string

    let private evalByResult xlRef dll =
        match dll with
        | Result.Ok dll ->
            useContext dll (fun assembly ->

                let boxedFunction, methodInfo = 
                    let boxedFunctionAndMethodInfo (assembly: Assembly) =
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

                    boxedFunctionAndMethodInfo assembly

                let paramValue = 
                    match methodInfo.GetParameters() with 
                    | [| param |] -> 
                        let convertToParam (paramType: Type) (xlRef: SerializableExcelReference) =
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

                
                        convertToParam param.ParameterType xlRef
                    | _ -> failwithf "only one parameter in %s is supported now" methodInfo.Name
        
                let returnValue = methodInfo.Invoke(boxedFunction.Value, [| paramValue |])

                let convertToReturn (value: obj) =
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
                        | _ -> failwithf "cannot convert %s to obj[,], please implement interface IToArray2D first" (value.GetType().FullName)
                    iToArray2D.ToArray2D()


                convertToReturn returnValue
            )

        | Result.Error (errors: FSharpErrorInfo []) ->
            let errorText =
                errors
                |> Array.map (fun error -> error.Message)
                |> String.concat "\n"
    
            array2D [[errorText]]



    let createAgent system: IActorRef<CompilerMsg> = 
        spawn system "cellScriptServerFcsCompiler" (propsPersist (fun ctx ->
            let rec loop (state: Map<string, Result<string,FSharpErrorInfo []>>) = actor {
                let! msg = ctx.Receive() : IO<obj>
                match msg with
                | SnapshotOffer snap ->
                    return! loop snap

                | :? CompilerMsg as fcsMsg -> 
                    match fcsMsg with 
                    | CompilerMsg.Eval (scriptsDir, xlRef, code) ->
                        let md5 = (md5 (Encoding.UTF8.GetBytes code))
                        let dll = 
                            let tmpFsx = Path.ChangeExtension(md5, ".dll")
                            Path.Combine(scriptsDir, tmpFsx)


                        match state.TryFind md5, File.Exists dll with 
                        | Some result, true -> 
                            ctx.Sender() <! (evalByResult xlRef result)
                    
                            return! loop state
                        | _ ->
                            let script = 
                                let xlRef = SerializableExcelReferenceWithoutContent.ofSerializableExcelReference xlRef
                                SerializableExcelReferenceWithoutContent.getFsxPath scriptsDir xlRef

                            File.WriteAllText(script,code,Text.Encoding.UTF8)
                            let result = 

                                let errors, exitCode = 
                                    let compilerArgvWithDebuggerInfo dll script =
                                        [| "fsc.exe"; "-o"; dll; "-a"; script;"--optimize-";"--debug"; |]

                                    checker.Value.Compile(compilerArgvWithDebuggerInfo dll script) |> Async.RunSynchronously
                                
                                if exitCode <> 0 then
                                    printfn "ERROS: %A" errors
                                    Result.Error errors
                                else
                                    Result.Ok dll 

                            let newState = Map.add md5 result state

                            saveSnapshot newState ctx

                            ctx.Sender() <! (evalByResult xlRef result)


                            return! loop newState

                | _ -> return Unhandled
            }
            loop Map.empty
        ))
        |> retype


