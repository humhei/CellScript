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
open CellScript.Core
open LiteDB
open LiteDB.FSharp
open LiteDB.FSharp.Extensions

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

        let validNamingPath = xlRef.WorkbookPath.Replace('/','-').Replace('\\','-').Replace(':','-').Replace('.','_')

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


    type CompileHistory =
        { Code: string
          /// dll, error
          CompileResult: Result<string, string>}

    [<CLIMutable>]
    type XlRefCache =
        { Id: int
          XlRef: SerializableExcelReferenceWithoutContent
          CompileHistorys: CompileHistory list }

    [<RequireQualifiedAccess>]
    module private XlRefCache =
        let private addOrUpdateHistory history (xlRefCache: XlRefCache) =
            match List.tryFind (fun historyInCache -> historyInCache.Code = history.Code) xlRefCache.CompileHistorys with 
            | Some _ ->
                let leftHistories = 
                    List.filter (fun historyInCache -> 
                        historyInCache.Code <> history.Code) xlRefCache.CompileHistorys
                { xlRefCache with CompileHistorys = history :: leftHistories}

            | None ->
                let deepth = 10
                let compileHistorys = xlRefCache.CompileHistorys
                if compileHistorys.Length < deepth then 
                    { xlRefCache with CompileHistorys = history :: compileHistorys }
                else 
                    { xlRefCache with 
                        CompileHistorys = history :: List.take (compileHistorys.Length - 1) compileHistorys }       
        

        let private dbPath = "CellScript.Server.Fcs.db"
        let private mapper = FSharpBsonMapper()

        let private useCollection (f: LiteCollection<'T> -> 'result) = 
            use db = new LiteDatabase(dbPath, mapper)
            let collection = db.GetCollection<'T>()
            f collection



        let getOrAddOrUpdateInLiteDb xlRef code valueFactory = 
            useCollection (fun col ->
                match col.tryFindOne(fun cache -> cache.XlRef = xlRef) with 
                | Some cache ->
                    match List.tryFind (fun history -> history.Code = code) cache.CompileHistorys with 
                    | Some history ->
                        match history.CompileResult with 
                        | Result.Ok dll when not (File.Exists dll) ->
                            let history = valueFactory()
                            let newCache = addOrUpdateHistory history cache
                            col.Update(newCache) |> ignore
                            newCache
                        | _ ->
                            cache
                    | None -> 
                        let compileHistory = valueFactory()
                        let newCache = addOrUpdateHistory compileHistory cache
                        col.Update(newCache) |> ignore
                        newCache
                | None ->
                    let compileHistory = valueFactory()
                    let cache = 
                        { Id = 0 
                          XlRef = xlRef 
                          CompileHistorys = [compileHistory] }
                    col.Insert(cache) |> ignore
                    cache
            )


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

    let textToMD5 (text: string) =
        Encoding.UTF8.GetBytes(text)
        |> md5

    let private evalByResult xlRef (script: string) dll =
        match dll with
        | Result.Ok dll ->
            useContext dll (fun assembly ->

                let boxedFunction, methodInfo = 
                    let boxedFunctionAndMethodInfo (assembly: Assembly) =
                        let userCodeModule =
                            let moduleName = Path.GetFileNameWithoutExtension script
                            assembly.GetExportedTypes()
                            |> Array.find (fun tp ->
                                let assemblyName = moduleName
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
                        | :? string as value ->
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
                    iToArray2D

                convertToReturn returnValue
            )

        | Result.Error (errorText: string) ->
    
            String errorText :> IToArray2D

    let createAgent system: IActorRef<CompilerMsg> = 
        spawnAnonymous system (propsPersist (fun ctx ->
            let rec loop (state: Map<SerializableExcelReferenceWithoutContent, XlRefCache>) = actor {
                let! msg = ctx.Receive() 
                match msg with
                | CompilerMsg.Eval (scriptsDir, xlRef, code) ->

                    let xlRefPos = SerializableExcelReferenceWithoutContent.ofSerializableExcelReference xlRef
                    let script = SerializableExcelReferenceWithoutContent.getFsxPath scriptsDir xlRefPos
                    
                    let dll = 
                        let md5 = md5 (Encoding.UTF8.GetBytes (script + code))
                        scriptsDir </> md5 + ".dll"

                    let compileToNewByState state =
                        let cache = 
                            XlRefCache.getOrAddOrUpdateInLiteDb xlRefPos code (fun () ->
                                File.WriteAllText(script, code, Text.Encoding.UTF8)
                                let result = 
                                    let errors, exitCode = 
                                        let compilerArgWithDebuggerInfo dll script =
                                            [| "fsc.exe"; "-o"; dll; "-a"; script;"--optimize-";"--debug"; |]

                                        checker.Value.Compile(compilerArgWithDebuggerInfo dll script) |> Async.RunSynchronously
                                
                                    if exitCode <> 0 then
                                        let errors = sprintf "ERROS: %A" errors
                                        Result.Error errors
                                    else
                                        Result.Ok dll 
                                { Code = code
                                  CompileResult = result }

                            )

                        let newState = Map.add xlRefPos cache state

                        let history =
                            cache.CompileHistorys |> List.find (fun history -> history.Code = code)

                        ctx.Sender() <! (evalByResult xlRef script history.CompileResult)

                        newState

                    return! loop (compileToNewByState state)

            }
            loop Map.empty
        ))


