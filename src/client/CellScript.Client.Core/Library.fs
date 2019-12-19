namespace CellScript.Client.Core
open Shrimp.Akkling.Cluster.Intergraction.Configuration
open Shrimp.Akkling.Cluster.Intergraction
open System.Reflection
open Fake.Core
open System.Linq.Expressions
open Akkling
open Microsoft.FSharp.Reflection
open Microsoft.FSharp.Quotations
open Microsoft.FSharp.Linq.RuntimeHelpers
open Microsoft.FSharp.Quotations.Patterns
open System
open System.Collections.Concurrent
open Akka.Actor
open CellScript.Core
open CellScript.Core.Cluster
open CellScript.Core.Types
open Akka.Event
open System.Collections.Generic
open System.Diagnostics
open System.IO


[<RequireQualifiedAccess>]
type ApiLambda =
    | Function of LambdaExpression
    | Command of (CommandSetting option * LambdaExpression)

[<RequireQualifiedAccess>]
module ApiLambda =
    let asFunction = function 
        | ApiLambda.Function func -> Some func 
        | _ -> None

    let asCommand = function 
        | ApiLambda.Command command -> Some command 
        | _ -> None

[<RequireQualifiedAccess>]
type EndPointClientMsg =
    | UpdateCellValues of xlRefs: ExcelRangeContactInfo list
    | SetRowHeight of xlRef: ExcelRangeContactInfo * height: float


type NetCoreClient<'Msg> = 
    { Value: Client<'Msg> 
      GetCaller: unit -> CommandCaller
      MsgMapping: 'Msg -> 'Msg }
with 

    member x.RemoteServer = 
        x.Value.RemoteServer

    member x.Logger = x.Value.Logger

    /// internal helper function
    member x.InvokeFunction (message): Async<obj[,]> = 
        async {     
            let message = x.MsgMapping message
            let! (result: Result<obj , string>) = x.RemoteServer.Ask (message,x.Value.MaxAskTime)
            let result: obj[,] = 
                match result with 
                | Result.Ok response -> 
                    let itoArray2D: IToArray2D = unbox response
                    itoArray2D.ToArray2D()
                    |> Array2D.map (fun v ->
                        match v with 
                        | null -> box ""
                        | :? DateTime as time -> box (time.ToString())
                        | _ -> v
                    )

                | Result.Error errorMsg ->
                    let errorMsg = array2D [[box errorMsg]]
                    unbox errorMsg

            return result

        } 

    member x.InvokeFunctionSync (message) =
        x.InvokeFunction message
        |> Async.RunSynchronously

    member x.InvokeCommand (message) = 
        async {
            let message = x.MsgMapping message
            match! x.RemoteServer.Ask(message, x.Value.MaxAskTime) with
            | Result.Ok (_: obj) -> x.Logger.Info (sprintf "Invoke command %A succesfully" message)
            | Result.Error error -> x.Logger.Error (sprintf "Erros when invoke command %A: %s" message error)
        } |> Async.Start

        0


[<RequireQualifiedAccess>]
module NetCoreClient =   

    type private LambdaExpressionStatic =
        static member ofFun(v: Expression<Func<'T1,'T2>>) =
            v :> LambdaExpression

    let normalTypesparamConversions =
        let arrayConversion convert =

            LambdaExpressionStatic.ofFun 
                (fun (input: obj[]) ->
                    Array.map (string >> convert) input
                )  
        let listConversion convert =
            LambdaExpressionStatic.ofFun 
                (fun (input: obj[]) ->
                    Array.map (string >> convert) input
                    |> List.ofArray
                )  
        let arrayAndListConversions convert =
            [ arrayConversion convert 
              listConversion convert ]

        [ arrayAndListConversions id
          arrayAndListConversions (Int32.Parse)
          arrayAndListConversions (Double.Parse)
        ] 
        |> List.concat

    [<AutoOpen>]
    module InternalExprHelper =

        [<RequireQualifiedAccess>]
        module private Assembly =
            let getAllExportMethods (assembly: Assembly) =
                assembly.ExportedTypes
                |> Seq.collect (fun exportType ->
                    exportType.GetMethods()
                )



        let private methods =
            lazy 
                Assembly.GetExecutingAssembly()
                |> Assembly.getAllExportMethods

        /// internal helper function
        let invoke2 func arg1 arg2 = 
            FSharpFunc<_,_>.InvokeFast (func, arg1, arg2)

        /// internal helper function
        let invoke3 func arg1 arg2 arg3 = 
            FSharpFunc<_,_>.InvokeFast (func, arg1, arg2, arg3)

        /// internal helper function
        let invoke4 func arg1 arg2 arg3 arg4 = 
            FSharpFunc<_,_>.InvokeFast (func, arg1, arg2, arg3, arg4)

        /// internal helper function
        let invoke5 func arg1 arg2 arg3 arg4 arg5 = 
            FSharpFunc<_,_>.InvokeFast (func, arg1, arg2, arg3, arg4, arg5)


        let internal invokeMethodInfos =
            lazy 
                [2..5]
                |> List.map (fun i ->
                    let methodName = sprintf "invoke%d" i
                    let method = 
                        methods.Value
                        |> Seq.find (fun method -> method.Name = methodName)

                    (i,method)
                )
                |> Map.ofList



    [<RequireQualifiedAccess>]
    module ApiLambda =
        let private paramExprsAndReturnType (methodCall: MethodCallExpression) =
            let rec loop accum (methodCall: MethodCallExpression) = 
                match methodCall.Type.FullName.StartsWith "Microsoft.FSharp.Core.FSharpFunc`2[[" with 
                | true ->
                    let arg =  Seq.exactlyOne methodCall.Arguments :?> LambdaExpression
                    let newAccum = Array.append accum (Array.ofSeq arg.Parameters)
                    match arg.Body with 
                    | :? MethodCallExpression as methodCall -> loop newAccum methodCall
                    | _ -> invalidArg "methodName" "methodCall's body should have inner methodCall"

                | false -> accum, methodCall.Type

            loop [||] methodCall

        let private flattenParamtersWithName (name: string) (methodCall: MethodCallExpression) =
            let paramExprs, returnType = paramExprsAndReturnType methodCall
            let length = paramExprs.Length
            match length with
            | _ when length = 0 ->
                let lambda = LambdaExpression.Lambda(methodCall, name, [||])
                lambda
            | _ when length = 1 ->
                let lambda = methodCall.Arguments.[0] :?> LambdaExpression
                LambdaExpression.Lambda(lambda.Body, name, lambda.Parameters)

            | _ when length > 1 && length < 6 ->
                let paramExprsUntyped = paramExprs |> Array.map (fun param -> param :> Expression)
                let methodInfo = 
                    let typeArguments = 
                        let paramTypes = 
                            paramExprs 
                            |> Array.map (fun param -> param.Type)
                        Array.append paramTypes [|returnType|]
                          
                    let invokeMethodInfo = invokeMethodInfos.Value.[length]
                    invokeMethodInfo.MakeGenericMethod(typeArguments)

                let call = 
                    let AllParams = Array.append [|methodCall :> Expression|] paramExprsUntyped
                    LambdaExpression.Call (methodInfo, AllParams)

                let lambda = LambdaExpression.Lambda(call, name, paramExprs)
                lambda
            | _ -> failwith "Parammeter number more than 5 is not supported now"






        let private tryGetConvertMethod =
            let convertParamTypeCache = new ConcurrentDictionary<Type, MethodInfo option>()

            fun (tp: Type) ->
                convertParamTypeCache.GetOrAdd(tp, fun _ ->
                    let method = 
                        tp.GetMethods()
                        |> Array.tryFind (fun methodInfo -> methodInfo.Name = "Convert" && methodInfo.GetParameters().Length = 1)
                    method
                )

        let private getConvertParamType tp = 
            tryGetConvertMethod tp |> Option.map (fun method ->
                method.GetParameters().[0].ParameterType
            )


        let private convertParam (lambda: Expr) =

            let fixLambda (var: Var) =
                match getConvertParamType var.Type with 
                | Some convertParamType ->
                    Var(var.Name, convertParamType)
                | None -> var

            let fixInner newVar (var: Var) =
                match tryGetConvertMethod var.Type with 
                | Some convertMethod ->
                    Expr.Call(convertMethod, [Expr.Var newVar])
                | None -> Expr.Var var

            let rec loop varAccums expr =  
                match expr with 
                | Lambda (var, expr) ->
                    let newVar = fixLambda var
                    let newAccums = Map.add var newVar varAccums
                    Expr.Lambda(newVar, loop newAccums expr)
                | Call (inst, methodInfo, args) ->
                    match inst with 
                    | Some inst ->
                        Expr.Call(inst, methodInfo, List.map (loop varAccums) args)
                    | None -> Expr.Call(methodInfo, List.map (loop varAccums) args)

                | NewUnionCase (uci, args) ->
                    Expr.NewUnionCase (uci, List.map (loop varAccums) args)
                | Var (expr) ->
                    fixInner varAccums.[expr] expr
                | _ -> expr

            loop Map.empty lambda

        let private hasUciImplementAttribute (attributeType: Type) (uci: UnionCaseInfo) =
            let method = 
                uci.DeclaringType.GetMethods()
                |> Array.find (fun methodInfo -> methodInfo.Name = "New" + uci.Name)

            method.CustomAttributes 
            |> Seq.exists (fun attr -> 
                attr.AttributeType = attributeType)
            

        let private (|InnerMsg|FunctionParams|Command|Event|) (uci: UnionCaseInfo) =
                let method = 
                    uci.DeclaringType.GetMethods()
                    |> Array.find (fun methodInfo -> methodInfo.Name = "New" + uci.Name)

                let parameters = 
                    method.GetParameters()

                match parameters with 
                | [| parameter |] ->
                    let paramType = parameter.ParameterType
                    if paramType = typeof<CommandCaller> then
                        Command (method, parameter)
                    
                    elif paramType.IsGenericType && paramType.GetGenericTypeDefinition() = typedefof<CellScriptEvent<_>> then
                        Event

                    else
                        if hasUciImplementAttribute typeof<SubMsgAttribute> uci
                            || hasUciImplementAttribute typeof<NameInheritedSubMsgAttribute> uci 
                        then 
                            InnerMsg paramType
                        else
                            FunctionParams (parameters)



                | [||] -> failwithf "at least one parameter should be defined in method %s when registered as a function" method.Name 
                | _ -> FunctionParams (parameters)

        let private nullValue = Expr.Value(null, typeof<unit>)

        let ofUci (log: NLog.FSharp.Logger) (commandSettings: IDictionary<UnionCaseInfo, CommandSetting>) (client: NetCoreClient<'Msg>) (uci: UnionCaseInfo) =
            let clientValue = Expr.Value client
            let callMethodInfo = typeof<NetCoreClient<'Msg>>.GetMethod("InvokeFunction",BindingFlags.Instance ||| BindingFlags.Public)
            let callMethodSyncInfo = typeof<NetCoreClient<'Msg>>.GetMethod("InvokeFunctionSync",BindingFlags.Instance ||| BindingFlags.Public)
            let actionMethodInfo = typeof<NetCoreClient<'Msg>>.GetMethod("InvokeCommand",BindingFlags.Instance ||| BindingFlags.Public)

            let commandCallerApp =
                let commandCaller = <@ fun () -> client.GetCaller() @>
                Expr.Application(commandCaller, nullValue)

            let rec loop uciAccum (uci: UnionCaseInfo) =

                let createLambda methodInfo parameters =
                    let paramExprs = 
        
                        let paramName (param: ParameterInfo) = 
                            param.Name.TrimStart('_')

                        parameters
                        |> Array.map (fun param ->
                            (Var(paramName param, param.ParameterType))
                        ) 
                        |> List.ofArray

                    let callExpr = 
                        let args = 
                            let msg = 
                                let arguments = List.map Expr.Var paramExprs
                                (arguments, uciAccum) ||> List.fold (fun arguments uci ->
                                    [Expr.NewUnionCase (uci, arguments)]
                                )
                            msg

                        Expr.Call(clientValue
                        , methodInfo, args)

                    (callExpr,List.rev paramExprs)
                    ||> List.fold (fun expr var ->
                        Expr.Lambda(var, expr)
                    )
                    |> convertParam

                let toCSharpLambda lambda =
                    let methodCall = LeafExpressionConverter.QuotationToExpression lambda :?> MethodCallExpression

                    let lambdaName = 
                        match uciAccum with 
                        | [uci] -> uci.Name
                        | m :: n :: _ ->
                            if hasUciImplementAttribute typeof<NameInheritedSubMsgAttribute> n then 
                                n.Name + "." + m.Name
                            else 
                                let declaringTypeNameWithoutEndingWithMsg = 
                                    let declaringTypeName = m.DeclaringType.Name
                                    if declaringTypeName.EndsWith("Msg",StringComparison.OrdinalIgnoreCase) then 
                                        declaringTypeName.Substring(0, declaringTypeName.Length - 3)
                                    else declaringTypeName
                                declaringTypeNameWithoutEndingWithMsg + "." + m.Name

                        | _ -> failwith "Uci accum is empty"

                    flattenParamtersWithName lambdaName methodCall

                match uci with 
                | InnerMsg paramType ->
                    FSharpType.GetUnionCases paramType
                    |> Array.collect (fun uci -> loop (uci :: uciAccum) uci)

                | FunctionParams (parameters) ->
                    let csharpLambda = 
                        if Array.exists (fun (param: ParameterInfo) -> param.ParameterType = typeof<ExcelRangeContactInfo>) parameters then
                            createLambda callMethodSyncInfo parameters
                            |> toCSharpLambda
                        else 
                            createLambda callMethodInfo parameters
                            |> toCSharpLambda

                    [| ApiLambda.Function csharpLambda |]

                | Command (_, parameter) ->

                    let lambda = createLambda actionMethodInfo [| parameter |]
                    let messageLambda = 

                        Expr.Application(lambda, commandCallerApp)

                    let csharpLambda = toCSharpLambda messageLambda

                    let commandSetting =
                        match commandSettings.TryGetValue uci with 
                        | true, commandSetting ->
                            Some commandSetting
                        | false, _ ->
                            log.Warn "cannot find command setting for %s" uci.Name
                            None

                    [| ApiLambda.Command (commandSetting, csharpLambda) |]

                | Event ->
                    [||]

            loop [uci] uci

    let rec private asCommandUci expr =
        match expr with 
        | Lambda(_, expr) ->
            asCommandUci expr

        | NewUnionCase (uci, exprs) ->
            if exprs.Length <> 1 then failwithf "command %A msg's paramters length should be equal to 1" uci
            else uci
        | _ -> failwithf "command expr %A should be a union case type" expr

    let apiLambdas (log: NLog.FSharp.Logger) (client: NetCoreClient<'Msg>) =
        let tp = typeof<'Msg>
        let commandSettings = 
            match tp.GetMethod("CommandSettings",BindingFlags.Public ||| BindingFlags.Static) with 
            | null -> 
                log.Info "cannot find static member CommandSettings in %s" tp.FullName
                []
            | commandSettingsMethodInfo ->
                unbox (commandSettingsMethodInfo.Invoke(null, [||]))
        
        let commandSettings =
            commandSettings
            |> List.map (fun (expr, command) -> asCommandUci expr, command)
            |> dict

        if FSharpType.IsUnion tp then
            FSharpType.GetUnionCases tp
            |> Array.collect (ApiLambda.ofUci log commandSettings client)
        else failwithf "type %s is not an union type" tp.FullName

    let create seedPort getCommandCaller processCustomMsg msgMapping = 
        let client = 
            Client.create seedPort (fun logger msg ->
                match msg with 
                | ClientCallbackMsg.Exec (toolName, args, workingDir) ->
                    let tool = 
                        match ProcessUtils.tryFindFileOnPath toolName with 
                        | Some tool -> tool
                        | None -> 
                            match Environment.environVarOrNone toolName with 
                            | Some tool -> tool
                            | None -> failwithf "cannot find %s on path" toolName

                    let exec tool args dir =
                        let argLine = Args.toWindowsCommandLine args
                        let fullArgLine = (sprintf "%s %s" tool argLine)
                        logger.Info fullArgLine
                        let result =
                            let startInfo = ProcessStartInfo(tool,argLine)
                            startInfo.WorkingDirectory <- dir
                            Process.Start(startInfo)
                        result.WaitForExit()
                        match result.ExitCode with 
                        | -1 when String.Compare (Path.GetFileNameWithoutExtension(tool),"devenv",true) = 0 ->
                            ()
                        | i when i <> 0 ->
                            failwithf "Error while running %s" 
                                (Args.toWindowsCommandLine ([tool] @ List.ofSeq args))
                        | 0 -> ()
                        | _ -> failwith "Invalid token"
                    exec tool args workingDir
                    Ignore

                | ClientCallbackMsg.UpdateCellValues xlRefs ->
                    processCustomMsg logger (EndPointClientMsg.UpdateCellValues xlRefs)
                | ClientCallbackMsg.SetRowHeight (xlRef, height) ->
                    processCustomMsg logger (EndPointClientMsg.SetRowHeight (xlRef,height))

            )

        { Value = client 
          GetCaller = getCommandCaller
          MsgMapping = msgMapping }
