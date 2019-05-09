namespace CellScript.Client.Core
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

[<RequireQualifiedAccess>]
type ApiLambda =
    | Function of LambdaExpression
    | Command of (Command * LambdaExpression)

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
    | UpdateCellValues of xlRefs: SerializableExcelReference list

type NetCoreClient<'Msg> = 
    { Value: Client<'Msg> 
      GetCaller: unit -> CommandCaller }
with 

    member x.RemoteServer = 
        x.Value.RemoteServer

    member x.Logger = x.Value.Logger

    /// internal helper function
    member internal x.InvokeFunction (message): Async<obj[,]> = 
        async {     
            let! (result: obj[,]) = x.RemoteServer <? (message)

            let flat2Darray array2D = 
                [| for x in [0..(Array2D.length1 array2D) - 1] do 
                                yield array2D.[x, *] |]

            let resultPrintable = 
                result 
                |> flat2Darray

            x.Logger.Info(sprintf "CellScript Client get respose %A" resultPrintable)
            return result

        } 

    member internal x.InvokeCommand (message) = 
        async {
            match! x.RemoteServer <? message with
            | Result.Ok () -> x.Logger.Info (sprintf "Invoke command %A succesfully" message)
            | Result.Error error -> x.Logger.Error (sprintf "Erros when invoke command %A: %s" message error)
        } |> Async.Start

        0


[<RequireQualifiedAccess>]
module NetCoreClient =   


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
                        Command (method, commandMapping.Value.[uci], parameter)
                    
                    elif paramType.IsGenericType && paramType.GetGenericTypeDefinition() = typedefof<CellScriptEvent<_>> then
                        Event

                    else
                        method.CustomAttributes
                        |> Seq.tryFind (fun attr -> attr.AttributeType = typeof<SubMsgAttribute>)
                        |> function

                            | Some _ -> InnerMsg (paramType)
                            | None -> FunctionParams (parameters)

                | [||] -> failwithf "at least one parameter should be defined in method %s when registered as a function" method.Name 
                | _ -> FunctionParams (parameters)

        let private nullValue = Expr.Value(null, typeof<unit>)

        let ofUci (client: NetCoreClient<'Msg>) (uci: UnionCaseInfo) =
            let clientValue = Expr.Value client
            let callMethodInfo = typeof<NetCoreClient<'Msg>>.GetMethod("InvokeFunction",BindingFlags.Instance ||| BindingFlags.NonPublic)
            let actionMethodInfo = typeof<NetCoreClient<'Msg>>.GetMethod("InvokeCommand",BindingFlags.Instance ||| BindingFlags.NonPublic)

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
                        if uciAccum.Length > 1 then 
                            let head = uciAccum.Head
                            head.DeclaringType.Name + "." + uci.Name
                        else uci.Name

                    flattenParamtersWithName lambdaName methodCall

                match uci with 
                | InnerMsg paramType ->
                    FSharpType.GetUnionCases paramType
                    |> Array.collect (fun uci -> loop (uci :: uciAccum) uci)

                | FunctionParams (parameters) ->
                    let csharLambda = 
                        createLambda callMethodInfo parameters
                        |> toCSharpLambda

                    [| ApiLambda.Function csharLambda |]

                | Command (_, command, parameter) ->
                    let lambda = createLambda actionMethodInfo [| parameter |]
                    let messageLambda = 

                        Expr.Application(lambda, commandCallerApp)

                    let csharpLambda = toCSharpLambda messageLambda

                    [| ApiLambda.Command (command, csharpLambda) |]

                | Event ->
                    [||]

            loop [uci] uci

    let apiLambdas (client: NetCoreClient<'Msg>) =
        let tp = typeof<'Msg>
        if FSharpType.IsUnion tp then
            FSharpType.GetUnionCases tp
            |> Array.collect (ApiLambda.ofUci client)
        else failwithf "type %s is not an union type" tp.FullName

    let create seedPort getCaller processCustomMsg = 
        let client = 
            Client.create seedPort (fun logger msg ->
                match msg with 
                | ClientMsg.Exec (toolName, args, workingDir) ->
                    let tool = 
                        match ProcessUtils.tryFindFileOnPath toolName with 
                        | Some tool -> tool
                        | None -> 
                            match Environment.environVarOrNone toolName with 
                            | Some tool -> tool
                            | None -> failwithf "cannot find %s on path" toolName

                    let exec tool args dir =
                        let argLine = Args.toWindowsCommandLine args
                        logger.Info (sprintf "%s %s" tool argLine)
                        let result =
                            args
                            |> CreateProcess.fromRawCommand tool
                            |> CreateProcess.withWorkingDirectory dir
                            |> Proc.run

                        if result.ExitCode <> 0
                        then failwithf "Error while running %s with args %A" tool (List.ofSeq args)

                    exec tool args workingDir
                    Ignore

                | ClientMsg.UpdateCellValues xlRefs ->
                    processCustomMsg logger (EndPointClientMsg.UpdateCellValues xlRefs)
            )

        let client =
            { client with 
                RemoteServer = 
                    let remoteServer = client.RemoteServer

                    { new ICanTell<_> with
                        member this.Ask(arg1, ?arg2: TimeSpan) = async {
                            let! (result: Result<obj , string>) = remoteServer.Ask(arg1, arg2)
                            match result with 
                            | Result.Ok response -> 
                                return unbox response
                            | Result.Error errorMsg ->
                                let errorMsg = array2D [[box errorMsg]]
                                return unbox errorMsg
                        }

                        member this.Tell(arg1, arg2): unit = 
                            remoteServer <! (arg1)

                        member this.Underlying = 
                            remoteServer.Underlying }  


            }

        { Value = client 
          GetCaller = getCaller }
