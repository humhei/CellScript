namespace CellScript.Core.Client
open System.Reflection
open System.Linq.Expressions
open CellScript.Core
open Akkling
open Microsoft.FSharp.Reflection
open Microsoft.FSharp.Quotations
open Microsoft.FSharp.Linq.RuntimeHelpers
open CellScript.Core.Extensions
open Akka.Util

[<AutoOpen>]
module InternalExprHelper =
    let mutable expected1: Expr option = None
    let mutable expected2: Expr option = None

    [<RequireQualifiedAccess>]
    module private Assembly =
        let getAllExportMethods (assembly: Assembly) =
            assembly.ExportedTypes
            |> Seq.collect (fun exportType ->
                exportType.GetMethods()
            )

    /// internal helper function
    let call (actor: IActorRef<Seriali>) message: obj[,] = 
        actor <? message
        |> Async.RunSynchronously

    let private methods =
        lazy 
            Assembly.GetExecutingAssembly()
            |> Assembly.getAllExportMethods

    let internal lazyCallMethodInfo = 
        lazy 
            methods.Value
            |> Seq.find (fun method -> method.Name = "call")

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
module private LambdaExpression =
    let paramExprsAndReturnType (methodCall: MethodCallExpression) =
        let rec loop accum (methodCall: MethodCallExpression) = 
            let arg =  Seq.exactlyOne methodCall.Arguments :?> LambdaExpression
            let newAccum = Array.append accum (Array.ofSeq arg.Parameters)
            match arg.Body, arg.ReturnType.FullName.StartsWith "Microsoft.FSharp.Core.FSharpFunc`2[[" with 
            | :? MethodCallExpression, false  -> newAccum,arg.ReturnType
            | :? MethodCallExpression as methodCall, true ->
                loop newAccum methodCall
            | _ -> invalidArg "methodName" "methodCall's body should have inner methodCall"

        loop [||] methodCall

    let flattenParamtersWithName (name: string) (methodCall: MethodCallExpression) =
        let paramExprs, returnType = paramExprsAndReturnType methodCall
        let length = paramExprs.Length
        match length with
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
        | 0 -> failwith "Parammeter number is 0"
        | _ -> failwith "Parammeter number more than 5 is not supported now"


    let private (|InnerMsg|Params|) (uci: UnionCaseInfo) =
        let method = 
            uci.DeclaringType.GetMethods()
            |> Array.find (fun methodInfo -> methodInfo.Name = "New" + uci.Name)
        
        let parameters = 
            method.GetParameters()

        match parameters with 
        | [| parameter |] ->
            let paramType = parameter.ParameterType
            let name = paramType.Name
            if name.EndsWith "Msg" 
                && Type.isSomeModuleOrNamespace paramType uci.DeclaringType
                && FSharpType.IsUnion paramType  
            then 
                InnerMsg (paramType)
            else Params (parameters)
        | _ -> Params (parameters)




    let convertParam (lambda: Expr) =
        let expected1 = expected1
        let expected2 = expected2
        lambda

    let ofUci clientAgent (uci: UnionCaseInfo) =
        let rec loop uciAccum uci =
            match uci with 
            | InnerMsg paramType ->
                FSharpType.GetUnionCases paramType
                |> Array.collect (fun uci -> loop (uci :: uciAccum) uci)

            | Params (parameters) ->
                let paramExprs = 
        
                    let paramName (param: ParameterInfo) = 
                        param.Name.TrimStart('_')

                    parameters
                    |> Array.map (fun param ->
                        (Var(paramName param, param.ParameterType))
                    ) 
                    |> List.ofArray

                let callExpr = 
                    let callMethodInfo = lazyCallMethodInfo.Value
                    let arg = 
                        let arguments = List.map Expr.Var paramExprs
                        (arguments, uciAccum) ||> List.fold (fun arguments uci ->
                            [Expr.NewUnionCase (uci, arguments)]
                        )

                    Expr.Call(callMethodInfo, clientAgent :: arg)

                let lambda =
                    (callExpr,List.rev paramExprs)
                    ||> List.fold (fun expr var ->
                        Expr.Lambda(var, expr)
                    )
                    |> convertParam

                let methodCall = LeafExpressionConverter.QuotationToExpression lambda :?> MethodCallExpression

                let lambdaName = 
                    if uciAccum.Length > 1 then 
                        let head = uciAccum.Head
                        head.DeclaringType.Name + "." + uci.Name
                    else uci.Name

                [| flattenParamtersWithName lambdaName methodCall |]

        loop [uci] uci

[<RequireQualifiedAccess>]
module Registration =
    let apiLambdas<'Msg> clientAgent =
        let tp = typeof<'Msg>
        if FSharpType.IsUnion tp then
            FSharpType.GetUnionCases tp
            |> Array.collect (LambdaExpression.ofUci clientAgent)
        else failwithf "type %s is not an union type" tp.FullName