module CellScript.Shared
open Microsoft.FSharp.Linq.RuntimeHelpers
open System.Linq.Expressions
open System
open CellScript.Core
open Quotations.DerivedPatterns
open Quotations.Patterns
open Microsoft.FSharp.Quotations
open FSharp.Quotations.Evaluator


let host = "127.0.0.1"

type RemotingPort = RemotingPort of int
with 
    member x.Value = 
        let (RemotingPort v) = x
        v

    member x.HostName = sprintf "http://%s:%d" host x.Value

    member x.RouteBuilder typeName methodName =
        sprintf "/api/%s/%s" typeName methodName

    member x.UrlBuilder typeName methodName =
        let route =
            x.RouteBuilder typeName methodName |> sprintf "%s%s" x.HostName
        route

[<RequireQualifiedAccess>]
type LambdaExpressionStatic =
    static member ofFun(exp: Expression<Func<_,_>>) =
        exp :> LambdaExpression

    static member ofFun2(exp: Expression<Func<_,_,_>>) =
        exp :> LambdaExpression

type ExpressionStatic =
    static member ofFun(exp: Expression<Func<_,_>>) =
        exp

module Fcs =
    
    type EvalParam = 
        { Input: string 
          Code: string }

    type ICellScriptFcsApi =
        { eval : EvalParam  -> Async<string>
          test: unit -> Async<string> }

module UDF =
    let port = RemotingPort 9000

    type UnitParam = 
        { unit: unit }

    let call (expr) =
        fun call -> call expr

    type ICellScriptUDFApi =
        { testFromCellScriptServer: string -> Async<string>
          testFromCellScriptServer2: int -> int -> Async<string>
          testTableParameter: Table -> Async<string>
          testTableTwoParameter: Table -> Table -> Async<string> }


    type Call<'a> = (Quotations.Expr<ICellScriptUDFApi -> Async<'a>> -> Async<'a>)

    let transfer2 desired call (expr: Expr<ICellScriptUDFApi -> 'a -> 'b -> Async<'c>>) =
        let expr = expr :> Expr

        let rec loop accum expr =
            match expr with 
            | Lambda (var1, body) ->
                loop (var1 :: accum) body

            | Application (app, param) ->
                loop accum app

            | PropertyGet (param, propInfo, _) ->
                let param = param.Value

                let call (input1: 'a) (input2: 'b) =
                    let param1 = Expr.Value input1
                    let propertyGet = Expr.PropertyGet (param, propInfo, [])
                    let app1 = Expr.Application (propertyGet, param1)
                    let param2 = Expr.Value input2
                    let app2 = Expr.Application (app1, param2)
                    let lambda = 
                        match param with 
                        | Var var ->
                            Expr.Lambda (var, app2)
                        | _ -> failwith "not implemtented"
                    call (unbox lambda) 

                let l = accum.Length

                let app = 
                    let func = Expr.ValueWithName (call, propInfo.PropertyType, "call") 
                    (func, List.rev accum.[0..l-2]) ||> List.fold (fun func var ->
                        Expr.Application (func, Expr.Var var)
                    )

                (app, accum.[0..l-2]) ||> List.fold (fun expr var ->
                    Expr.Lambda (var, expr)
                )
            | _ -> failwith "Not implemented"

        loop [] expr
    //type Ext = Ext
    //    with
    //        static member inline Bar (ext : Ext, call, expr : Quotations.Expr<ICellScriptUDFApi -> ^a -> Async<'b>> when ^a: (static member Convert : obj[,] -> ^a)) = 
    //            transfer call (expr :> Expr) 

    //        static member inline Bar (ext : Ext, call, flt : Quotations.Expr<ICellScriptUDFApi -> ^a -> ^b -> Async<'b>> when ^a: (static member Convert : obj[,] -> ^a) and  ^b: (static member Convert : obj[,] -> ^b)) = 
    //            flt
    //        //static member Bar (ext : Ext, flt : float32) = 1.0f + flt


    //let inline bar (call: Call<_>) (x : ^a) =
    //    ((^b or ^a) : (static member Bar : ^b * Call<_> * ^a -> Expr) (Ext, call, x))

    //let inline bar (expr : Quotations.Expr<ICellScriptUDFApi -> ^a -> Async<string>> when ^a: (static member Convert : obj[,] -> ^a)) =

        //()

    let apiLambdas call =

        let s = 
            let call input1 input2 =
                call <@ fun api -> api.testTableTwoParameter input1 input2 @>

            <@ fun input1 input2 -> call input1 input2 @>


        let exprTwoParams = <@ fun api input1 input2 -> api.testTableTwoParameter input1 input2 @>
        let s2 = 
            <@ fun input1 input2 -> call <@ fun api -> api.testTableTwoParameter input1 input2 @> @>
        
        let transferd = transfer2 s call exprTwoParams
        let compiled = transferd.CompileUntyped()
        let lambda = LeafExpressionConverter.QuotationToExpression transferd :?> MethodCallExpression
        //let s2Transferd = transfer s2 exprTwoParams
        let s3 = <@ fun input1 input2 -> <@ fun api -> api.testTableTwoParameter input1 input2 @> @>
        let s4 = <@ fun input1 input2 -> call <@ fun api -> api.testTableTwoParameter input1 input2 @> @>

        //let s = bar call expr
        //let s2 = bar call exprTwoParams
        //let s = expr.ToString()
        
        let s = ""

        [ 
            <@ fun api input -> api.testFromCellScriptServer input @>
        ]