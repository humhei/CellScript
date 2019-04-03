namespace CellScript.Client
open ExcelDna.Registration
open System
open CellScript.Core.Registration
open System.Linq.Expressions
open CellScript.Core
open Akkling
open CellScript.Core.Client
open FSharp.Quotations
open Akka.Util


[<RequireQualifiedAccess>]
module Registration =

    [<AutoOpen>]
    module Akka =
        let client = System.create "client" <| Configuration.parse Routed.clientConfig

        let clientAgent =
            Routed.remote <- Remote.Client
            spawn client "remote-actor" (Routed.remoteProps())


    let excelFunctions() = 

        expected1 <- 
            let expr = <@ fun table -> call clientAgent (OuterMsg.TestTable table) @> :> Expr
            Some expr

        expected2 <- 
            let expr = 
                (<@ fun (table: obj[,]) -> 
                    call clientAgent (OuterMsg.TestTable (((Array2D table) :> ISurrogate).FromSurrogate client :?> Table))
                @>) :> Expr
            Some expr

        Registration.apiLambdas<OuterMsg> (Expr.Value clientAgent)
        |> Array.map (fun lambda -> ExcelFunctionRegistration lambda)

    [<RequireQualifiedAccess>]
    module CustomParamConversion =
        let register originType (param: ExcelParameterRegistration) =
            CustomParamConversion.register originType (fun conversion ->
                param.ArgumentAttribute.AllowReference <- conversion.IsRef
            )

    let paramConvertConfig =
        let funcParam f =
            Func<Type,ExcelParameterRegistration,LambdaExpression>(f)

        ParameterConversionConfiguration()
            .AddParameterConversion(funcParam(CustomParamConversion.register),null)
