namespace CellScript.Client
open ExcelDna.Integration
open ExcelDna.Registration
open System
open System.Reflection
open CellScript.Core.Registration
open System.Linq.Expressions
open Elmish
open Elmish.Bridge
open CellScript.Core
open System.Timers
open Microsoft.FSharp.Reflection
open FSharp.Quotations.Evaluator

[<RequireQualifiedAccess>]
module Registration =

    [<RequireQualifiedAccess>]
    module Assembly =
        let getAllExportMethods (assembly: Assembly) =
            assembly.ExportedTypes
            |> Seq.collect (fun exportType ->
                exportType.GetMethods()
            )

        let getAllExportMethodsWithSpecificAttribute (attributeType: Type) (assembly: Assembly) =
            getAllExportMethods assembly
            |> Seq.filter (fun methodInfo ->
                methodInfo.CustomAttributes |> Seq.exists (fun attribute ->
                    attribute.AttributeType = attributeType
                )
            )

        let excelFunctions assembly =
            getAllExportMethodsWithSpecificAttribute typeof<ExcelFunctionAttribute> assembly
            |> Seq.map ExcelFunctionRegistration
            |> List.ofSeq




    let init() =
        OuterMsg.None, Cmd.none

    let update msg model =
        Bridge.Send msg
        model, Cmd.none


    let view model (dispatch: Dispatch<OuterMsg>) =
        let s = typeof<OuterMsg>
        let ucis = FSharpType.GetUnionCases s
        let s = ucis.[0].DeclaringType.GetMethods().[0]
        let p = s.GetParameters()
        let k = p.[0].Name
        let expr = 
            <@ 
                fun (table) -> Table.op_ErasedCast table

            @>
        let kp = expr.Compile()
        []

    let excelFunctions() =
        Program.mkProgram init update view
        |> Program.withBridge Routed.wsUrl
        |> Program.run
        //functions
        []

    [<RequireQualifiedAccess>]
    module CustomParamConversion =

        let register originType (param: ExcelParameterRegistration) =
            CustomParamConversion.register originType (fun conversion ->
                param.ArgumentAttribute.AllowReference <- conversion.IsRef
            )

    [<RequireQualifiedAccess>]
    module ICustomReturn =
        let register (tp: Type) (_: ExcelReturnRegistration) =
            ICustomReturn.register tp

    let paramConvertConfig =
        let funcParam f =
            Func<Type,ExcelParameterRegistration,LambdaExpression>(f)

        let funcReturn f =
            Func<Type,ExcelReturnRegistration,LambdaExpression>(f)

        ParameterConversionConfiguration()
            .AddParameterConversion(funcParam(CustomParamConversion.register),null)
            .AddReturnConversion(funcReturn(ICustomReturn.register),null)