namespace CellScript.Client
open ExcelDna.Integration
open ExcelDna.Registration
open System
open System.Reflection
open CellScript.Core.Registration
open System.Linq.Expressions
open Fable.Remoting.DotnetClient


[<RequireQualifiedAccess>]
module Registration =
    open CellScript.Core.Extensions
    open CellScript.Shared

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


    [<ExcelFunction>]
    let hellowo2 string =
        string + "666"


    let excelFunctions() =
        let proxy = Proxy.create<UDF.ICellScriptUDFApi> UDF.port.UrlBuilder
        let propNames = typeof<UDF.ICellScriptUDFApi>.GetProperties() |> Array.map (fun prop -> prop.Name)
        let lambdas = UDF.apiLambdas proxy.call
        //let functions = 
        //    lambdas 
        //    |> List.mapi (fun i lambda ->
        //        let excelFunc = ExcelFunctionRegistration(lambda)
        //        excelFunc.FunctionAttribute.Name <- propNames.[i]
        //        excelFunc
        //    )
        []
        //functions

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