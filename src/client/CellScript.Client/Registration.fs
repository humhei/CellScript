namespace CellScript.Client
open ExcelDna.Integration
open ExcelDna.Registration
open System
open System.Reflection
open CellScript.Core.Registration
open System.Linq.Expressions
open CellScript.Core.Extensions

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






    let excelFunctions() =
        let assembly = Assembly.GetExecutingAssembly()
        Assembly.excelFunctions assembly

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