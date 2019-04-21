namespace CellScript.Client
open ExcelDna.Registration
open CellScript.Core
open ExcelDna.Integration
open System.IO
open CellScript.Core.Types
open ExcelDna.Integration
open System
open System.Linq.Expressions
open Remote

[<RequireQualifiedAccess>]
module Registration =
    let private toSerialbleExcelReference (input: obj): SerializableExcelReference =
        let xlRef = input :?> ExcelReference

        let retrievedSheetName = XlCall.Excel(XlCall.xlSheetNm,xlRef) :?> string

        let workbookPath =
            let dir = XlCall.Excel(XlCall.xlfGetDocument,2,retrievedSheetName) :?> string
            let name = XlCall.Excel(XlCall.xlfGetDocument,88,retrievedSheetName) :?> string
            Path.Combine(dir,name)

        let sheetName =
            let rightOf (pattern: string) (input: string) =
                let index = input.IndexOf(pattern)
                input.Substring(index + 1)

            rightOf "]" retrievedSheetName

        { ColumnFirst = xlRef.ColumnFirst
          ColumnLast = xlRef.ColumnLast
          RowFirst = xlRef.RowFirst
          Content = 
            let value = xlRef.GetValue()
            let tp = value.GetType()
            if tp.IsArray then unbox value
            else array2D[[value]]
          RowLast = xlRef.RowLast 
          WorkbookPath = workbookPath 
          SheetName = sheetName }

    let private parameterConfig =
        ParameterConversionConfiguration()
            .AddParameterConversion(fun (input: obj) ->
                toSerialbleExcelReference input
            )

    let excelFunctions<'Msg>() = 
        let client: CellScriptClient<'Msg> = CellScriptClient.create()
        CellScriptClient.apiLambdas client
        |> Array.map (fun lambda -> 
            let excelFunction = ExcelFunctionRegistration lambda
            excelFunction.FunctionLambda.Parameters
            |> Seq.iteri (fun i parameter ->
                if parameter.Type = typeof<SerializableExcelReference> then
                    excelFunction.ParameterRegistrations.[i].ArgumentAttribute.AllowReference <- true
            )
            excelFunction.FunctionAttribute.IsMacroType <- 
                excelFunction.ParameterRegistrations |> Seq.exists (fun paramReg -> paramReg.ArgumentAttribute.AllowReference)
            excelFunction
        )
        |> fun fns ->
            ParameterConversionRegistration.ProcessParameterConversions (fns, parameterConfig)


    let evalFunction()=
        excelFunctions<FcsMsg>()