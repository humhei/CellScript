namespace CellScript.Client
open ExcelDna.Registration
open CellScript.Core
open ExcelDna.Integration
open System.IO
open CellScript.Core.Types
open ExcelDna.Integration
open System
open System.Linq.Expressions

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
          Content = xlRef.GetValue() |> unbox
          RowLast = xlRef.RowLast 
          WorkbookPath = workbookPath 
          SheetName = sheetName }

    let private parameterConfig =
        ParameterConversionConfiguration()
            .AddParameterConversion(fun (input: obj) ->
                toSerialbleExcelReference input
            )

    let excelFunctions (client: Client<'Msg>) = 
        let remote = Remote(RemoteKind.Client client)
        Client.apiLambdas remote.RemoteActor.Value client
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

