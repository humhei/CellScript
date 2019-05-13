namespace CellScript.Client.Desktop
open ExcelDna.Registration
open CellScript.Core
open ExcelDna.Integration
open System.IO
open CellScript.Core.Types
open CellScript.Core.Cluster
open Microsoft.Office.Interop.Excel
open Akkling
open CellScript.Client.Core

[<AutoOpen>]
module internal Extensions =
    let value2ToArray2D (value2: obj) =
        let tp = value2.GetType()
        if tp.IsArray then unbox value2
        else array2D[[value2]]

[<RequireQualifiedAccess>]
module private ExcelReference =

    let workbookPathAndSheetName (xlRef: ExcelReference) =
        let retrievedSheetName = XlCall.Excel(XlCall.xlSheetNm, xlRef) :?> string

        let workbookPath =
            let dir = XlCall.Excel(XlCall.xlfGetDocument,2,retrievedSheetName) :?> string
            let name = XlCall.Excel(XlCall.xlfGetDocument,88,retrievedSheetName) :?> string
            Path.Combine(dir,name)

        let sheetName =
            let rightOf (pattern: string) (input: string) =
                let index = input.IndexOf(pattern)
                input.Substring(index + 1)

            rightOf "]" retrievedSheetName

        workbookPath,sheetName

    let toArray2D (xlRef: ExcelReference) =
        let value = xlRef.GetValue()
        value2ToArray2D value

[<RequireQualifiedAccess>]
module private SerializableExcelReference =
    let ofExcelReference (xlRef: ExcelReference) =
        let workbookPath, sheetName = ExcelReference.workbookPathAndSheetName xlRef
        { ColumnFirst = xlRef.ColumnFirst 
          RowFirst = xlRef.RowFirst 
          ColumnLast = xlRef.ColumnLast
          RowLast = xlRef.RowLast
          Content = ExcelReference.toArray2D xlRef
          WorkbookPath = workbookPath
          SheetName = sheetName }

    let ofRange workbookPath sheetName (range: Range) =

        { ColumnFirst = range.Column
          RowFirst = range.Row
          ColumnLast = range.Columns.Count - 1 + range.Column
          RowLast = range.Rows.Count - 1 + range.Row
          Content = value2ToArray2D range.Value2
          WorkbookPath = workbookPath
          SheetName = sheetName }

[<RequireQualifiedAccess>]
module NetCoreClient =
    let create<'Msg> seedPort  = 

        let getCaller() =
            let app = ExcelDnaUtil.Application :?> Application
            let activeWorkbookPath = app.ActiveWorkbook.Path
                        
            let activeSheetName = 
                let sheet = app.ActiveSheet :?> Worksheet
                sheet.Name

            let selection = app.Selection :?> Range
            SerializableExcelReference.ofRange activeWorkbookPath activeSheetName selection
            |> CommandCaller

        let client: NetCoreClient<'Msg> = 
            NetCoreClient.create seedPort getCaller (fun logger msg ->
                match msg with 
                | EndPointClientMsg.UpdateCellValues xlRefs ->
                    try 
                        let app = ExcelDnaUtil.Application :?> Application

                        let activeWorkbookPath = app.ActiveWorkbook.Path
                        
                        let activeSheetName = 
                            let sheet = app.ActiveSheet :?> Worksheet
                            sheet.Name

                        xlRefs 
                        |> List.iter (fun xlRef ->
                            if activeWorkbookPath = xlRef.WorkbookPath && activeSheetName = xlRef.SheetName then
                                for i = xlRef.RowFirst to xlRef.RowLast do
                                    for j = xlRef.ColumnFirst to xlRef.ColumnLast do
                                        app.Cells.[i,j] <- xlRef.Content.[i - xlRef.RowFirst,j - xlRef.ColumnFirst]

                            else failwithf "except: %s_%s; actual: %s_%s" xlRef.WorkbookPath xlRef.SheetName activeWorkbookPath activeSheetName
                        )
                        Ignore

                    with ex ->
                        logger.Error (sprintf "%A" ex)
                        Ignore

            )

        client

[<RequireQualifiedAccess>]
module Registration =

    let private toSerialbleExcelReference (xlRef: ExcelReference): SerializableExcelReference =
        SerializableExcelReference.ofExcelReference xlRef

    let private parameterConfig =
        ParameterConversionConfiguration()
            .AddParameterConversion(fun (input: obj) ->
                toSerialbleExcelReference (input :?> ExcelReference)
            )

    let excelFunctions (apiLambdas: ApiLambda []) = 
        apiLambdas
        |> Array.choose ApiLambda.asFunction  
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


    let excelCommands (apiLambdas: ApiLambda []) = 
        apiLambdas
        |> Array.choose ApiLambda.asCommand  
        |> Array.map (fun (commandSetting, lambda) -> 
            let excelCommand = ExcelCommandRegistration lambda
            excelCommand.CommandAttribute.MenuText <- lambda.Name
            match commandSetting with 
            | Some comandSetting ->
                match comandSetting.Shortcut with 
                | Some shortcut -> excelCommand.CommandAttribute.ShortCut <- shortcut
                | None -> ()
            | None -> ()
            
            excelCommand

        )
