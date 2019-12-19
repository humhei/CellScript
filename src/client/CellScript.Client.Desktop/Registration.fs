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
open OfficeOpenXml
open System.Threading
open System
open Akka.Event

[<AutoOpen>]
module internal Extensions =
    let value2ToArray2D (value2: obj) =
        match value2 with 
        | null ->
            (array2D [[box ExcelErrorEnum.Null]])
        | _ ->

            let tp = value2.GetType()
            if tp.IsArray then 
                let values: obj [,] = unbox value2
                values
            else 
                array2D[[value2]]

            |> Array2D.map (fun value ->
                match value with 
                | :? ExcelEmpty -> box ExcelErrorEnum.Null
                | :? ExcelError as error -> 
                    let byte = int16 error
                    let errorEnum: ExcelErrorEnum = LanguagePrimitives.EnumOfValue byte
                    box errorEnum
                | _ -> value
            )

[<RequireQualifiedAccess>]
module private ExcelReference =

    let workbookPathAndSheetName (xlRef: ExcelReference) =
        let retrievedSheetName = XlCall.Excel(XlCall.xlSheetNm, xlRef) :?> string

        let workbookPath =
            let app = ExcelDnaUtil.Application :?> Application
            app.ActiveWorkbook.FullName
            //let dir = 
            //    XlCall.Excel(XlCall.xlfGetDocument,2,retrievedSheetName) :?> string

            //let name = XlCall.Excel(XlCall.xlfGetDocument,88,retrievedSheetName) :?> string
            //Path.Combine(dir,name)

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
module ExcelRangeContactInfo =
    let ofExcelReference (xlRef: ExcelReference) =
        let workbookPath, sheetName = ExcelReference.workbookPathAndSheetName xlRef
        { ColumnFirst = xlRef.ColumnFirst + 1 
          RowFirst = xlRef.RowFirst + 1
          ColumnLast = xlRef.ColumnLast + 1
          RowLast = xlRef.RowLast + 1
          Content = ExcelReference.toArray2D xlRef
          WorkbookPath = workbookPath
          SheetName = sheetName }

    let ofRange workbookPath sheetName (range: Range) =
        let value2 = 
            try range.Value2
            with _ -> box ""

        { ColumnFirst = range.Column
          RowFirst = range.Row
          ColumnLast = range.Columns.Count - 1 + range.Column
          RowLast = range.Rows.Count - 1 + range.Row
          Content = value2ToArray2D value2
          WorkbookPath = workbookPath
          SheetName = sheetName }

[<RequireQualifiedAccess>]
module NetCoreClient =
    let create<'Msg> seedPort msgMapping = 

        let getCommandCaller() =
            let app = ExcelDnaUtil.Application :?> Application
            let activeWorkbookPath = app.ActiveWorkbook.FullName
                        
            let activeSheetName = 
                let sheet = app.ActiveSheet :?> Worksheet
                sheet.Name

            let selection = app.Selection :?> Range
            ExcelRangeContactInfo.ofRange activeWorkbookPath activeSheetName selection
            |> CommandCaller


        let client: NetCoreClient<'Msg> = 
            let actionsInCorrentWorkbook (logger: ILoggingAdapter) (xlRefs: ExcelRangeContactInfo list) f =
                try 
                    let app = ExcelDnaUtil.Application :?> Application

                    let activeWorkbookPath = app.ActiveWorkbook.FullName
                        
                    let activeSheetName = 
                        let sheet = app.ActiveSheet :?> Worksheet
                        sheet.Name
                    xlRefs 
                    |> List.iter (fun xlRef ->
                        if activeWorkbookPath = xlRef.WorkbookPath && activeSheetName = xlRef.SheetName then
                            f app xlRef
                        else failwithf "except: %s_%s; actual: %s_%s" xlRef.WorkbookPath xlRef.SheetName activeWorkbookPath activeSheetName
                    )
                with ex ->
                    logger.Error (sprintf "%A" ex)

            let actionInCorrentWorkbook logger xlRef f =
                actionsInCorrentWorkbook logger [xlRef] f

            NetCoreClient.create seedPort getCommandCaller (fun logger msg ->
                match msg with 
                | EndPointClientMsg.UpdateCellValues xlRefs ->
                    actionsInCorrentWorkbook logger xlRefs (fun app xlRef ->
                    for i = xlRef.RowFirst to xlRef.RowLast do
                        for j = xlRef.ColumnFirst to xlRef.ColumnLast do
                            app.Cells.[i,j] <- xlRef.Content.[i - xlRef.RowFirst,j - xlRef.ColumnFirst]
                    )
                    Ignore

                | EndPointClientMsg.SetRowHeight (xlRef, height) ->
                    actionInCorrentWorkbook logger xlRef (fun app xlRef ->
                        let cell = app.Range(xlRef.CellAddress)
                        cell.EntireRow.RowHeight <- height
                    )
                    Ignore

            ) msgMapping

        client

[<RequireQualifiedAccess>]
module Registration =

    let private toExcelRangeContactInfo (xlRef: ExcelReference): ExcelRangeContactInfo =
        ExcelRangeContactInfo.ofExcelReference xlRef

    let private createFunctionCaller() =
        let app = ExcelDnaUtil.Application :?> Application
        let activeWorkbookPath = app.ActiveWorkbook.FullName
        let activeSheetName = 
            let sheet = app.ActiveSheet :?> Worksheet
            sheet.Name

        { WorkbookPath = activeWorkbookPath 
          SheetName = activeSheetName }
        |> FunctionSheetReference


    let private parameterConfig =
        let paramConfig = 
            ParameterConversionConfiguration()
                .AddParameterConversion(fun (input: obj) ->
                    toExcelRangeContactInfo (input :?> ExcelReference)
                )
                .AddParameterConversion(fun (input: obj) ->
                    value2ToArray2D (input)
                )
                .AddParameterConversion (fun (input: obj) ->
                    createFunctionCaller()
                )


        (paramConfig, NetCoreClient.normalTypesparamConversions) 
        ||> List.fold (fun paramConfig conversion ->
            let targetType = conversion.ReturnType
            let func = Func<_,_,_>(fun (_: Type) _ -> conversion)
            paramConfig.AddParameterConversion(func, targetType)
        )


    let excelFunctions (apiLambdas: ApiLambda []) = 
        apiLambdas
        |> Array.choose ApiLambda.asFunction  
        |> Array.map (fun lambda -> 
            let excelFunction = ExcelFunctionRegistration lambda
            excelFunction.FunctionLambda.Parameters
            |> Seq.iteri (fun i parameter ->
                if parameter.Type = typeof<ExcelRangeContactInfo> then
                    excelFunction.ParameterRegistrations.[i].ArgumentAttribute.AllowReference <- true
            )
            excelFunction.FunctionAttribute.IsMacroType <- 
                excelFunction.ParameterRegistrations |> Seq.exists (fun paramReg -> paramReg.ArgumentAttribute.AllowReference)

            excelFunction.FunctionAttribute.IsMacroType <- 
                excelFunction.FunctionLambda.Parameters |> Seq.exists (fun parameter -> 
                    parameter.Type = typeof<FunctionSheetReference>
                )
            
            
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
