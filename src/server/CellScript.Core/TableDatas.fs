namespace CellScript.Core
open Shrimp.FSharp.Plus

[<AutoOpen>]
module _TableDatas =
    type ExcelPackageWithXlsxFile with 
        member excelPackage.ReadTable_R(?sheetGettingOptions, ?rangeGettingOptions) =
            let sheet = 
                excelPackage.ExcelPackage.TryGetValidWorksheet(defaultArg sheetGettingOptions SheetGettingOptions.DefaultValue)

            match sheet with 
            | Result.Error error -> Result.Error error
            | Result.Ok sheet ->
                let rangeGettingOptions = defaultArg rangeGettingOptions RangeGettingOptions.UserRange
                let rangeGettingOptions = rangeGettingOptions.SetIncludeHided(true)

                match sheet.TryReadDatas(rangeGettingOptions) with 
                | Result.Error error -> Result.Error error.Message
                | Result.Ok datas ->
                    let tb = 
                        datas.Content
                        |> Table.OfArray2D

                    Result.Ok tb

        member excelPackage.ReadTableDatas_R(f, ?sheetGettingOptions, ?rangeGettingOptions) =
            match excelPackage.ReadTable_R(?sheetGettingOptions = sheetGettingOptions, ?rangeGettingOptions = rangeGettingOptions) with 
            | Result.Error error -> Result.Error error
            | Result.Ok tb ->
                let rows = 
                    tb.Rows.Values
                    |> List.ofSeq
                    |> List.filter(fun m -> m.ValueCount <> 0)

                match AtLeastOneList.TryCreate rows with 
                | None -> Result.Error "EmptyTable"
                | Some rows ->
                    rows
                    |> AtLeastOneList.map(fun row ->
                        f row
                    )
                    |> Result.Ok


        member excelPackage.ReadTableDatas(f, ?sheetGettingOptions, ?rangeGettingOptions) =
            excelPackage.ReadTableDatas_R(f = f, ?sheetGettingOptions = sheetGettingOptions, ?rangeGettingOptions = rangeGettingOptions)
            |> Result.getOrFail

        member excelPackage.ReadTableDatas_R(sheetName, tableName, f) =
            excelPackage.ReadTableDatas_R(
                f,
                sheetGettingOptions = SheetGettingOptions.SheetName(StringIC sheetName), 
                rangeGettingOptions = RangeGettingOptions.TableName tableName)
            
        member excelPackage.ReadTableDatas(sheetName, tableName, f) =
            excelPackage.ReadTableDatas_R(sheetName = sheetName, tableName = tableName, f = f)
            |> Result.getOrFail
