﻿

namespace 眼镜复杂外箱贴
#nowarn "0104"
module Module_眼镜复杂外箱贴 =
    open System.IO
    open Deedle
    open OfficeOpenXml

    open CellScript.Core
    open Shrimp.FSharp.Plus

    let parse(xlsxFile: XlsxFile) =
        let excelRangeContactInfo = 
            ExcelRangeContactInfo.readFromFile 
                RangeGettingOptions.UserRange 
                (SheetGettingOptions.SheetNameOrSheetIndex ("Sheet1", 0)) xlsxFile
        
        let table = 

            Table.OfArray2D excelRangeContactInfo.Content
            |> Table.removeEmptyRows
            |> Table.mapFrame (fun frame -> 
                let cartonNumber = 
                    frame.ColumnKeys
                    |> List.ofSeq
                    |> List.find(fun m -> m = stringIC "件数")
                Frame.filterRowValues (fun row ->
                    (row.FillMissing "").[cartonNumber].ToString() <> ""
                ) frame
            )
            |> Table.fillEmptyUp
            |> Table.splitRowToMany ["箱号_Left"; "箱号_Right"] (fun key row ->
                let cartonIdLeft, cartonIdRight =
                    let cartonIdRangeText = row.[stringIC "箱号"].ToString()
                    let texts = cartonIdRangeText.Split([| "--"; "-" |], System.StringSplitOptions.None)
                    System.Int32.Parse(texts.[0]), System.Int32.Parse(texts.[1])

                Seq.replicate (cartonIdRight - cartonIdLeft + 1) row
                |> Seq.mapi (fun i row ->
                    Seq.append row.ValuesAll [ box (cartonIdLeft + i); box cartonIdRight ]
                )
            )

        table

    let storages() =
        let tb = parse(XlsxFile @"D:\Users\Jia\Documents\Tencent Files\857701188\FileRecv\EU80-21014  &21015 装箱明细2.xlsx")
        let m = tb
        tb.SaveToXlsx(@"C:\Users\Jia\Desktop\新建 Microsoft Excel 工作表.xlsx")
        ()
