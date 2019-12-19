namespace CellScript.Core
open Deedle
open System
open OfficeOpenXml
open Deedle.Internal
open System.IO
open System.Collections.Generic
open System.Runtime.CompilerServices

module  Extensions =

    [<RequireQualifiedAccess>]
    module internal String =

        let ofCharList chars = chars |> List.toArray |> String

        let equalIgnoreCase (text1: string) (text2: string) =
            String.Equals(text1,text2,StringComparison.InvariantCultureIgnoreCase)

    type internal String with 
        member x.LeftOf(pattern: string) =
            let index = x.IndexOf(pattern)
            x.Substring(0,index)

        member x.RightOf(pattern: string) =
            let index = x.IndexOf(pattern)
            x.Substring(index + 1)

            

    [<RequireQualifiedAccess>]
    module internal Type =

        /// no inheritance
        let tryGetAttribute<'Attribute> (tp: Type) =
            tp.GetCustomAttributes(false)
            |> Seq.tryFind (fun attr ->
                let t1 =  attr.GetType()
                let t2 =typeof<'Attribute>
                t1 = t2
            )


    [<RequireQualifiedAccess>]
    module Array2D =

        let toSeqs (input: 'a[,]) =
            let l1 = input.GetLowerBound(0)
            let u1 = input.GetUpperBound(0)
            seq {
                for i = l1 to u1 do
                    yield input.[i,*] :> seq<'a>
            }

        let transpose (input: 'a[,]) =
            let l1 = input.GetLowerBound(1)
            let u1 = input.GetUpperBound(1)
            seq {
                for i = l1 to u1 do
                    yield input.[*,i]
            }
            |> array2D

    

    [<RequireQualifiedAccess>]
    module Frame =
        let mapValuesString mapping frame =
            let mapping raw = mapping (raw.ToString())
            Frame.mapValues mapping frame
        

    type ExcelRangeBase with
        member x.LoadFromArray2D(array2D: obj [,]) =
            let baseArray = Array2D.toSeqs array2D |> Seq.map Array.ofSeq
            x.LoadFromArrays(baseArray)


    type RangeGettingArg =
        | RangeIndexer of string
        | UserRange

    type ExcelWorksheet with
        member sheet.GetRange arg = 
            match arg with
            | RangeIndexer indexer ->
                sheet.Cells.[indexer]

            | UserRange ->
                let indexer = sheet.Dimension.Start.Address + ":" + sheet.Dimension.End.Address
                sheet.Cells.[indexer]

    type SheetGettingArg =
        | SheetName of string
        | SheetIndex of int
        | SheetNameOrSheetIndex of sheetName: string * index: int

    type ExcelPackage with
        member excelPackage.GetWorkSheet(arg) =
            match arg with 
            | SheetGettingArg.SheetName sheetName -> excelPackage.Workbook.Worksheets.[sheetName]
            | SheetGettingArg.SheetIndex index -> excelPackage.Workbook.Worksheets.[index]
            | SheetGettingArg.SheetNameOrSheetIndex (sheetName, index) ->
                match Seq.tryFind (fun (worksheet: ExcelWorksheet) -> worksheet.Name = sheetName) excelPackage.Workbook.Worksheets with 
                | Some worksheet -> worksheet
                | None -> excelPackage.Workbook.Worksheets.[index]


