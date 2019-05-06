namespace CellScript.Core
open Deedle
open System
open System.Reflection
open System.Linq.Expressions
open OfficeOpenXml

module internal Extensions =

    [<RequireQualifiedAccess>]
    module String =

        let ofCharList chars = chars |> List.toArray |> String

        let equalIgnoreCaseAndEdgeSpace (text1: string) (text2: string) =
            let trimedText1 = text1.Trim()
            let trimedText2 = text2.Trim()

            String.Equals(trimedText1,trimedText2,StringComparison.InvariantCultureIgnoreCase)

        let leftOf (pattern: string) (input: string) =
            let index = input.IndexOf(pattern)
            input.Substring(0,index)

        let rightOf (pattern: string) (input: string) =
            let index = input.IndexOf(pattern)
            input.Substring(index + 1)
            

    [<RequireQualifiedAccess>]
    module Type =

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
                    yield input.[i,*]
            }
            |> array2D

    [<RequireQualifiedAccess>]
    module Frame =
        let mapValuesString mapping frame =
            let mapping raw =
                mapping (raw.ToString())
            Frame.mapValues mapping frame

    [<RequireQualifiedAccess>]
    module ExcelRangeBase =
        let ofArray2D (p: ExcelPackage) (values: obj[,]) =
            let baseArray = Array2D.toSeqs values |> Seq.map Array.ofSeq
            let ws = p.Workbook.Worksheets.Add("Sheet1")
            ws.Cells.["A1"].LoadFromArrays(baseArray)

