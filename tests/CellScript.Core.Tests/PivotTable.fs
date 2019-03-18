namespace CellScript.Core.Tests
open CellScript.Core.Registration
open CellScript.Core.Types
open CellScript.Core
open ExcelProcess.MatrixParsers
open FParsec
open OfficeOpenXml
open CellScript.Core.Extensions
open Deedle
open CellScript.Core.Tests.Extensions

module PivotTable =

    type PivotHeader =
        { HeaderKey: string * string
          KeeperHeaders: string list
          TransferAbleHeaders: string list }

    [<RequireQualifiedAccess>]
    module PivotHeader =
        let ofHeaders (headers: seq<string>) =
            let p =
                let biggerThan3 i = i > 3
                [ mxManyWith biggerThan3 (!^^ pfloat Shoes.isEurFloat)
                  mxManyWith biggerThan3 (!^^ pfloat Shoes.isUK)
                  mxManyWith biggerThan3 (!^^ pfloat Shoes.isUS) ]

            let sizes,sizeHeaders,others = Seq.partitionByMatrixParsers p headers

            let sectionText = Shoes.sectionText sizes

            { HeaderKey = sectionText, "Number"
              KeeperHeaders = others
              TransferAbleHeaders = sizeHeaders }

    [<CustomParamConversion>]
    type PlainPivotTable =
        { Frame: ExcelFrame<int,string>
          Header: PivotHeader }
    with
        static member Convert() =
            Table.Convert()
            |> CustomParamConversion.mapM (fun (table) ->
                { Frame = Table.value table
                  Header = PivotHeader.ofHeaders table.Headers }
            )

    [<RequireQualifiedAccess>]
    module PlainPivotTable =
        let transfer (pvt: PlainPivotTable) =
            let frame = pvt.Frame.AsFrame
            let transferAbleHeaders = pvt.Header.TransferAbleHeaders
            let keeperHeaders = pvt.Header.KeeperHeaders
            let transferHeaderKey, valueHeaderKey = pvt.Header.HeaderKey

            frame
            |> Frame.mapRowValues (fun row ->
                let baseRow =
                    row
                    |> Series.filter (fun k v ->
                        List.contains k keeperHeaders
                    )
                let resizedRows =
                    transferAbleHeaders
                    |> List.map (fun header ->
                        [transferHeaderKey => box header; valueHeaderKey => row.[header]]
                        |> Series.ofObservations
                        |> Series.merge baseRow
                    )
                resizedRows
            )
            |> Series.values
            |> Seq.concat
            |> Frame.ofRowsOrdinal
            |> Frame.indexRowsOrdinally
            |> ExcelFrame
            |> Table


    // type TransferAbleHeader =
    //     { HeaderKey: string * string
    //       Headers: string list }

    // [<RequireQualifiedAccess>]
    // type TwoRowHeader =
    //     | Normal of string
    //     | TransferAble of TransferAbleHeader

    // [<CustomParamConversion>]
    // type TwoRowHeaderPivotTable =
    //     { Frame: ExcelFrame<int, TwoRowHeader> }
    // with
    //     static member Convert() =
    //         fun (inputs: obj[,]) ->
    //             let rebased =
    //                 inputs
    //                 |> Array2D.rebase

    //             let headers =

    //                 let firstRow = rebased.[0,*]
    //                 let secondRow = rebased.[1,*]
    //                 (firstRow, secondRow)
    //                 ||> Array.map2 (fun firstValue sencondValue ->
    //                     match (Some firstValue,Some sencondValue) with
    //                     | CellValue.HasSense _, CellValue.NoSense ->
    //                         TwoRowHeader.Normal (string firstValue)
    //                     | CellValue.HasSense _, CellValue.HasSense _ ->


    //                 )
