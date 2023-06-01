namespace CellScript.Core
open System
open System.IO
open Shrimp.FSharp.Plus
open System.Diagnostics
open CellScript.Core.Extensions
open OfficeOpenXml

[<DebuggerDisplay("{ExcelCellAddress}")>]
type ComparableExcelCellAddress =
    { Row: int 
      Column: int }
with 
    static member OfExcelCellAddress(address: ExcelCellAddress) =
        { Row = address.Row 
          Column = address.Column }

    static member OfAddress(address: string) =
        ComparableExcelCellAddress.OfExcelCellAddress(ExcelCellAddress(address))

    member x.ExcelCellAddress =
        ExcelCellAddress(x.Row, x.Column)

    member x.Address = x.ExcelCellAddress.Address

    member x.Offset(rowOffset, columnOffset) =
        { Row    = x.Row + rowOffset 
          Column = x.Column + columnOffset }


[<DebuggerDisplay("{ExcelAddress}")>]
type ComparableExcelAddress =
    { StartRow: int 
      EndRow: int
      StartColumn: int 
      EndColumn: int 
      }
with 
    member x.Start: ComparableExcelCellAddress =
        { Row = x.StartRow 
          Column = x.StartColumn }

    member x.End: ComparableExcelCellAddress =
        { Row = x.EndRow
          Column = x.EndColumn }

    member x.AsCellAddress_Array2D() =
        [x.StartRow..x.EndRow]
        |> List.map(fun row ->
            [x.StartColumn..x.EndColumn]
            |> List.map(fun col ->
                { Row = row; Column = col }
            )
        )
        |> array2D
       
    member x.Offset(rowsOffset, columnsOffset) =
        { StartRow =  x.StartRow + rowsOffset
          StartColumn = x.StartColumn + columnsOffset
          EndRow  = x.EndRow + rowsOffset
          EndColumn = x.EndColumn + rowsOffset }

    member x.Offset(rowOffset, columnOffset, numberOfRows, numberOfColumns) =
        let rowStart = x.StartRow + rowOffset
        let columnStart = x.StartColumn + columnOffset
        let rowEnd = rowStart + numberOfRows
        let columnEnd = columnStart + numberOfColumns

        { StartRow = rowStart
          StartColumn = columnStart
          EndRow  = rowEnd
          EndColumn = columnEnd }

    member x.AsCellAddresses() =
        [
            for row = x.StartRow to x.EndRow do
                for column = x.StartColumn to x.EndColumn do
                    yield { Row = row; Column = column }
        ]

    member x.Rows = x.EndRow - x.StartRow + 1

    member x.Columns = x.EndColumn - x.StartColumn + 1

    static member OfAddress(excelAddress: ExcelAddress) =
        let startCell = excelAddress.Start

        let endCell = excelAddress.End
        {
            StartRow = startCell.Row
            EndRow = endCell.Row
            StartColumn = startCell.Column
            EndColumn = endCell.Column
        }

    static member OfAddress(address: string) =
        ComparableExcelAddress.OfAddress(ExcelAddress(address))

    static member OfRange(range: ExcelRangeBase) =
        let startCell = range.Start

        let endCell = range.End
        {
            StartRow = startCell.Row
            EndRow = endCell.Row
            StartColumn = startCell.Column
            EndColumn = endCell.Column
        }



    member x.ExcelAddress =
        ExcelAddress(x.StartRow, x.StartColumn, x.EndRow, x.EndColumn)
    
    member x.Address = x.ExcelAddress.Address

    member x.Contains(y: ComparableExcelAddress) = 
        match x.Start.Column, x.Start.Row, x.End.Column, x.End.Row with 
        | SmallerOrEqual y.Start.Column, SmallerOrEqual y.Start.Row, BiggerOrEqual y.End.Column, BiggerOrEqual y.End.Row ->
            true
        | _ -> false

    member x.IsIncludedIn(y: ComparableExcelAddress) = y.Contains(x)


type ComparableExcelCellAddress with 
    member x.RangeTo(y: ComparableExcelCellAddress) =
        x.Address + ":" + y.Address

    member x.Offset(rowOffset, columnOffset, numberOfRows, numberOfColumns) =
        let start = x.Offset(rowOffset, columnOffset)
        let endValue =
            x.Offset(
                rowOffset + numberOfRows,
                columnOffset + numberOfColumns
            )

        start.RangeTo(endValue)
        |> ComparableExcelAddress.OfAddress