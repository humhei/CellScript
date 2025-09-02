namespace CellScript.Core
open System
open System.IO
open Shrimp.FSharp.Plus
open System.Diagnostics
open OfficeOpenXml

[<DebuggerDisplay("{ExcelCellAddressText}")>]
type ComparableExcelCellAddress =
    { Row: int 
      Column: int }
with 
    static member OfExcelCellAddress(address: ExcelCellAddress) =
        { Row = address.Row 
          Column = address.Column }

    static member OfAddress(address: string) =
        ComparableExcelCellAddress.OfExcelCellAddress(ExcelCellAddress(address))

    static member OfRange(address: ExcelRangeBase) =
        match address.Columns, address.Rows with 
        | 1, 1 ->
            address.Start
            |> ComparableExcelCellAddress.OfExcelCellAddress

        | _ -> failwithf "Cannot create ComparableExcelCellAddress from %s" address.Address

    member x.ExcelCellAddress =
        ExcelCellAddress(x.Row, x.Column)

    member private x.ExcelCellAddressText = x.ExcelCellAddress.Address

    member x.Address = x.ExcelCellAddress.Address

    member x.Offset(rowOffset, columnOffset) =
        { Row    = x.Row + rowOffset 
          Column = x.Column + columnOffset }


[<DebuggerDisplay("{ExcelAddressText}")>]
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
    
    member private x.ExcelAddressText = x.ExcelAddress.Address

    member x.Address = x.ExcelAddress.Address

    member x.Contains(y: ComparableExcelAddress) = 
        match x.StartColumn, x.StartRow, x.EndColumn, x.EndRow with 
        | SmallerOrEqual y.StartColumn, SmallerOrEqual y.StartRow, BiggerOrEqual y.EndColumn, BiggerOrEqual y.EndRow ->
            true
        | _ -> false

    member x.Contains(y: ComparableExcelCellAddress) = 
        y.Column.IsBetween(x.StartColumn, x.EndColumn) 
            &&
                y.Row.IsBetween(x.StartRow, x.EndRow)
                
    member x.IntersectTo(y: ComparableExcelAddress) =
        let rowsIntersect =
            let x = 
                [x.StartRow..x.EndRow]
                |> Set.ofList

            let y = 
                [y.StartRow..y.EndRow] 
                |> Set.ofList

            Set.intersect x y 

        match rowsIntersect.IsEmpty with 
        | true -> None
        | false ->
            let columnIntersect =
                let x = 
                    [x.StartColumn..x.EndColumn]
                    |> Set.ofList

                let y = 
                    [y.StartColumn..y.EndColumn] 
                    |> Set.ofList

                Set.intersect x y 

            match columnIntersect.IsEmpty with 
            | true -> None
            | false ->
                {
                    ComparableExcelAddress.StartRow = rowsIntersect.MinimumElement
                    EndRow = rowsIntersect.MaximumElement
                    StartColumn = columnIntersect.MinimumElement
                    EndColumn = columnIntersect.MaximumElement
                }
                |> Some



    member x.IsIncludedIn(y: ComparableExcelAddress) = y.Contains(x)

    static member Concat(addrs: al1List<ComparableExcelAddress>) =
        let startRow = 
            addrs.AsList
            |> List.map(fun m -> m.StartRow)
            |> List.min

        let startColumn = 
            addrs.AsList
            |> List.map(fun m -> m.StartColumn)
            |> List.min

        let endRow = 
            addrs.AsList
            |> List.map(fun m -> m.EndRow)
            |> List.max

        let endColumn = 
            addrs.AsList
            |> List.map(fun m -> m.EndColumn)
            |> List.max


        { StartRow    = startRow 
          EndRow      = endRow
          StartColumn = startColumn 
          EndColumn   = endColumn }

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