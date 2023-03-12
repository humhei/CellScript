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


type AddressedArray =
    { Address: ComparableExcelCellAddress 
      Array: ConvertibleUnion [,]
      SpecificName: string option }
with 
    //member x.Name = 
    //    let name = 
    //        match x.SpecificName with 
    //        | None -> x.Array.[0,0].Text
    //        | Some name -> name

    //    name.ForceEndingWith("表")


    static member SetName(name) x =
        { x with SpecificName = Some name }

    /// No Table
    static member ofValue address (value: string) =
        let array = 
            [
                [
                    ConvertibleUnion.Convert value
                ]
            ]
            |> array2D

        { Address = ComparableExcelCellAddress.OfAddress address 
          Array = array
          SpecificName = None }

    /// As Table
    static member ofObservations address (observations: list<_ * ConvertibleUnion>) =
        let array = 
            observations
            |> List.map(fun (a, b) ->
                [ConvertibleUnion.Convert a; b]
            )
            |> array2D

        { Address = ComparableExcelCellAddress.OfAddress address 
          Array = array
          SpecificName = Some (array.[0, 0].Text.ForceEndingWith("表")) }

    /// As Table
    static member ofColumn address (value: list<ConvertibleUnion>) =
        let array = 
            value
            |> List.map(fun (a) ->
                [a]
            )
            |> array2D

        { Address = ComparableExcelCellAddress.OfAddress address 
          Array = array
          SpecificName = Some (array.[0, 0].Text.ForceEndingWith("表")) }



    /// As Table
    static member ofTitledArray address  (title: string, array: ConvertibleUnion list) =
        
        { Address = ComparableExcelCellAddress.OfAddress address 
          Array = 
            ConvertibleUnion.Convert title :: array 
            |> List.map List.singleton
            |> array2D
          SpecificName = Some (title.ForceEndingWith("表"))
        }

    /// As Table
    static member ofTitledArray2D address (title: string list, array: ConvertibleUnion [,]) =
        let __ensureNotTitleDuplicated =
            title
            |> List.filter(fun m -> m.Trim() <> "")
            |> List.map StringIC
            |> AtLeastOneSet.create true
        let array = 
            let colNum = Array2D.length2 array
            let title =
                [0..colNum-1]
                |> List.map(fun i ->
                    match List.tryItem i title with 
                    | Some title -> ConvertibleUnion.Convert title
                    | _ -> ConvertibleUnion.Convert ""
                )
                 
            let array = Array2D.toLists array

            title :: array
            |> array2D
                    


        { Address = ComparableExcelCellAddress.OfAddress address 
          Array = array
          SpecificName = Some (title.[0].ForceEndingWith("表")) }


    /// As Table
    static member ofTable address (table: Table) =
        let content = 
            table.ToArray2D()
            |> Array2D.map ConvertibleUnion.Convert

        { Address = ComparableExcelCellAddress.OfAddress address 
          Array = content
          SpecificName = Some (content.[0,0].Text.ForceEndingWith("表")) }

type AddressedArrays = AddressedArrays of Map<ComparableExcelCellAddress, AddressedArray>
with 
    member x.AsList =
        let (AddressedArrays v) = x
        v
        |> List.ofSeq
        |> List.map(fun m -> m.Value)

    member x.Value =
        let (AddressedArrays v) = x
        v

    member x.Item(address: string) =
        x.Value.[ComparableExcelCellAddress.OfAddress address].Array

    member x.Item(address: ComparableExcelCellAddress) =
        x.Value.[address].Array

    static member OfList(arrays: AddressedArray list) =
        arrays
        |> List.map(fun m -> m.Address, m)
        |> Map.ofList
        |> AddressedArrays

    


[<AutoOpenAttribute>]
module _Address_Extensions =

    type ExcelWorksheet with 

        member x.ReadToAddressedArray(address: ComparableExcelCellAddress) =
            let array = 
                let rec getBottom (address: ComparableExcelCellAddress) =
                    let nextAddress = { address with Row = address.Row + 1 }
                    let nextText = x.Cells.[nextAddress.Address].Text
                    match nextText with 
                    | null -> address
                    | nextText ->
                        match nextText.Trim() with 
                        | "" -> address
                        | _ -> getBottom nextAddress

                let rec getRight (address: ComparableExcelCellAddress) =
                    let nextAddress = { address with Column = address.Column + 1 }
                    let nextText = x.Cells.[nextAddress.Address].Text
                    match nextText with 
                    | null -> address
                    | nextText ->
                        match nextText.Trim() with 
                        | "" -> address
                        | _ -> getRight nextAddress


                let leftBottom, rightBottom =
                    let firtRowRight = getRight address
                    let leftBottom = getBottom address
                    let rightBottom = getBottom firtRowRight
                    let bottom = max leftBottom.Row firtRowRight.Row


                    { leftBottom with Row = bottom }, {rightBottom with Row = bottom }


                //let leftBottom = getBottom address

                //let rightBottom =  
                //    let bottom1 = getRight address
                //    let bottom2 = getRight leftBottom

                //    { leftBottom with 
                //        Column = max bottom1.Column bottom2.Column
                //    }

                [address.Row .. rightBottom.Row]
                |> List.map(fun row ->
                    [address.Column..rightBottom.Column]
                    |> List.map(fun column ->
                        x.Cells.[row, column].Value :?> IConvertible
                        |> ConvertibleUnion.Convert
                    )

                )
                |> array2D
            
            { Array = array 
              Address = address
              SpecificName = None }


        //member x.ReadToAddressedArrays(addresses: ComparableExcelCellAddress list) =
        //    addresses
        //    |> List.map(x.ReadToAddressedArray)
        //    |> List.map(fun m -> m.Address, m)
        //    |> Map.ofList
        //    |> AddressedArrays

        member x.ReadToAddressedArray(address: string) =
            x.ReadToAddressedArray(ComparableExcelCellAddress.OfAddress address)

        member x.ReadToObsevations(address: ComparableExcelCellAddress) =
            x.ReadToAddressedArray(address).Array
            |> Array2D.toLists
            |> List.map(fun m -> m.[0].Text => m.[1])
            |> Observations.Create

        member x.ReadToObsevations(address: string) =
            x.ReadToAddressedArray(address).Array
            |> Array2D.toLists
            |> List.map(fun m -> m.[0].Text => m.[1])
            |> Observations.Create

        //member x.ReadToAddressedArrays(addresses: string list) =
        //    let addresses = addresses |> List.map ComparableExcelCellAddress.OfAddress
        //    x.ReadToAddressedArrays(addresses)


        member x.LoadFromAddressedArrays(addressArrays: AddressedArray list) =
            for addressedArray in addressArrays do 
                let array2D = 
                    addressedArray.Array
                    |> Array2D.map(fun m ->
                        m.Value
                    )

                match addressedArray.SpecificName with 
                | Some tableName ->

                    (VisibleExcelWorksheet.Create x).LoadFromArraysAsTable(
                        array2D,
                        tableName         = tableName,
                        addr              = addressedArray.Address.Address,
                        includingFormula  = true,
                        allowRerangeTable = true
                    )

                | None ->
                    (VisibleExcelWorksheet.Create x).LoadFromArrays(
                        array2D,
                        addr = addressedArray.Address.Address,
                        includingFormula = true
                    )

                //x.Cells.[addressedArray.Address.Address].LoadFromArray2D(array2D)
                //|> ignore


        member x.LoadFromAddressedArrays_NoAddingTableNames(addressArrays: AddressedArray list) =
            for addressedArray in addressArrays do 
                let array2D = 
                    addressedArray.Array
                    |> Array2D.map(fun m ->
                        m.Value
                    )

                let worksheet = x

                let addr = addressedArray.Address.Address


                let range = worksheet.Cells.[addr].LoadFromArray2D(array2D) 

                ()

                //x.Cells.[addressedArray.Address.Address].LoadFromArray2D(array2D)
                //|> ignore


    type AddressedArrays with 
        member x.SaveToXlsx(xlsxPath: XlsxPath, ?sheetName) =
            File.Delete(xlsxPath.Path)
            let excelPackage = new ExcelPackage(xlsxPath.Path)
            let sheet = excelPackage.Workbook.Worksheets.Add(defaultArg sheetName "Sheet1")
            sheet.LoadFromAddressedArrays(x.AsList)
            excelPackage.Save()
            excelPackage.Dispose()
