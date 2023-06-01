namespace CellScript.Core
open System
open System.IO
open Shrimp.FSharp.Plus
open System.Diagnostics
open CellScript.Core.Extensions
open OfficeOpenXml


[<RequireQualifiedAccess>]
type AddressKind =
    | Exactly
    | UnderTableCase of string * distance: int
with 
    /// defaultArg distance 1
    static member UnderTable(tableName, ?distance) =
        AddressKind.UnderTableCase(tableName, defaultArg distance 1)

type AddressedArray =
    { Address: ComparableExcelCellAddress 
      Array: ConvertibleUnion [,]
      SpecificName: string option
      AddressKind: AddressKind
    }
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
    static member ofValue address (value: IConvertible) =
        let array = 
            [
                [
                    ConvertibleUnion.Convert value
                ]
            ]
            |> array2D

        { Address = ComparableExcelCellAddress.OfAddress address 
          Array = array
          SpecificName = None
          AddressKind = AddressKind.Exactly }

    static member OfValue (address, value: IConvertible) =
        let array = 
            [
                [
                    ConvertibleUnion.Convert value
                ]
            ]
            |> array2D

        { Address = ComparableExcelCellAddress.OfAddress address 
          Array = array
          SpecificName = None
          AddressKind = AddressKind.Exactly }

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
          SpecificName = Some (array.[0, 0].Text.ForceEndingWith("表"))
          AddressKind = AddressKind.Exactly }

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
          SpecificName = Some (array.[0, 0].Text.ForceEndingWith("表"))
          AddressKind = AddressKind.Exactly }



    /// As Table
    static member ofTitledArray address  (title: string, array: ConvertibleUnion list) =
        
        { Address = ComparableExcelCellAddress.OfAddress address 
          Array = 
            ConvertibleUnion.Convert title :: array 
            |> List.map List.singleton
            |> array2D
          SpecificName = Some (title.ForceEndingWith("表"))
          AddressKind = AddressKind.Exactly
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
          SpecificName = Some (title.[0].ForceEndingWith("表")) 
          AddressKind = AddressKind.Exactly }


    /// As Table
    static member ofTable address (table: Table) =

        let content = 
            table.ToArray2D()
            |> Array2D.map ConvertibleUnion.Convert

        match table.RowCount with 
        | 0 -> 
            { Address = ComparableExcelCellAddress.OfAddress address 
              Array = Array2D.zeroCreate 0 0 
              SpecificName = None 
              AddressKind = AddressKind.Exactly
             }

        | _ ->
            { Address = ComparableExcelCellAddress.OfAddress address 
              Array = content
              SpecificName = Some (content.[0,0].Text.ForceEndingWith("表"))
              AddressKind = AddressKind.Exactly
             }
       

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
              SpecificName = None
              AddressKind = AddressKind.Exactly }


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


        member x.LoadFromAddressedArrays(addressArrays: AddressedArray list, ?columnAutofitOptions) =
            for addressedArray in addressArrays do 
                let array2D = 
                    addressedArray.Array
                    |> Array2D.map(fun m ->
                        m.Value
                    )

                match Array2D.length1 array2D * Array2D.length2 array2D with 
                | 0 -> failwithf "AddressedArray %A data is empty" (addressedArray)
                | _ ->

                    //if addressedArray.Address.Address = "G319"
                    //then ()

                    let addr =
                        match addressedArray.AddressKind with 
                        | AddressKind.Exactly -> addressedArray.Address.Address
                        | AddressKind.UnderTableCase (tableName, distance) ->
                            let findedTable =
                                x.Tables
                                |> Seq.tryFind(fun m -> StringIC m.Name = StringIC tableName)

                            match findedTable with 
                            | None -> 
                                let tableNames =
                                    x.Tables
                                    |> List.ofSeq
                                    |> List.map(fun m -> m.Name)


                                failwithf 
                                    "Cannot find table %s in worksheets\navaliable tableNames are %A" 
                                    tableName
                                    tableNames

                            | Some table ->
                                let row = table.Address.End.Row + distance
                                let column = addressedArray.Address.Column
                                { Row = row; Column = column }.Address


                    match addressedArray.SpecificName with 
                    | Some tableName ->

                        (VisibleExcelWorksheet.Create x).LoadFromArraysAsTable(
                            array2D,
                            tableName         = tableName,
                            addr              = addr,
                            includingFormula  = true,
                            allowRerangeTable = true,
                            columnAutofitOptions = defaultArg columnAutofitOptions ColumnAutofitOptions.None
                        )

                    | None ->
                        (VisibleExcelWorksheet.Create x).LoadFromArrays(
                            array2D,
                            addr = addr,
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





    type NamedAddressedArrays =
        { SheetName: string 
          AddressedArrays: AddressedArrays
          SpecificTargetSheetName: string option }
    with 
        member x.TargetSheetName =
            match x.SpecificTargetSheetName with 
            | None -> x.SheetName
            | Some sheetName -> sheetName


    type NamedAddressedArrayLists = NamedAddressedArrayLists of NamedAddressedArrays al1List
    with 
        member x.Value =
            let (NamedAddressedArrayLists v) = x
            v

        member x.SaveToXlsx_FromTemplate(template: XlsxFile, xlsxPath: XlsxPath) =
            let dir = Path.GetDirectoryName (xlsxPath.Path)
            let _ = Directory.CreateDirectory(dir) |> ignore
            
            match template.XlsxPath = xlsxPath with 
            | true -> ()
            | false -> File.Copy(template.Path, xlsxPath.Path, true)

            let excelPackage = new ExcelPackage(xlsxPath.Path)

            x.Value.AsList
            |> List.iter(fun namedAddressedArray ->
                
                let sheet = excelPackage.Workbook.Worksheets.[(namedAddressedArray.SheetName)]
                match sheet with 
                | null -> failwithf "excelWorksheet is null"
                | _ -> ()
                match namedAddressedArray.SpecificTargetSheetName with 
                | None -> ()
                | Some sheetName ->
                    sheet.Name <- sheetName

                sheet.LoadFromAddressedArrays(namedAddressedArray.AddressedArrays.AsList, columnAutofitOptions = ColumnAutofitOptions.None)
            )

            excelPackage.Save()
            excelPackage.Dispose()


    type NamedAddressedArrays with 
        member x.SaveToXlsx_FromTemplate(template: XlsxFile, xlsxPath: XlsxPath) =
            let x = 
                x
                |> AtLeastOneList.singleton
                |> NamedAddressedArrayLists
            x.SaveToXlsx_FromTemplate(template, xlsxPath)
    

    type AddressedArrays with 
        member x.SaveToXlsx(xlsxPath: XlsxPath, ?sheetName) =
            if File.Exists(xlsxPath.Path)
            then File.Delete(xlsxPath.Path)

            let dir = Path.GetDirectoryName (xlsxPath.Path)
            let _ = Directory.CreateDirectory(dir) |> ignore


            let excelPackage = new ExcelPackage(xlsxPath.Path)
            let sheet = excelPackage.Workbook.Worksheets.Add(defaultArg sheetName "Sheet1")
            sheet.LoadFromAddressedArrays(x.AsList, columnAutofitOptions = ColumnAutofitOptions.DefaultValue)
            excelPackage.Save()
            excelPackage.Dispose()



        member x.SaveToXlsx_FromTemplate(template: XlsxFile, xlsxPath: XlsxPath, ?templateSheetName, ?targetSheetName) =
            let namedAddressedArray =
                { SheetName = defaultArg templateSheetName "Sheet1" 
                  SpecificTargetSheetName = targetSheetName 
                  AddressedArrays = x }
            
            namedAddressedArray.SaveToXlsx_FromTemplate(template, xlsxPath)