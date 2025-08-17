namespace CellScript.Core
open Deedle
open System
open OfficeOpenXml
open Deedle.Internal
open System.IO
open System.Collections.Generic
open System.Runtime.CompilerServices
open Shrimp.FSharp.Plus
open Shrimp.FSharp.Plus.Text

[<AutoOpen>]
module private _Utils =
    
    let internal fixRawContent_MapText  (content: obj) f =
        match content with
        | :? ConvertibleUnion as v -> v.Value
        | null -> null
        | :? ExcelErrorValue -> null
        | :? string as text -> 
            let text = text.Replace(Text.MAX_CHAR_65534, "")
            match text.Trim() with 
            | "" -> null 
            | text -> (f text) :> IConvertible
        | :? IConvertible as v -> 
            match v with 
            | :? DBNull -> null
            | _ -> v
        | _ -> failwithf "Cannot convert to %A to Iconvertible" (content.GetType())


    let internal fixRawContent (content: obj) =
        match content with
        | :? ConvertibleUnion as v -> v.Value
        | null -> null
        | :? ExcelErrorValue -> null
        | :? string as text -> 
            let text = text.Replace(Text.MAX_CHAR_65534, "")
            match text.Trim() with 
            | "" -> null 
            | text -> text :> IConvertible
        | :? IConvertible as v -> 
            match v with 
            | :? DBNull -> null
            | _ -> v
        | _ -> failwithf "Cannot convert to %A to Iconvertible" (content.GetType())

    let private patterns = 
        lazy 
            System.Globalization.DateTimeFormatInfo.CurrentInfo.GetAllDateTimePatterns()
            |> List.ofArray

    type ExcelRangeBase with 
        member internal x.IsHidden() =
            x.EntireColumn.Hidden || x.EntireRow.Hidden

    let internal fixRawCell (content: ExcelRangeBase, includeHided: bool) =
        
        match (not includeHided) && (content.IsHidden()) with
        | true -> "" :> IConvertible
        | _ ->
            match content.Style.Numberformat.Format with 
            | "yyyy\-mm\-dd" ->
                let v = 
                    match content.Value with 
                    | :? double as v -> (System.DateTime.FromOADate v) :>  IConvertible
                    | v -> v :?> IConvertible
                v 

            | String.TrimStartIC "General" ending -> 
                //match ending with 
                //| "" -> fixRawContent content.Value
                //| _ -> 
                //    match content.Value with 
                //    | null -> null
                //    | :? ExcelErrorValue -> null
                //    | :? DBNull -> null
                //    | _ -> fixRawContent (content.Text)
                  

                let ending = ending.Trim('\"').TrimEnding("_)")
                match ending with 
                | "" -> fixRawContent content.Value
                | ending ->
                    let v = fixRawContent content.Value |> ConvertibleUnion.Convert
                    match v.Text.Trim() with 
                    | "" -> "" :> IConvertible
                    | _ -> (v.Text + ending) :> IConvertible

            | "#,##0" ->
                fixRawContent_MapText content.Value (fun text ->
                    let text2 = text.Replace(",", "")
                    match Double.tryParse text2 with 
                    | Some v -> v :> IConvertible
                    | None -> text :> IConvertible
                )


            | _ -> fixRawContent content.Text

[<RequireQualifiedAccess>]
module CellText =
    let parseToDateTime(m: obj) =
        match m with 
        | :? double as convertible -> System.DateTime.FromOADate convertible 
        | :? DateTime as v -> v
        | :? string as v -> System.DateTime.Parse v
        | v -> failwithf "Cannot parse %A to datetime" (v.GetType(), v.ToString())

    let getAsODBCNumber(m: string) =
        match m with 
        | String.StartsWith "0" -> 
            match m.Length with 
            | 1 -> Some (System.Double.Parse m)
            | _ -> None
        | String.Contains "," -> None
        | _ ->
            match System.Double.TryParse (m) with 
            | true, v -> 
                match m.RightOf "."   with 
                | Some right ->
                    match right.Trim() with 
                    | "" -> None
                    | right ->
                        let lastChar = right.Chars(right.Length-1)
                        match lastChar with 
                        | '0' -> None
                        | _ ->
                            match System.Int32.TryParse right with 
                            | true, 0 -> None
                            | false, _ -> None
                            | _ ->  
                                //Some v
                                match right.Length < 10 with 
                                | true -> Some (v)
                                | false -> None
                    
                | None ->
                    //Some v
                    match v.ToString().Length < 10 with 
                    | true -> Some (v)
                    | false -> None
            | false, _ -> None



module Constrants = 
    let [<Literal>] SHEET1 = "Sheet1"
    let [<Literal>] ``#N/A`` = "#N/A"

    let [<Literal>] internal  DefaultTableName = "Table1"
    let [<Literal>] internal  CELL_SCRIPT_COLUMN = "CellScriptColumn"




module Extensions =

    [<AutoOpenAttribute>]
    module _Frame_FillEmptyUp =
        [<RequireQualifiedAccess>]
        module Frame =
            [<RequireQualifiedAccess>]
            module Column =
                let fillEmptyUp (column: ObjectSeries<_>) =
                    let values: list<obj option> = 
                        let values = 
                            column.GetAllValues()
                            |> List.ofSeq

                        values
                        |> List.map OptionalValue.asOption
                        |> List.map(fun m ->
                            match m with 
                            | Some (v) ->
                                match v with 
                                | :? string as text -> 
                                    match text.Trim() with 
                                    | "" -> None
                                    | _ -> Some text

                                | _ -> Some v
                            | None -> None
                        )

                    let rec loop values accum accumValues =
                        match values with 
                        | h :: t ->
                            match h with 
                            | Some v -> loop t (Some v) (v :: accumValues)
                            | None -> 
                                match accum with 
                                | Some accum ->
                                    loop t (Some accum) (accum :: accumValues)
                                | None -> 
                                    match h with 
                                    | Some h ->
                                        loop t None (h :: accumValues)
                                    | None -> loop t None (null :: accumValues)

                        | [] -> accumValues

                    let newValues = 
                        loop values None []
                        |> List.rev

                    let series = 
                        newValues
                        |> Series.ofValues

                    series
                    |> Series.indexWith (column.Keys)


            let fillEmptyUpForColumns (columnKeys: _ list) frame =
                Frame.mapCols (fun colKey (column: ObjectSeries<_>) ->
                    match List.contains colKey columnKeys with 
                    | true ->
                        Column.fillEmptyUp column

                    | false ->
                        column :> Series<_, _>

                    ) frame




            let fillEmptyUp frame =
                Frame.mapColValues (fun (column: ObjectSeries<_>) ->
                    Column.fillEmptyUp column
                ) frame

    [<RequireQualifiedAccess>]
    module Array2D =

        let private toSeqs (input: 'a[,]) =
            let l1 = input.GetLowerBound(0)
            let u1 = input.GetUpperBound(0)
            seq {
                for i = l1 to u1 do
                    yield input.[i,*] :> seq<'a>
            }

        let toLists (input: 'a[,]) =
            let l1 = input.GetLowerBound(0)
            let u1 = input.GetUpperBound(0)
            [
                for i = l1 to u1 do
                    yield List.ofArray input.[i,*] 
            ]

        let transpose (input: 'a[,]) =
            let l1 = input.GetLowerBound(1)
            let u1 = input.GetUpperBound(1)
            seq {
                for i = l1 to u1 do
                    yield input.[*,i]
            }
            |> array2D

        let concat__addSuffixEmptyColumnsToForceSameColumnLength emptyValue (space: int) (rowLists: list<'a [,]>) =
            let maxLength = 
                rowLists
                |> List.map Array2D.length2
                |> List.max

            let emptyRow = List.replicate maxLength emptyValue
            let spacedRows = List.replicate space emptyRow 

            rowLists
            |> List.indexed
            |> List.collect(fun (i, rows) ->
                let columnLength = Array2D.length2 rows
                let rows = 
                    match columnLength = maxLength with 
                    | true -> toLists rows
                    | false ->
                        let rows = toLists rows
                        let substract = maxLength - columnLength
                        rows
                        |> List.map(fun row -> 
                            row @ List.replicate substract emptyValue
                        )

                match i = rowLists.Length - 1 with 
                | false -> rows @ spacedRows
                | true -> rows 
            )

            |> array2D



    let lists_ForceSameLength emptyValue (rowLists: list<list<'a>>) =
        let maxLength = 
            rowLists
            |> List.map List.length
            |> List.max

        rowLists
        |> List.map(fun rows ->
            let columnLength = List.length rows
            match columnLength = maxLength with 
            | true -> rows
            | false ->
                let substract = maxLength - columnLength
                rows @ List.replicate substract emptyValue
        )


    let lists_ForceSameLength_fillEmptyHeadersWithNumber emptyValue (rowLists: list<list<string>>) =
        match lists_ForceSameLength emptyValue rowLists with 
        | [] -> []
        | headers :: contents ->
            let headers =
                let mutable i = 
                    headers
                    |> List.choose(Int32.tryParse)
                    |> function
                        | [] -> 0
                        | vs -> List.max vs

                headers
                |> List.map(fun header ->
                    match header.Trim() with 
                    | "" -> 
                        i <- i + 1
                        i.ToString()
                    | _ -> header
                )

            headers :: contents

    let lists_ForceSameLength_fillEmptyHeadersWith_UnderLines emptyValue (rowLists: list<list<string>>) =
        match lists_ForceSameLength emptyValue rowLists with 
        | [] -> []
        | headers :: contents ->
            let headers =
                let mutable i = 
                    headers
                    |> List.choose(Int32.tryParse)
                    |> function
                        | [] -> 0
                        | vs -> List.max vs

                headers
                |> List.map(fun header ->
                    match header.Trim() with 
                    | "" -> 
                        i <- i + 1

                        List.replicate i "_"
                        |> String.concat ""

                    | _ -> header
                )

            headers :: contents


    [<RequireQualifiedAccess>]
    module internal Frame =
        /// Frame.indexRowsOrdinally
        let Concat_RefreshRowKeys(frames: AtLeastOneList<Frame<_, _>>) =
            match frames.AsList with 
            | [ frame ] -> Frame.indexRowsOrdinally frame
            | _ ->
                let headerLists =
                    frames.AsList
                    |> List.map (fun m ->
                        m.ColumnKeys
                        |> Set.ofSeq
                    )


                headerLists
                |> List.reduce(fun headers1 headers2 -> 
                    if headers1 <> headers2 then failwithf "headers1 %A <> headers2 %A when concating table" (Set.toList headers1) (Set.toList headers2)
                    else headers2
                )
                |> ignore

                let rows =

                    let indexedColkeys =
                        frames.Head.ColumnKeys
                        |> List.ofSeq
                        |> List.indexed
                        |> List.map (fun (id ,colKey) ->
                            colKey, (id, colKey)
                        )
                        |> dict

                    frames.AsList
                    |> List.collect (fun frame ->
                        let frame =
                            frame
                            |> Frame.mapColKeys(fun colKey -> indexedColkeys.[colKey])
                            |> Frame.sortColsByKey
                            |> Frame.mapColKeys snd

                        frame.Rows.ValuesAll
                        |> List.ofSeq
                    )

                Frame.ofRowsOrdinal rows
                |> Frame.indexRowsOrdinally

        [<System.Obsolete("This method is obsolted, using RefreshRowKeys instead")>]
        let Concat_RemoveRowKeys(frames: AtLeastOneList<Frame<_, _>>) =
            Concat_RefreshRowKeys(frames)

        let Concat_KeepRowKeys(frames: AtLeastOneList<Frame<_, _>>) =
            match frames.AsList with
            | [frame] -> frame
            | _ ->
                let headerLists =
                    frames.AsList
                    |> List.map (fun m ->
                        m.ColumnKeys
                        |> Set.ofSeq
                    )

                headerLists
                |> List.reduce(fun headers1 headers2 -> 
                    if headers1 <> headers2 then failwithf "headers1 %A <> headers2 %A when concating table" headers1 headers2
                    else headers2
                )
                |> ignore

                let indexedColkeys =
                    frames.Head.ColumnKeys
                    |> List.ofSeq
                    |> List.indexed
                    |> List.map (fun (id ,colKey) ->
                        colKey, (id, colKey)
                    )
                    |> dict

                let rows =
                    frames.AsList
                    |> List.collect (fun frame ->

                        let frame =
                            frame
                            |> Frame.mapColKeys(fun colKey -> indexedColkeys.[colKey])
                            |> Frame.sortColsByKey
                            |> Frame.mapColKeys snd


                        frame.GetRows()
                        |> Series.observations
                        |> List.ofSeq
                    )

                Frame.ofRows rows

        let chooseCols chooser (frame) =
            frame
            |> Frame.filterCols(fun colKey _ ->
                match chooser colKey with 
                | Some _ -> true
                | None -> false
            )
            |> Frame.mapColKeys(fun colKey ->
                (chooser colKey).Value
            )




        let internal splitRowToMany addtionalHeaders (mapping : 'R -> ObjectSeries<_> -> seq<seq<obj>>)  (frame: Frame<'R,'C>) =
            let headers = 
                let keys = frame.ColumnKeys
                Seq.append keys addtionalHeaders

            let values = 
                frame.Rows.ValuesAll
                |> Seq.mapi (fun i value -> (mapping (Seq.item i frame.RowKeys) value))
                |> Seq.concat
                |> array2D

            match values.Length with 
            | 0 -> failwith "Cannot split row to many as current frame is empty"
            | _ -> 
                Frame.ofArray2D values
                |> Frame.indexColsWith headers

    type internal MergeCellAddrs = 
        { Addresses: ComparableExcelAddress list 
          HiddenCache: ConcurrentDictionary<ComparableExcelAddress, bool> }
    with 
        member x.AsList = x.Addresses

        static member TryCreate(mergedCells: ExcelWorksheet.MergeCellsCollection) =
            let addresses = 
                mergedCells
                |> List.ofSeq
                |> List.map(fun m -> ComparableExcelAddress.OfAddress m)
            
            match addresses with 
            | [] -> None
            | _ ->
                { Addresses = addresses
                  HiddenCache = ConcurrentDictionary() }
                |> Some

        member x.TryGetMergedOf(sheet: ExcelWorksheet, addr: ComparableExcelCellAddress) =
            let addr = 
                x.AsList
                |> List.tryFind(fun m ->
                    m.Contains addr
                )

            addr
            |> Option.map(fun addr ->
                let isAllCellsHidden = 
                    x.HiddenCache.GetOrAdd(addr, valueFactory = fun addr ->
                        let r = sheet.Cells.[addr.Address]
                        let isAllCellsHidden = 
                            r
                            |> Seq.forall(fun m ->
                                m.IsHidden()
                            )

                        isAllCellsHidden
                    )

                {|
                    IsAllCellsHidden = isAllCellsHidden
                    MergedAddr = addr
                |}

                
            )


    type ExcelWorksheet with 
        member internal x.TryGetMergeCellAddrs() =
            x.MergedCells
            |> MergeCellAddrs.TryCreate


    type ExcelRangeBase with
        member internal x.LoadFromArray2D(array2D: IConvertible [,]) =
            let baseArray = 
                Array2D.toLists array2D |> List.map (Array.ofList >> Array.map box)
            let range = x.LoadFromArrays(baseArray)
            range

        member private range.ToLists() =
            let content =
                [ for i = 0 to range.Rows-1 do 
                    yield
                        [ for j = 0 to range.Columns-1 do yield (range.Offset(i, j, 1, 1)) ]
                ] 

            content

        member internal range.ReadDatasWithUserState_TrackMergeRange(fUserState, includeHided, mergedCellAddrs: MergeCellAddrs option) =
            //let rowStart = range.Start.Row
            //let rowEnd = range.End.Row
            //let columnStart = range.Start.Column
            //let columnEnd = range.End.Column
            let fixRawCellEx (currentRange: ExcelRangeBase, includedHided) =
                match mergedCellAddrs with 
                | None -> fixRawCell(currentRange, includedHided)
                | Some mergedCellAddrs ->
                    match includedHided with 
                    | true -> fixRawCell(currentRange, includedHided)
                    | false ->
                        match currentRange.Merge with 
                        | false -> fixRawCell(currentRange, includedHided)

                        | true ->
                            let r = fixRawCell(currentRange, true)
                            let isEmpty = 
                                match r with 
                                | null -> true
                                | :? string as v ->
                                    match v.Trim() with 
                                    | "" -> true
                                    | _ -> false

                                | _ -> false

                            match isEmpty with 
                            | true -> ""
                            | false -> 
                                match currentRange.IsHidden() with 
                                | true -> 
                                    let currentRangeAddr = 
                                        currentRange
                                        |> ComparableExcelCellAddress.OfRange

                                    match mergedCellAddrs.TryGetMergedOf(range.Worksheet, currentRangeAddr) with 
                                    | None -> failwith "Invalid token"
                                    | Some merged ->
                                        match merged.IsAllCellsHidden with 
                                        | true -> ""
                                        | false -> r

                                | false -> r

            let content =
                array2D
                    [ for i = 0 to range.Rows-1 do 
                        yield
                            [ for j = 0 to range.Columns-1 do 
                                let currentRange = (range.Offset(i, j, 1, 1))
                                yield (
                                    {|
                                        Content = 
                                            fixRawCellEx (currentRange, includeHided)
                                            |> ConvertibleUnion.Convert

                                        UserState = fUserState currentRange
                                        CurrentRange = currentRange
                                    |}
                                )
                            ]
                    ] 

            let content = 
                match includeHided with 
                | true -> content
                | false ->
                    let rows = 
                        Array2D.toLists content
                        
                    let row1 = rows.[0]
                    let colIndexes_hided =
                        row1
                        |> List.indexed
                        |> List.filter(fun (i, m) ->
                            m.CurrentRange.EntireColumn.Hidden
                        )
                        |> List.map fst

                    let rows =
                        rows
                        |> List.map(fun row ->
                            row
                            |> List.indexed
                            |> List.filter(fun (i, _) ->
                                not(List.contains i colIndexes_hided)
                            )
                            |> List.map snd
                        )

                    array2D rows

            content

        member range.ReadDatasWithUserState(fUserState, includeHided) =
            range.ReadDatasWithUserState_TrackMergeRange(fUserState, includeHided, mergedCellAddrs = None)


        member range.ReadDatas(includeHided) =
            //let rowStart = range.Start.Row
            //let rowEnd = range.End.Row
            //let columnStart = range.Start.Column
            //let columnEnd = range.End.Column
            range.ReadDatasWithUserState(ignore, includeHided)
            |> Array2D.map(fun m -> m.Content)
