namespace CellScript.Core
open Deedle
open System
open OfficeOpenXml
open Deedle.Internal
open System.IO
open System.Collections.Generic
open System.Runtime.CompilerServices
open Shrimp.FSharp.Plus

[<AutoOpen>]
module private _Utils =
    let internal fixRawContent (content: obj) =
        match content with
        | :? ConvertibleUnion as v -> v.Value
        | null -> null
        | :? ExcelErrorValue -> null
        | :? string as text -> 
            match text.Trim() with 
            | "" -> null 
            | _ -> text :> IConvertible
        | :? IConvertible as v -> 
            match v with 
            | :? DBNull -> null
            | _ -> v
        | _ -> failwithf "Cannot convert to %A to iconvertible" (content.GetType())

    let private patterns = 
        lazy 
            System.Globalization.DateTimeFormatInfo.CurrentInfo.GetAllDateTimePatterns()
            |> List.ofArray

    let internal fixRawCell (content: ExcelRangeBase) =
        
        match content.Style.Numberformat.Format with 
        | EqualTo "yyyy\-mm\-dd" ->
            let v = 
                match content.Value with 
                | :? double as v -> (System.DateTime.FromOADate v) :>  IConvertible
                | v -> v :?> IConvertible

            v 

        | _ -> fixRawContent content.Value
    

module Constrants = 
    let [<Literal>] SHEET1 = "Sheet1"
    let [<Literal>] ``#N/A`` = "#N/A"

    let [<Literal>] internal  DefaultTableName = "Table1"
    let [<Literal>] internal  CELL_SCRIPT_COLUMN = "CellScriptColumn"



module Extensions =


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

 

    [<RequireQualifiedAccess>]
    module internal Frame =

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

        let internal fillEmptyUpForColumns (columnKeys: _ list) frame =
            Frame.mapCols (fun colKey (column: ObjectSeries<_>) ->
                match List.contains colKey columnKeys with 
                | true ->
                    let values: list<obj option> = 
                        column.GetAllValues()
                        |> List.ofSeq
                        |> List.map OptionalValue.asOption

                    let rec loop accum accumValues values =
                        match values with 
                        | h :: t ->
                            match h with 
                            | Some v -> loop  (Some v) (v :: accumValues) t
                            | None -> 
                                match accum with 
                                | Some accum ->
                                    loop  (Some accum) (accum :: accumValues) t
                                | None -> 
                                    match h with 
                                    | Some h ->
                                        loop  None (h :: accumValues) t
                                    | None -> loop  None (null :: accumValues) t

                        | [] -> accumValues

                    let r =  loop None [] values
                    r
                    |> List.rev
                    |> Series.ofValues
                    |> Series.indexWith (column.Keys)

                | false ->
                    column :> Series<_, _>

                ) frame


        let internal fillEmptyUp frame =
            Frame.mapColValues (fun (column: ObjectSeries<_>) ->
                let values: list<obj option> = 
                    column.GetAllValues()
                    |> List.ofSeq
                    |> List.map OptionalValue.asOption

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

                loop values None []
                |> List.rev
                |> Series.ofValues
                |> Series.indexWith (column.Keys)

                ) frame


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


    type ExcelRangeBase with
        member internal x.LoadFromArray2D(array2D: IConvertible [,]) =
            let baseArray = 
                Array2D.toLists array2D |> List.map (Array.ofList >> Array.map box)
            let range = x.LoadFromArrays(baseArray)
            range


        member range.ReadDatas() =
            //let rowStart = range.Start.Row
            //let rowEnd = range.End.Row
            //let columnStart = range.Start.Column
            //let columnEnd = range.End.Column
            let content =
                array2D
                    [ for i = 0 to range.Rows-1 do 
                        yield
                            [ for j = 0 to range.Columns-1 do yield fixRawCell (range.Offset(i, j, 1, 1)) ]
                    ] 
                |> Array2D.map ConvertibleUnion.Convert

            content

