namespace CellScript.Core
open Deedle
open System
open OfficeOpenXml
open Deedle.Internal
open System.IO
open System.Collections.Generic
open System.Runtime.CompilerServices
open Shrimp.FSharp.Plus


module Constrants = 
    let [<Literal>] SHEET1 = "Sheet1"


module Extensions =


    [<RequireQualifiedAccess>]
    module CellValue =

        let private isTextEmpty (v: obj) =
            match v with
            | :? string as v -> v = ""
            | _ -> false

        let private isEmpty (v: obj) =
            isTextEmpty v

        let private isNotEmpty (v: obj) =
            isEmpty v |> not

        /// no missing or cell empty
        let (|HasSense|NoSense|) (v: obj option) =
            match v with
            | Some v -> if isEmpty v then NoSense else HasSense v
            | None -> NoSense




    [<RequireQualifiedAccess>]
    module Array2D =

        let toSeqs (input: 'a[,]) =
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
    module Frame =

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

                        frame.Rows.Values
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


        let mapValuesString mapping frame =
            let mapping raw = mapping (raw.ToString())
            Frame.mapValues mapping frame
        
        let internal fillEmptyUp frame =
            Frame.mapColValues (fun column ->
                let values = 
                    column.GetAllValues()
                    |> List.ofSeq
                    |> List.map OptionalValue.asOption

                let rec loop values accum accumValues =
                    match values with 
                    | h :: t ->
                        match h with 
                        | CellValue.HasSense v -> loop t (Some v) (v :: accumValues)
                        | CellValue.NoSense -> 
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
                frame.Rows.Values
                |> Seq.mapi (fun i value -> mapping (Seq.item i frame.RowKeys) value)
                |> Seq.concat
                |> array2D

            match values.Length with 
            | 0 -> failwith "Cannot split row to many as current frame is empty"
            | _ -> 
                Frame.ofArray2D values
                |> Frame.indexColsWith headers


    type ExcelRangeBase with
        member internal x.LoadFromArray2D(array2D: obj [,]) =
            let baseArray = Array2D.toSeqs array2D |> Seq.map Array.ofSeq
            x.LoadFromArrays(baseArray)



