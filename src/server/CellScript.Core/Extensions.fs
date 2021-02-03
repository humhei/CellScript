namespace CellScript.Core
open Deedle
open System
open OfficeOpenXml
open Deedle.Internal
open System.IO
open System.Collections.Generic
open System.Runtime.CompilerServices


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
            | Some value -> if isEmpty v then NoSense else HasSense value
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
        let mapValuesString mapping frame =
            let mapping raw = mapping (raw.ToString())
            Frame.mapValues mapping frame
        
        let internal fillEmptyUp frame =
            Frame.mapColValues (fun column ->
                column 
                |> Series.mapAll (fun index value ->
                    let rec searchValue index value =
                        match value with 
                        | CellValue.HasSense v -> Some v
                        | _ -> 
                            let newIndex = (index - 1)
                            if Seq.contains newIndex column.Keys 
                            then searchValue newIndex (Series.tryGet newIndex column)
                            else value
                    searchValue index value
                )) frame


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
        member x.LoadFromArray2D(array2D: obj [,]) =
            let baseArray = Array2D.toSeqs array2D |> Seq.map Array.ofSeq
            x.LoadFromArrays(baseArray)



