namespace CellScript.Core
open Deedle
open Extensions
open System
open OfficeOpenXml
open System.IO
open Shrimp.FSharp.Plus

[<AutoOpen>]
module Observations = 

    [<RequireQualifiedAccess>]
    type DataAddingPosition =
        | BeforeFirst
        | ByIndex of int
        | After of StringIC
        | Before of StringIC
        | AfterLast 

    [<StructuredFormatDisplay("Observation {ObservationValue}")>]
    type Observation = Observation of StringIC * ConvertibleUnion
    with 
        member x.ObservationValue = 
            let (Observation (key, value)) = x
            key, value.Value

        member x.Key =
            let (Observation (key, value)) = x
            key

        member x.Value =
            let (Observation (key, value)) = x
            value

        member x.Pair = x.Key, x.Value

    let observation (a,b) =
        Observation (a,b)

    type Observations = private Observations of Observation list
    with 
        member x.AsList =
            let (Observations v) = x
            v

        
        member x.Keys = x.AsList |> List.map(fun m -> m.Key)
        member x.Values = x.AsList |> List.map(fun m -> m.Value)

        member x.Item(key: string) =
            let key = StringIC key
            let r = 
                x.AsList
                |> List.tryFind(fun observation ->
                    observation.Key = key
                )

            match r with 
            | Some r -> r
            | None -> failwithf "Cannot found key %A in observations %A" key x.AsList

        member x.Item(keys: string seq) =
            keys
            |> List.ofSeq
            |> List.map(fun key ->
                x.[key]
            )
            |> Observations


        member x.MapValue(mapping) =
            x.AsList
            |> List.map(fun m ->
                (m.Key, mapping m)
                |> observation
            )
            |> Observations

        member x.UpdateBy(observations: Observations) =
            let __ensureAllUpdatingKeysExistsInOriginKeys =
                let xKeys = x.Keys
                observations.Keys
                |> List.iter(fun m ->
                    match List.contains m xKeys with 
                    | true -> ()
                    | false -> failwithf "Not all updating keys %A are included in originKeys %A" observations.Keys xKeys
                )

            x.AsList
            |> List.map(fun pair ->
                observations.AsList
                |> List.tryFind(fun pair' ->
                    pair'.Key = pair.Key
                )
                |> function
                    | Some v -> observation(pair.Key, v.Value)
                    | None -> pair
            )
            |> Observations


        member x.ObservationValues =
            let (Observations v) = x
            v
            |> List.map(fun m-> m.ObservationValue)

        static member Create(list: Observation seq, checkKeysAreNotDuplicated) =
            let list = List.ofSeq list
            let __ensureKeyNotDuplicated =
                match checkKeysAreNotDuplicated with 
                | true ->
                    let ditinctedKeys =
                        list
                        |> List.countBy (fun m -> m.Key)

                    ditinctedKeys
                    |> List.tryFind (fun (_ ,l) -> l <> 1)
                    |> function
                        | Some v ->
                            failwithf "Cannot create Observations as duplicated keys %A are found in %A" v list
                        | None -> ()
                | false -> ()
            list
            |> Observations


        static member Create(list: Observation seq) =
            Observations.Create(list, false)

        static member Create(list: seq<StringIC * ConvertibleUnion>) =
            Observations.Create(list |> List.ofSeq |> List.map observation, false)

        static member (@@) (v1: Observations, v2: Observations) =
            v1.AsList @ v2.AsList
            |> Observations

        member x.Add(obsevations: Observations, ?position) =
            let position = defaultArg position DataAddingPosition.AfterLast
            match position with 
            | DataAddingPosition.BeforeFirst -> 
                obsevations @@ x  

            | DataAddingPosition.AfterLast ->
                x @@ obsevations

            | _ ->
                let originKeys = 
                    x.AsList
                    |> List.map(fun m -> m.Key)

                let addingIndex = 
                    match position with 
                    | DataAddingPosition.BeforeFirst  
                    | DataAddingPosition.AfterLast -> failwith "Invalid token"
                    | DataAddingPosition.ByIndex i -> i 
                    | DataAddingPosition.After k ->
                        originKeys
                        |> List.tryFindIndex(fun (k') -> k' = k)
                        |> function
                            | Some i -> i
                            | None -> failwithf "Invalid data adding position %A, all keys are %A" position originKeys

                    | DataAddingPosition.Before k ->
                        originKeys
                        |> List.tryFindIndex(fun (k') -> k' = k)
                        |> function
                            | Some i -> i-1
                            | None -> failwithf "Invalid data adding position %A, all keys are %A" position originKeys

                x.AsList.[0..addingIndex] @ obsevations.AsList @ x.AsList.[addingIndex+1..]
                |> Observations
                
        member x.Add(observation, ?position) =
            x.Add(Observations [observation], ?position = position)

        member x.Add(key, value, ?position) =
            x.Add(observation(key, value), ?position = position)


    let observations (list: seq<Observation>) =
        list
        |> List.ofSeq
        |> Observations.Create
