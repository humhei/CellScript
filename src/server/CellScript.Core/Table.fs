namespace CellScript.Core
open Deedle
open Registration
open CellScript.Core.Extensions
open Types
open Newtonsoft.Json

[<CustomParamConversion;JsonObject(MemberSerialization.OptIn)>]
type Table(frame: ExcelFrame<int, string>) =

    [<JsonProperty>]
    let serializedData = ExcelFrame.toArray2DWithHeader frame

    [<JsonConstructor>]
    new (array2D: obj[,]) =
        let frame = ExcelFrame.ofArray2DWithHeader array2D
        new Table(frame)

    member x.AsFrame = frame.AsFrame

    member x.AsExcelFrame = frame
            
    static member op_ErasedCast(array2D: obj[,]) = 
        Table(array2D)


[<RequireQualifiedAccess>]
module Table =

    let mapFrame mapping (table: Table) =
        mapping table.AsFrame
        |> ExcelFrame
        |> Table

    let item (indexes: seq<string>) table =
        let mapping =
            Frame.filterCols (fun key _ ->
                Seq.exists (String.equalIgnoreCaseAndEdgeSpace key) indexes
            )

        mapFrame mapping table
