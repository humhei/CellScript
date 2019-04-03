namespace CellScript.Core
open Deedle
open Registration
open CellScript.Core.Extensions
open Types
open Newtonsoft.Json
open Akka.Util



[<CustomParamConversion;JsonObject(MemberSerialization.OptIn)>]
type Table(frame: ExcelFrame<int, string>) =

    [<JsonProperty>]
    let serializedData = ExcelFrame.toArray2DWithHeader frame

    [<JsonConstructor>]
    private new (array2D: obj[,]) =
        let frame = ExcelFrame.ofArray2DWithHeader array2D
        new Table(frame)

    member x.AsFrame = frame.AsFrame

    member x.AsExcelFrame = frame
            
    member x.AsArray2D = serializedData

    static member TempData =
        let param =
            array2D [[box "Header1";box "Header2"]; [1; 2]]

        Table(param)


    interface ISurrogated with 
        member x.ToSurrogate(system) = 
            Array2D x.AsArray2D :> ISurrogate


and Array2D = Array2D of obj [,]
with 
    interface ISurrogate with 
        member x.FromSurrogate(system) = 
            let (Array2D array2D) = x
            let excelFrame  = ExcelFrame.ofArray2DWithHeader array2D
            new Table(excelFrame) :> ISurrogated

        

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
