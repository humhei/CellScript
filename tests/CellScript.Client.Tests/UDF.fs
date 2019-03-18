module CellScript.Client.Tests.UDF

open ExcelDna.Integration
open CellScript.Core.Tests
open ExcelProcess
open MatrixParsers
open FParsec
open CellScript.Core
open CellScript.Core.Tests.Extensions
open CellScript.Core.Tests.PivotTable

[<ExcelFunction>]
let returnZH text =
    String.returnZH

[<ExcelFunction>]
let eurHeader (values: string []) =
    let p = mxMany (!^^ pint32 Shoes.isEur)
    let state,stateValues,others = Seq.partitionByMatrixParsers [p] values
    state
    |> ExcelSeries.row

[<RequireQualifiedAccess>]
module PlainPivotTable =

    let [<Literal>] prefix = "pvt."
    [<ExcelFunction(Name= prefix + "transfer")>]
    let transfer (pvt: PlainPivotTable) =
        PlainPivotTable.transfer pvt