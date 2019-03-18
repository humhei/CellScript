module CellScript.Client.UDF

open CellScript.Core
open ExcelDna.Integration
open CellScript.Core.UDF
open CellScript.Client.Extensions
open Fable.Remoting.DotnetClient
open CellScript.Shared
open Newtonsoft.Json

[<RequireQualifiedAccess>]
module ExcelArray =

    [<ExcelFunction>]
    let leftOf (input: ExcelArray) (pattern: string)  =
        ExcelArray.leftOf pattern input

[<RequireQualifiedAccess>]
module Table =
    let [<Literal>] private Prefix = "tb."

    [<ExcelFunction(Name = "it")>]
    let it table (indexes: string []) =
        Table.item indexes table


[<ExcelFunction>]
let raw (data: obj[,]) =
    data

[<ExcelFunction(IsMacroType=true)>]
let range ([<ExcelArgument(AllowReference = true)>] range: obj) =
    let xlRef = range :?> ExcelReference
    ExcelRangeBase.mapByXlRef xlRef (fun range ->
        range.Text
    )

let proxy = Proxy.create<ICellScriptApi>(Route.urlBuilder)

[<ExcelFunction>]
let eval (input: obj[,]) (script: string) =
    async {
        let input = JsonConvert.SerializeObject(input)
        let! result = proxy.call(<@fun server -> server.eval input script@>)
        return JsonConvert.DeserializeObject<obj[,]> result
    } |> Async.StartAsTask