
module CellScript.Server.UDF
open Fable.Remoting
open CellScript.Shared.UDF

let cellScriptUDFApi: ICellScriptUDFApi =
    { testFromCellScriptServer = fun _ -> async { return "Hello world from Remoting UDF"} }