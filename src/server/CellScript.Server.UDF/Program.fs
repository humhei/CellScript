
module CellScript.Server.UDF
open CellScript.Shared.UDF

let cellScriptUDFApi: ICellScriptUDFApi =
    { 
        testFromCellScriptServer = fun input -> async { return "Hello world from Remoting UDF" + input }
        testTableParameter = fun table -> async {return "table as paramter"} 
    }