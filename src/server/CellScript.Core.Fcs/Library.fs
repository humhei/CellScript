namespace CellScript.Core

open Types
open Cluster

[<RequireQualifiedAccess>]
type FcsMsg =
    | EditCode of CommandCaller
    | Eval of input: ExcelRangeContactInfo * code: string
    | Sheet_Active of CellScriptEvent<SheetReference>
    | Workbook_BeforeClose of CellScriptEvent<string>
with 
    static member CommandSettings() =
        [ <@@ FcsMsg.EditCode @@>, { Shortcut = Some "+^{F11}" }]

[<RequireQualifiedAccess>]
module Route =
    let port = 9060
