namespace CellScript.Core

open Types
open Cluster

[<RequireQualifiedAccess>]
type FcsMsg =
    | EditCode of CommandCaller
    | Eval of input: SerializableExcelReference * code: string
    | Sheet_Active of CellScriptEvent<SheetActiveArg>
    | Workbook_BeforeClose of CellScriptEvent<string>
with 
    static member CommandSettings() =
        [ <@@ FcsMsg.EditCode @@>, { Shortcut = Some "+^{F11}" }]

