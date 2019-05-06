namespace CellScript.Core.Tests

open CellScript.Core
open CellScript.Core.Types

[<AutoOpen>]
module Msg =

    [<RequireQualifiedAccess>]
    type InnerMsg =
        | TestString of string
        | TestExcelReference of SerializableExcelReference

    [<RequireQualifiedAccess>]
    type OuterMsg = // That's the one the update expects
        | TestTable of Table
        | [<SubMsg>] TestInnerMsg of InnerMsg
        | TestString of string
        | TestExcelVector of ExcelVector
        | TestString2Params of text1: string * text2: string
