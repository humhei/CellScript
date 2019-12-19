namespace CellScript.Core.Tests

open CellScript.Core
open CellScript.Core.Types

[<AutoOpen>]
module Msg =

    [<RequireQualifiedAccess>]
    type InnerMsg =
        | TestString of string
        | TestExcelReference of ExcelRangeContactInfo
    
    [<RequireQualifiedAccess>]
    type NameInheritedInnerMsg =
        | TestString of string

    type MyRecord =
        { Name: string 
          Id: int }

    [<RequireQualifiedAccess>]
    type OuterMsg = // That's the one the update expects
        | TestTable of Table
        | [<SubMsg>] TestInnerMsg of InnerMsg
        | TestFunctionSheetReference of FunctionSheetReference
        | TestRecords of Records<MyRecord>
        | TestString of string
        | TestStringArray of string []
        | TestExcelVector of ExcelVector
        | TestString2Params of text1: string * text2: string
        | TestCommand of CommandCaller
        | TestFailWith of string
        | [<NameInheritedSubMsg>] TestNameInherited  of NameInheritedInnerMsg