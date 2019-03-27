namespace CellScript.Core.Tests
open CellScript.Core.Types
module Common =
    let getRef (ref: ExcelReference) =
        sprintf "ColumnFirst:%d \n RowFirst:%d" ref.ColumnFirst ref.RowFirst