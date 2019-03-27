namespace CellScript.Core.Tests
open CellScript.Core.Types
module Common =
    let getRef (ref: SerializableExcelReference) =
        sprintf "ColumnFirst:%d \n RowFirst:%d" ref.ColumnFirst ref.RowFirst