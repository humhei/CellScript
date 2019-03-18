namespace CellScript.Core.Tests
open LiteDB.FSharp
open LiteDB
module LiteDB =

    let db =
        let mapper = FSharpBsonMapper()
        new LiteDatabase("simple.db", mapper)
