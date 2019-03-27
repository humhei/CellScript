namespace CellScript.Core


[<RequireQualifiedAccess>]
module Routed =
    let port = 9000
    let endpoint = "/socket"
    let host = "127.0.0.1"

    let private url = sprintf "%s:%d%s" host port endpoint
    let httpUrl = sprintf "http://%s" url
    let wsUrl = sprintf "ws://%s" url


[<RequireQualifiedAccess>]
type OuterMsg = // That's the one the update expects
  | SomeMsg of table: Table
  | None

