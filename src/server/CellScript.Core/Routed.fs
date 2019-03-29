namespace CellScript.Core


[<RequireQualifiedAccess>]
module Routed =
    let port = 9000
    let endpoint = "/socket"
    let host = "127.0.0.1"

    let private url = sprintf "%s:%d%s" host port endpoint
    let httpUrl = sprintf "http://%s" url
    let wsUrl = sprintf "ws://%s" url


type FirstInnerMsg =
  | FIA
  | FIB
type SecondInnerMsg =
  | SIA
  | SIB

type OuterMsg = // That's the one the update expects
  | No
  | First of FirstInnerMsg
  | Second of SecondInnerMsg
