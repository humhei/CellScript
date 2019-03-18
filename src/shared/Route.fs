namespace CellScript.Shared

[<RequireQualifiedAccess>]
module Route =
    let host = "127.0.0.1"
    let port = 9090
    /// Defines how routes are generated on server and mapped from client
    let hostName = sprintf "http://%s:%d" host port

    let routeBuilder typeName methodName =
        sprintf "/api/%s/%s" typeName methodName

    let urlBuilder typeName methodName =
        let route = 
            routeBuilder typeName methodName |> sprintf "%s%s" hostName
        route


/// A type that specifies the communication protocol between client and server
/// to learn more, read the docs at https://zaid-ajaj.github.io/Fable.Remoting/src/basics.html
type ICellScriptApi =
    { eval : string -> string -> Async<string>
      test: unit -> Async<string> }
