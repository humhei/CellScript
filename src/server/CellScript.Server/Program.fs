// Learn more about F# at http://fsharp.org

open Suave
open CellScript.Shared
open Suave.Filters
open Suave.Operators
open CellScript.Server.UDF
open CellScript.Shared
open Elmish
open Elmish.Bridge

[<EntryPoint>]
let main argv =
    let webApp = 
        let webPart =
            Remoting.createApi()
            |> Remoting.fromValue cellScriptUDFApi
            |> Remoting.withRouteBuilder UDF.port.RouteBuilder
            |> Remoting.buildWebPart

        choose [
            path "/ping"   >=> Successful.OK "pong"
            webPart ]

    let suaveConfig =
        { defaultConfig with
            bindings   = [ HttpBinding.createSimple HTTP host UDF.port.Value ] }

    startWebServer suaveConfig webApp

    0 // return an integer exit code
