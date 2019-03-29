// Learn more about F# at http://fsharp.org

open Suave
open Suave.Filters
open Suave.Operators
open Elmish
open Elmish.Bridge2
open CellScript.Core

[<EntryPoint>]
let main argv =


    let webApp =

        let init clientDispatch () =
            OuterMsg.No, Cmd.none

        let update clientDispatch msg model =
            clientDispatch (First FIA)
            printfn "%A" msg
            OuterMsg.No, Cmd.none

        let server =
            Bridge.mkServer Routed.endpoint init update
            |> Bridge.register First
            |> Bridge.run Suave2.server


        choose [
            path "/ping"   >=> Successful.OK "pong"
            server ]

    let suaveConfig =
        { defaultConfig with
            bindings   = [ HttpBinding.createSimple HTTP Routed.host Routed.port ] }


    startWebServer suaveConfig webApp

    0 // return an integer exit code
