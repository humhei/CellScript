namespace CellScript.Server
open System.Threading.Tasks
open Suave
open Fable.Remoting
open CellScript.Shared
open Suave.Filters
open Suave.Operators
open CellScript.Shared
open Fable.Remoting.Suave
open Fable.Remoting.Server
open Fable.Remoting.Suave.SuaveUtil
open CellScript.Server
open System
open Newtonsoft.Json

module Main =


    let runCellScriptServer (config: Config) =

        let webApp =
            let cellScriptApi: ICellScriptApi =
                { eval =
                    fun input code  ->
                        async {return Fcs.eval config input code}
                  test = fun () -> async { return "Hello world from Fable.Remoting"}
                  go =
                    fun s ->
                        printf "%A" s
                        async {return "Hello"}
                }

            let webPart =
                Remoting.createApi()
                |> Remoting.fromValue cellScriptApi
                |> Remoting.withRouteBuilder Route.routeBuilder
                |> Remoting.buildWebPart

            choose [
                path "/ping"   >=> Successful.OK "pong"
                webPart ]

        let suaveConfig =
            { defaultConfig with
                bindings   = [ HttpBinding.createSimple HTTP Route.host Route.port ] }

        startWebServer suaveConfig webApp