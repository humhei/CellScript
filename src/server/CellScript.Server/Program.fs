﻿// Learn more about F# at http://fsharp.org

open Suave
open Fable.Remoting
open CellScript.Shared
open Suave.Filters
open Suave.Operators
open Fable.Remoting.Suave
open Fable.Remoting.Server
open CellScript.Server.UDF

[<EntryPoint>]
let main argv =
    let webApp = 
        let webPart =
            Remoting.createApi()
            |> Remoting.fromValue cellScriptUDFApi
            |> Remoting.withRouteBuilder Route.routeBuilder
            |> Remoting.buildWebPart

        choose [
            path "/ping"   >=> Successful.OK "pong"
            webPart ]

    let suaveConfig =
        { defaultConfig with
            bindings   = [ HttpBinding.createSimple HTTP Route.host Route.port ] }

    startWebServer suaveConfig webApp

    0 // return an integer exit code