module CellScript.Server.Tests.App

open System
open System.IO
open Microsoft.AspNetCore.Builder
open Microsoft.AspNetCore.Cors.Infrastructure
open Microsoft.AspNetCore.Hosting
open Microsoft.Extensions.Logging
open Microsoft.Extensions.DependencyInjection
open Giraffe
open NLog.Web
open Akkling
open Hyperion

// ---------------------------------
// Models
// ---------------------------------

type Message =
    {
        Text : string
    }

// ---------------------------------
// Views
// ---------------------------------

module Views =
    open GiraffeViewEngine

    let layout (content: XmlNode list) =
        html [] [
            head [] [
                title []  [ encodedText "CellScript.Server.Tests" ]
                link [ _rel  "stylesheet"
                       _type "text/css"
                       _href "/main.css" ]
            ]
            body [] content
        ]

    let partial () =
        h1 [] [ encodedText "CellScript.Server.Tests" ]

    let index (model : Message) =
        [
            partial()
            p [] [ encodedText model.Text ]
        ] |> layout

// ---------------------------------
// Web app
// ---------------------------------

let indexHandler (name : string) =
    let greetings = sprintf "Hello %s, from Giraffe!" name
    let model     = { Text = greetings }
    let view      = Views.index model
    htmlView view

let webApp =
    choose [
        GET >=>
            choose [
                route "/" >=> indexHandler "world"
                routef "/hello/%s" indexHandler
            ]
        setStatusCode 404 >=> text "Not Found" ]

// ---------------------------------
// Error handler
// ---------------------------------

let errorHandler (ex : Exception) (logger : ILogger) =
    logger.LogError(ex, "An unhandled exception has occurred while executing the request.")
    clearResponse >=> setStatusCode 500 >=> text ex.Message

// ---------------------------------
// Config and Main
// ---------------------------------

let configureCors (builder : CorsPolicyBuilder) =
    builder.WithOrigins("http://localhost:8080")
           .AllowAnyMethod()
           .AllowAnyHeader()
           |> ignore

let configureApp (app : IApplicationBuilder) =
    let env = app.ApplicationServices.GetService<IHostingEnvironment>()
    (match env.IsDevelopment() with
    | true  -> app.UseDeveloperExceptionPage()
    | false -> app.UseGiraffeErrorHandler errorHandler)
        .UseHttpsRedirection()
        .UseCors(configureCors)
        .UseStaticFiles()
        .UseGiraffe(webApp)

let configureServices (services : IServiceCollection) =
    services.AddCors()    |> ignore
    services.AddGiraffe() |> ignore

let configureLogging (builder : ILoggingBuilder) =
    builder.AddFilter(fun l -> l.Equals LogLevel.Error)
           .AddConsole()
           .AddDebug()
           .SetMinimumLevel(LogLevel.Trace) |> ignore


let buildWebhost logger args =
    let contentRoot = Directory.GetCurrentDirectory()
    let webRoot     = Path.Combine(contentRoot, "WebRoot")

    RemoteServer.run logger

    WebHostBuilder()
        .UseKestrel()
        .UseIISIntegration()
        .Configure(Action<IApplicationBuilder> configureApp)
        .ConfigureServices(configureServices)
        .ConfigureLogging(configureLogging)
        .UseNLog()
        .UseContentRoot(contentRoot)
        .UseWebRoot(webRoot)
        .Build()
        .Run()


[<EntryPoint>]
let main args =
    let logger = NLog.FSharp.Logger(NLog.Web.NLogBuilder.ConfigureNLog("nlog.config").GetCurrentClassLogger())
    try 
        logger.Debug "int main"
        buildWebhost logger args
    with ex ->
        //NLog: catch setup errors
        logger.Error "Stopped program because of exception %A" ex
        NLog.LogManager.Shutdown()
    0