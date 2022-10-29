namespace CellScript.Core
open Deedle
open Akka.Util
open System
open Shrimp.FSharp.Plus
open System.IO
open OfficeOpenXml
open Newtonsoft.Json
open System.Runtime.CompilerServices
open FParsec
open FParsec.CharParsers
open System.Collections.Generic
open CsvHelper
open Akka.Configuration
open System.Text
open Shrimp.Akkling.Cluster.Intergraction.Configuration
open Shrimp.Akkling.Cluster.Intergraction


[<AutoOpen>]
module private _Config =
    type private AssemblyFinder = AssemblyFinder
    let config = 
        lazy
            ConfigurationFactory
                .FromResource<AssemblyFinder>("CellScript.Core.reference.conf")
            |> Configuration.fallBackByApplicationConf
