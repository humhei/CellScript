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
open System.Text
open CsvHelper.Configuration
open System.Globalization


module CsvUtils =
    let private currentCultureEncoding =
        lazy
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance)
            Encoding.GetEncoding(config.Value.GetInt("CellScript.Core.CodePageIdentifier"))


    type CsvReader with 
        static member ReadRecordsI(csvFile: CsvFile, fRecord, ?fConfig) =
            let encoding = currentCultureEncoding.Value
            let config = new CsvConfiguration(CultureInfo.InvariantCulture, Encoding = encoding)
            let config =
                match fConfig with 
                | None -> config
                | Some fConfig -> fConfig config

            use stream = new FileStream(csvFile.Path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite)
            use streamReader = new StreamReader(stream, encoding)
            use csvReader = new CsvReader(streamReader, config)
            csvReader.Read() |> ignore
                
            csvReader.ReadHeader() |> ignore

            let mutable i = 0 

            let rec read accum =
                if csvReader.Read()
                then 
                    let record = fRecord i csvReader
                    i <- i + 1
                    read (record :: accum) 
                else accum

            read []
            |> List.rev

        static member ReadRecords(csvFile: CsvFile, fRecord, ?fConfig) =
            CsvReader.ReadRecordsI(csvFile, fRecord = (fun i reader -> fRecord reader), ?fConfig = fConfig)