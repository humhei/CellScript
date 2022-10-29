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
        static member ReadRecords(csvFile: CsvFile, fRecord) =
            let encoding = currentCultureEncoding.Value
            let config = new CsvConfiguration(CultureInfo.InvariantCulture, Encoding = encoding)
            use stream = new FileStream(csvFile.Path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite)
            use streamReader = new StreamReader(stream, encoding)
            use csvReader = new CsvReader(streamReader, config)
            csvReader.Read() |> ignore
                
            csvReader.ReadHeader() |> ignore

            let rec read accum =
                if csvReader.Read()
                then 
                    let record = fRecord csvReader
                            
                    read (record :: accum) 
                else accum

            read []
            |> List.rev

                