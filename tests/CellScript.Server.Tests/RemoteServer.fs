module internal RemoteServer
open Akkling
open Shrimp.Akkling.Cluster.Intergraction
open CellScript.Core.Tests
open CellScript.Core
open CellScript.Core.Cluster
open CellScript.Core.Types
open Fake.IO.FileSystemOperators
open Fake.IO.Globbing.Operators
open Fake.IO
open Akka.Actor
open CellScript.Core.COM
open System.Threading

type Config = 
    { ServerPort: int }




let run (logger: NLog.FSharp.Logger) =

    let comClient = COMClient.create (COMRouted.port)

    let tryToSaveXlsToXlsx paths =
        let convertablePaths = 
            paths 
            |> List.filter(fun path ->
                let xlsxFile = Path.changeExtension ".xlsx" path
                not (File.exists xlsxFile)
        
            )
        
        let (result: Result<obj,string>) = 
            comClient.RemoteServer <? (COMMsg.SaveXlsToXlsx convertablePaths)
            |> Async.RunSynchronously
            
        match result with 
        | Result.Ok ok -> ok.ToString()
        | Result.Error error -> sprintf "ERRORS: %O" error

    let handleMsg (ctx: Actor<_>) (msg: obj) customModel = 
        try 
            match msg with  
            | :? OuterMsg as outerMsg ->
                match outerMsg with 
                | OuterMsg.TestString text ->
                    ctx.Sender() <! String text

                | OuterMsg.TestString2Params (text1, text2) ->
                    ctx.Sender() <! String text1

                | OuterMsg.TestTable table ->
                    ctx.Sender() <! table

                | OuterMsg.TestInnerMsg innerMsg ->
                    match innerMsg with 

                    | InnerMsg.TestExcelReference xlRef -> 
                        ctx.Sender() <! Record xlRef

                    | InnerMsg.TestString text ->
                        ctx.Sender() <! String text

                | OuterMsg.TestExcelVector vector ->
                    ctx.Sender() <! vector

                | OuterMsg.TestCommand commandCaller ->
                    ctx.Sender() <! CommandResponse.Ok

                | OuterMsg.TestStringArray stringArray ->
                    ctx.Sender() <! ExcelVector.Row (Array.map box stringArray)

                | OuterMsg.TestFunctionSheetReference (FunctionSheetReference caller) ->
                    let xlsxPath = caller.WorkbookPath
                    let dir = Path.getDirectory xlsxPath
                    let globbing = (dir </> "Orders" </> "**/*.xls")
                    let orders = 
                        !! globbing
                        |> List.ofSeq
                    let result = tryToSaveXlsToXlsx orders 

                    ctx.Sender() <! String result

                | OuterMsg.TestRecords (Records v) ->
                    ctx.Sender() <! (Records v)

                | OuterMsg.TestFailWith _ ->
                    failwith "Hello"
            | _ -> 
                ()

        with ex ->
            ctx.Sender() <! String (ex.Message)

    
    let system = Server.createAgent 9050 () handleMsg

    ()
