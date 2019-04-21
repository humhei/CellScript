module Tests
open CellScript.Core.Tests
open Akkling
open CellScript.Core.Remote
open Expecto
open CellScript.Core.Types
open System.IO
open System.Reflection
open Fake.IO.FileSystemOperators
open CellScript.Client
open CellScript.Core

let dir = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location)

let book1 = dir </> "datas/book1.xlsx"

let equal actual expected = Expect.equal expected actual "passed"

let equalByResponse expected response = 
    let array2D: obj[,] = unbox response 
    Expect.equal array2D.[0,0] expected "passed"

let udfTests =
    testList "udf tests" [
        testCase "get excel functions" (fun _ ->
            /// ensure no exceptions
            let excelFunctions = Registration.excelFunctions<OuterMsg>()
            ()
        )

        testCase "UDF tests" (fun _ ->
            let client = Client.create()
            let actor = client.RemoteServer
            let response =
                actor <? OuterMsg.TestString "Hello world"
                |> Async.RunSynchronously

            equalByResponse "Hello world" response
        )
    ]

let evalTests =
    testList "eval tests" [
    
        testCase "string test" (fun _ ->
            let client = Client.create()
            let actor = client.RemoteServer

            let tempData = 
                { ColumnFirst = 1
                  RowFirst = 1
                  ColumnLast = 1
                  RowLast = 1
                  WorkbookPath = book1
                  SheetName = "Script"
                  Content = array2D[["Hello"]]
                }

            let response =
                actor <? FcsMsg.Eval (tempData, "fun a -> a + \"8\"")
                |> Async.RunSynchronously

            equalByResponse "Hello8" response

        )

        testCase "table" (fun _ ->
            let client = Client.create()
            let actor = client.RemoteServer

            let tempData = 
                { ColumnFirst = 1
                  RowFirst = 1
                  ColumnLast = 1
                  RowLast = 1
                  WorkbookPath = book1
                  SheetName = "Script"
                  Content = (Table.TempData :> IToArray2D).ToArray2D()
                }

            let response =
                actor <? FcsMsg.Eval (tempData, 
                    """
#load "..\Packages.fsx"   
open CellScript.Core
fun (table: Table) -> table
                    """)
                |> Async.RunSynchronously

            equalByResponse "Hello8" response

        )
    
    ]