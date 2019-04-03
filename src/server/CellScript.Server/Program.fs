open System.Runtime.Serialization
open Newtonsoft.Json
open System

// Learn more about F# at http://fsharp.org
#nowarn "0064"

open CellScript.Core
open Akkling
open System.Threading

let testTable (table: Table) =
    table



type Ext = Ext
    with
        static member Bar (ext : Ext, flt : float) = 
            array2D [[box flt]]

        static member Bar (ext : Ext, flt : string) = 
            array2D [[box flt]]

        static member Bar (ext : Ext, flt : string []) = 
            array2D [Array.map box flt]

        static member Bar (ext : Ext, flt : int) = 
            array2D [[box flt]]

        static member Bar (ext : Ext, flt : Table) = 
            flt.AsArray2D

        static member Bar (ext : Ext, flt : ExcelArray) = 
            flt.AsArray2D

        static member inline Bar (ext : Ext, flt : ^a) = 
            (^a : (member AsArray2D : obj [,]) (flt))


let inline bar (x : ^a) =
    ((^b or ^a) : (static member Bar : ^b * ^a -> obj[,]) (Ext, x))


let ext = Ext

let inline (<!) (actor: IActorRef<_>) (msg: 'Message) =
    let array2DMsg = bar msg
    actor <! array2DMsg


[<EntryPoint>]
let main argv =
    Routed.remote <- Remote.Server (fun ctx msg ->
        match msg with 
        | OuterMsg.TestString text ->
            ctx.Sender() <! text
            ignored ()
        | OuterMsg.TestString2Params (text1, text2) ->
            ctx.Sender() <! text1
            ignored ()
        | OuterMsg.TestTable table ->
            ctx.Sender() <! testTable table
            ignored ()
        | OuterMsg.TestArray texts ->
            ctx.Sender() <! texts
            ignored ()
        | OuterMsg.TestInnerMsg innerMsg ->
            ctx.Sender() <! "TestInnerMsg"
            ignored ()

    )

    let server = System.create "server" <| Configuration.parse Routed.serverConfig
    let manualSet = new ManualResetEventSlim()
    manualSet.Wait()
    0 // return an integer exit code
