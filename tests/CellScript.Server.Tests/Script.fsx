open System
open System.Threading

let a = 1
let n = TimeSpan.FromSeconds(2.)
Thread.Sleep(2000)
n.Milliseconds
//n.TotalMilliseconds