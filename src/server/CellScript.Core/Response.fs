namespace CellScript.Core
open Extensions
open System
type IToArray2D =
    abstract member ToArray2D: unit -> IConvertible [,]

//type CommandResponse =
//    | Ok = 0


//type Records<'T> = Records of 'T list
//with 
//    interface IToArray2D with 
//        member x.ToArray2D() =
//            let result = 
//                let (Records recordValues) = x 
//                let tp = typeof<'T>
//                tp.GetProperties()
//                |> Array.map (fun prop ->
//                    let values = recordValues |> List.map (prop.GetValue >> string)
//                    prop.Name :: values
//                )
//                |> array2D
//                |> Array2D.transpose
//                |> Array2D.map box
//            result

//type Record<'T> = Record of 'T
//with 
//    interface IToArray2D with 
//        member x.ToArray2D() =
//            let (Record recordValue) = x 
//            (Records [recordValue] :> IToArray2D).ToArray2D()

//type Array2D = Array2D of obj [,] 
//with 
//    interface IToArray2D with 
//        member x.ToArray2D() =
//            let (Array2D array2D) = x 
//            array2D