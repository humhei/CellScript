namespace CellScript.Core
open Deedle
open System
open System.Reflection
open System.Linq.Expressions
open OfficeOpenXml

module Extensions =


    [<RequireQualifiedAccess>]
    type LambdaExpressionStatic =
        static member ofFun(exp: Expression<Func<_,_>>) =
            exp :> LambdaExpression


    [<RequireQualifiedAccess>]
    module LambdaExpression =

        /// retrive anonymous in module
        let ofRuntimeMemberInfo(memberInfo: MemberInfo) =
            let memberInfoType = memberInfo.GetType()

            if memberInfoType.FullName <> "System.RuntimeType" then failwithf "member info named %s is not runtime type" memberInfoType.Name

            let getValue propertyName = memberInfo.GetType().GetProperty(propertyName).GetValue(memberInfo)

            let baseType = getValue "BaseType" :?> Type
            assert (baseType.Name.StartsWith "FSharpFunc`")


            let methodInfo = getValue "DeclaredMethods" :?> seq<MethodInfo> |> Seq.exactlyOne

            let instance =
                let constructor = getValue "DeclaredConstructors" :?> seq<ConstructorInfo> |> Seq.exactlyOne

                let boxedFunction = constructor.Invoke(null)

                LambdaExpression.Constant boxedFunction

            let param =
                methodInfo.GetParameters()
                |> Array.map (fun pi -> Expression.Parameter(pi.ParameterType,pi.Name))

            let paramExprs = param |> Seq.cast<Expression>

            Expression.Lambda(Expression.Call(instance,methodInfo,paramExprs),param)




    [<RequireQualifiedAccess>]
    module Cast =
        let toBoolean (value: obj) =
            match value with
            | :? bool as v -> v
            | _ -> failwithf "Cannot cast value %A to boolean" value

        let toInt32 (value: obj) =
            let text = value.ToString()
            Int32.Parse text

        let toDouble (value: obj) =
            let text = value.ToString()
            Double.Parse text

    [<RequireQualifiedAccess>]
    module Type =

        let tryGetAttribute<'Attribute> (tp: Type) =
            tp.GetCustomAttributes(false)
            |> Seq.tryFind (fun attr ->
                let t1 =  attr.GetType()
                let t2 =typeof<'Attribute>
                t1 = t2
            )

        let tryGetStaticMethod methodName (tp: Type) =
            try
                tp.GetMethod(methodName, BindingFlags.Static|||BindingFlags.Public) |> Some
            with _ -> None

        let getBoxedBaseTypeConversion (tp: Type) =
            match tp.FullName with
            | "System.String" -> string |> box
            | "System.Boolean" -> Cast.toBoolean |> box
            | "System.Int32" -> Cast.toInt32 |> box
            | "System.Object" -> box |> box
            | _ -> failwith "not implemented"

        let isSomeModuleOrNamespace (type1: Type) (type2: Type) =
            let prefixName (tp: Type) = 
                let name = tp.Name
                let fullName = tp.FullName
                fullName.Remove(fullName.Length - name.Length - 1)
            prefixName type1 = prefixName type2

    [<RequireQualifiedAccess>]
    module String =
        open System
        let ofCharList chars = chars |> List.toArray |> String

        let equalIgnoreCaseAndEdgeSpace (text1: string) (text2: string) =
            let trimedText1 = text1.Trim()
            let trimedText2 = text2.Trim()

            String.Equals(trimedText1,trimedText2,StringComparison.InvariantCultureIgnoreCase)

        let leftOf (pattern: string) (input: string) =
            let index = input.IndexOf(pattern)
            input.Substring(0,index)

        let rightOf (pattern: string) (input: string) =
            let index = input.IndexOf(pattern)
            input.Substring(index + 1)

            

    [<RequireQualifiedAccess>]
    module Array2D =

        let toSeqs (input: 'a[,]) =
            let l1 = input.GetLowerBound(0)
            let u1 = input.GetUpperBound(0)
            seq {
                for i = l1 to u1 do
                    yield input.[i,*] :> seq<'a>
            }

        let transpose (input: 'a[,]) =
            let l1 = input.GetLowerBound(1)
            let u1 = input.GetUpperBound(1)
            seq {
                for i = l1 to u1 do
                    yield input.[i,*]
            }
            |> array2D

    [<RequireQualifiedAccess>]
    module Frame =
        let mapValuesString mapping frame =
            let mapping raw =
                mapping (raw.ToString())
            Frame.mapValues mapping frame


    [<RequireQualifiedAccess>]
    module ExcelRangeBase =
        let ofArray2D (p: ExcelPackage) (values: obj[,]) =
            let baseArray = Array2D.toSeqs values |> Seq.map Array.ofSeq
            let ws = p.Workbook.Worksheets.Add("Sheet1")
            ws.Cells.["A1"].LoadFromArrays(baseArray)