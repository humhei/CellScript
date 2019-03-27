namespace CellScript.Core
open System.Collections.Concurrent
module Registration =
    open System
    open System.Linq.Expressions
    open Extensions
    open Types


    [<AttributeUsage (AttributeTargets.Class ||| AttributeTargets.Struct,AllowMultiple=false)>]
    type AutoSerializableAttribute() = inherit Attribute()

    type CustomParamConversionAttribute()=
        inherit Attribute()

    [<RequireQualifiedAccess>]
    module CustomParamConversion =

        [<AllowNullLiteral>]
        type IConversion =
            /// fun obj[, ] -> 'T
            abstract member AsSingletonLambda : LambdaExpression
            /// fun obj[] -> 'T []
            abstract member AsArrayLambda : LambdaExpression
            abstract member IsRef: bool


        [<RequireQualifiedAccess>]
        module IConversion =
            let tryCreate (tp: Type) =
                match Type.tryGetAttribute<CustomParamConversionAttribute> tp with
                | Some _ ->
                    match Type.tryGetStaticMethod "Convert" tp with
                    | Some method ->
                        if tp.IsGenericType then
                            let innerType = tp.GetGenericArguments() |> Array.exactlyOne

                            let cast = Type.getBoxedBaseTypeConversion innerType

                            let method = method.MakeGenericMethod(innerType)

                            method.Invoke(null,[|cast|]) :?> IConversion

                        else
                            let conversion = method.Invoke(null,null)
                            conversion :?> IConversion
                        |> Some

                    | None ->
                        failwithf "Type %s implement CustomConversionAttribute should implement static method: Convert: unit -> _Conversion" tp.FullName
                | None ->
                    let fromCast cast =
                        { new IConversion with
                             member x.AsSingletonLambda = LambdaExpressionStatic.ofFun (fun input -> cast input)
                             member x.AsArrayLambda = LambdaExpressionStatic.ofFun (fun (input: obj[]) ->
                                input |> Array.map cast
                              )
                             member x.IsRef = false }

                    match tp.FullName with
                    | "System.Double" ->
                        Some (fromCast Cast.toDouble)
                    | "System.String" ->
                        Some (fromCast string)
                    | "System.Int32" ->
                        Some (fromCast Cast.toInt32)
                    | "System.Boolean" ->
                        Some (fromCast Cast.toBoolean)
                    | _ -> None

            let withArrayConversion (originType:Type) (originConversionGen: Type -> IConversion option) =
                if originType.IsArray then
                    let elementType = originType.GetElementType()
                    let baseConversionOp = originConversionGen elementType
                    match baseConversionOp with
                    | None -> None
                    | (Some conversion) -> Some (conversion.AsArrayLambda, conversion)
                else
                    let baseConversionOp = (originConversionGen originType)
                    match baseConversionOp with
                    | None -> None
                    | (Some conversion) -> Some (conversion.AsSingletonLambda, conversion)


        let private toArrayConversion originConversion (values) =
            values |> Array.map (fun value ->
                let value = value
                originConversion value
            )


        [<RequireQualifiedAccess>]
        type Conversion<'T> =
            private
                | Array2D of (obj [,] -> 'T)
                | Object of (obj -> 'T)
                | Ref of (ExcelReference -> 'T)


        with
            interface IConversion with
                member x.AsSingletonLambda =
                    match x with
                    | Array2D array2D -> LambdaExpressionStatic.ofFun(fun inputs -> array2D inputs)
                    | Object object ->
                        LambdaExpressionStatic.ofFun(fun (input: obj) -> object input)
                    | Ref ref ->
                        LambdaExpressionStatic.ofFun(fun (input: ExcelReference) -> ref input)

                member x.AsArrayLambda =

                    match x with
                    | Array2D array2D ->
                        LambdaExpressionStatic.ofFun (fun (values: obj []) ->
                            toArrayConversion (fun (baseValue: obj) -> array2D (baseValue :?> obj[,]) ) values
                        )
                    | Object object ->
                        LambdaExpressionStatic.ofFun (fun (values: obj []) ->
                            toArrayConversion object values
                        )
                    | Ref ref ->
                        LambdaExpressionStatic.ofFun (fun (values: ExcelReference []) ->
                            toArrayConversion ref values
                        )

                member x.IsRef =
                    match x with
                    | Array2D _ -> false
                    | Object _ -> false
                    | Ref _ -> true

        let private paramConversionMap = ConcurrentDictionary<Type, option<LambdaExpression * IConversion>>()

        let mapM mapping = function
            | Conversion.Array2D convert ->
                convert >> mapping
                |> Conversion.Array2D

            | Conversion.Object convert ->
                convert >> mapping
                |> Conversion.Object

            | Conversion.Ref convert ->
                convert >> mapping
                |> Conversion.Ref

        let array2D conversion =
            Conversion.Array2D conversion

        let object conversion =
            Conversion.Object conversion

        let ref (conversion: ExcelReference -> 'T) =
            Conversion.Ref conversion

        let register (originType: Type) (setParam: IConversion -> unit) =
            let paramConversionOp =
                paramConversionMap.GetOrAdd(originType, fun _ ->
                    IConversion.withArrayConversion originType IConversion.tryCreate
                )

            match paramConversionOp with
            | Some (lambda, conversion) ->
                setParam conversion
                lambda
            | None -> null

    type ICustomReturn =
        abstract member ReturnValue: unit -> obj[,]

    [<RequireQualifiedAccess>]
    module ICustomReturn =
        let returnValue (material: 'T when 'T :> ICustomReturn) =
            material.ReturnValue()


        let private returncConversionTypeMapping =
            new ConcurrentDictionary<Type,LambdaExpression>()

        let private asArray2d<'value>() =
            LambdaExpressionStatic.ofFun (fun (value: 'value) ->
                [[box value]]
                |> array2D
            )

        let register (paramType: Type) =
            returncConversionTypeMapping.GetOrAdd(paramType,fun _ ->
                paramType.GetInterfaces()
                |> Array.tryFind (fun it -> it.FullName = typeof<ICustomReturn>.FullName)
                |> function
                    | Some _ ->
                        LambdaExpressionStatic.ofFun (fun (value: obj) ->
                            match value with
                            | :? ICustomReturn as iReturn ->
                                iReturn.ReturnValue()
                            | _ -> null
                        )
                    | _ ->
                        match paramType.FullName with
                        | "System.Int32" -> asArray2d<int>()
                        | "System.String" -> asArray2d<string>()
                        | "System.Boolean" -> asArray2d<bool>()
                        | "System.Double" -> asArray2d<float>()
                        | _ ->
                            null
            )