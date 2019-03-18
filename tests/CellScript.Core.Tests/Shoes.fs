
namespace CellScript.Core.Tests
open FParsec

[<RequireQualifiedAccess>]
module String =
    let isInt32 s =
        match System.Int32.TryParse s with
        | true,_ -> true
        | false,_ -> false

[<RequireQualifiedAccess>]
module Float =
    let isHalf float =
        ((float * 2.).ToString() |> (String.isInt32))

[<RequireQualifiedAccess>]
module Shoes =

    let eur = [19..46]

    let isEur i = i>=19 && i<46

    let isEurFloat i = i>=19. && i<46.

    let isUK i = i>=3. && i<12. && Float.isHalf i

    let isUS i = i>=5. && i<12.5 && Float.isHalf i

    let sectionText sizes =
        if sizes |> Seq.forall isEurFloat then "Eur"
        elif sizes |> Seq.forall isUK then "UK"
        elif sizes |> Seq.forall isUS then "US"
        else failwithf "unknown size section %A" (List.ofSeq sizes)

    let isFraction i = i < 5

    let isManOfEur eur =
        eur > 38 && eur < 47

    let isWomanOfEur eur =
        eur > 34 && eur < 43

    let isManOfUK size =
        size > 5.5 && size < 12.

    let isWomanOfUK size =
        size > 2.5 && size < 9.5


[<RequireQualifiedAccess>]
module CellText =
    let private (!!) parser predicate (text: string) =
        let text = text.Trim()
        match run parser text with
        | Success (r,_,_) -> predicate r
        | Failure _ -> false

    let isEUR  = !! (pint32 .>> eof) Shoes.isEur
    let isEURRange  = !! (pint32 .>> pstring "-" .>> pint32 .>> eof) Shoes.isEur