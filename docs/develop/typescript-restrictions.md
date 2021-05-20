---
title: Officeスクリプトにおける TypeScript の制限
description: Officeスクリプト コード エディターで使用される TypeScript コンパイラおよびリンターの詳細。
ms.date: 02/05/2021
localization_priority: Normal
ms.openlocfilehash: a4198e0e56224ac5da89e89c43c8d2f3ef44d6d7
ms.sourcegitcommit: 4687693f02fc90a57ba30c461f35046e02e6f5fb
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 05/19/2021
ms.locfileid: "52545020"
---
# <a name="typescript-restrictions-in-office-scripts"></a><span data-ttu-id="91402-103">Officeスクリプトにおける TypeScript の制限</span><span class="sxs-lookup"><span data-stu-id="91402-103">TypeScript restrictions in Office Scripts</span></span>

<span data-ttu-id="91402-104">Officeスクリプトは、タイプスクリプト言語を使用します。</span><span class="sxs-lookup"><span data-stu-id="91402-104">Office Scripts use the TypeScript language.</span></span> <span data-ttu-id="91402-105">ほとんどの場合、TypeScript または JavaScript コードはOfficeスクリプトで動作します。</span><span class="sxs-lookup"><span data-stu-id="91402-105">For the most part, any TypeScript or JavaScript code will work in Office Scripts.</span></span> <span data-ttu-id="91402-106">ただし、スクリプトが一貫して、Excelブックで意図したとおりに動作するように、コード エディターによって適用される制限がいくつかあります。</span><span class="sxs-lookup"><span data-stu-id="91402-106">However, there are a few restrictions enforced by the Code Editor to ensure your script works consistently and as intended with your Excel workbook.</span></span>

## <a name="no-any-type-in-office-scripts"></a><span data-ttu-id="91402-107">Officeスクリプトに 「任意」 タイプがありません</span><span class="sxs-lookup"><span data-stu-id="91402-107">No 'any' type in Office Scripts</span></span>

<span data-ttu-id="91402-108">[TypeScript](https://www.typescriptlang.org/docs/handbook/typescript-in-5-minutes.html)では型を推論できるため、型の書き込みはオプションです。</span><span class="sxs-lookup"><span data-stu-id="91402-108">Writing [types](https://www.typescriptlang.org/docs/handbook/typescript-in-5-minutes.html) is optional in TypeScript, because the types can be inferred.</span></span> <span data-ttu-id="91402-109">ただし、Officeスクリプトでは、変数を[任意の型](https://www.typescriptlang.org/docs/handbook/basic-types.html#any)にすることはできません。</span><span class="sxs-lookup"><span data-stu-id="91402-109">However, Office Scripts requires that a variable can't be of [type any](https://www.typescriptlang.org/docs/handbook/basic-types.html#any).</span></span> <span data-ttu-id="91402-110">`any`Office スクリプトでは、明示的なスクリプトと暗黙的な両方を使用できません。</span><span class="sxs-lookup"><span data-stu-id="91402-110">Both explicit and implicit `any` are not allowed in Office Scripts.</span></span> <span data-ttu-id="91402-111">これらのケースはエラーとして報告されます。</span><span class="sxs-lookup"><span data-stu-id="91402-111">These cases are reported as errors.</span></span>

### <a name="explicit-any"></a><span data-ttu-id="91402-112">暁 `any`</span><span class="sxs-lookup"><span data-stu-id="91402-112">Explicit `any`</span></span>

<span data-ttu-id="91402-113">`any`Officeスクリプト (つまり) で変数を型として明示的に宣言することはできません `let someVariable: any;` 。</span><span class="sxs-lookup"><span data-stu-id="91402-113">You cannot explicitly declare a variable to be of type `any` in Office Scripts (that is, `let someVariable: any;`).</span></span> <span data-ttu-id="91402-114">`any`この型は、Excelで処理されるときに問題を引き起こします。</span><span class="sxs-lookup"><span data-stu-id="91402-114">The `any` type causes issues when processed by Excel.</span></span> <span data-ttu-id="91402-115">たとえば、 の `Range` 値が `string` `number` 、 、または であることを知る必要があります `boolean` 。</span><span class="sxs-lookup"><span data-stu-id="91402-115">For example, a `Range` needs to know that a value is a `string`, `number`, or `boolean`.</span></span> <span data-ttu-id="91402-116">スクリプト内で型として明示的に定義されている変数がある場合、コンパイル エラー (スクリプトの実行前にエラー) `any` が表示されます。</span><span class="sxs-lookup"><span data-stu-id="91402-116">You will receive a compile-time error (an error prior to running the script) if any variable is explicitly defined as the `any` type in the script.</span></span>

:::image type="content" source="../images/explicit-any-editor-message.png" alt-text="コード エディターのホバー テキスト内の明示的な 'any' メッセージ":::

:::image type="content" source="../images/explicit-any-error-message.png" alt-text="コンソール ウィンドウでの明示的な 'any' エラー":::

<span data-ttu-id="91402-119">前のスクリーンショットでは `[5, 16] Explicit Any is not allowed` 、行#5、列#16が型を定義していることを示 `any` しています。</span><span class="sxs-lookup"><span data-stu-id="91402-119">In the previous screenshot `[5, 16] Explicit Any is not allowed` indicates that line #5, column #16 defines `any` type.</span></span> <span data-ttu-id="91402-120">これにより、エラーを見つけることができます。</span><span class="sxs-lookup"><span data-stu-id="91402-120">This helps you locate the error.</span></span>

<span data-ttu-id="91402-121">この問題を回避するには、常に変数の型を定義します。</span><span class="sxs-lookup"><span data-stu-id="91402-121">To get around this issue, always define the type of the variable.</span></span> <span data-ttu-id="91402-122">変数の型が不明な場合は、 [共用体型](https://www.typescriptlang.org/docs/handbook/unions-and-intersections.html)を使用できます。</span><span class="sxs-lookup"><span data-stu-id="91402-122">If you are uncertain about the type of a variable, you can use a [union type](https://www.typescriptlang.org/docs/handbook/unions-and-intersections.html).</span></span> <span data-ttu-id="91402-123">これは、値を保持する変数 `Range` (型 `string` `number` 、、または `boolean` 値の型 `Range` がそれらの和集合である) `string | number | boolean` に役立ちます。</span><span class="sxs-lookup"><span data-stu-id="91402-123">This can be useful for variables that hold `Range` values, which can be of type `string`, `number`, or `boolean` (the type for `Range` values is a union of those: `string | number | boolean`).</span></span>

### <a name="implicit-any"></a><span data-ttu-id="91402-124">暗黙の `any`</span><span class="sxs-lookup"><span data-stu-id="91402-124">Implicit `any`</span></span>

<span data-ttu-id="91402-125">TypeScript 変数型は [、暗黙的に](https://www.typescriptlang.org/docs/handbook/type-inference.html) 定義できます。</span><span class="sxs-lookup"><span data-stu-id="91402-125">TypeScript variable types can be [implicitly](https://www.typescriptlang.org/docs/handbook/type-inference.html) defined.</span></span> <span data-ttu-id="91402-126">TypeScript コンパイラが変数の型を判断できない場合 (型が明示的に定義されていないか、型の推論が不可能な場合)、暗黙的なエラー `any` が発生します。</span><span class="sxs-lookup"><span data-stu-id="91402-126">If the TypeScript compiler is unable to determine the type of a variable (either because type is not defined explicitly or type inference isn't possible), then it's an implicit `any` and you will receive a compilation-time error.</span></span>

<span data-ttu-id="91402-127">暗黙的な場合の最も一般的なケース `any` は、変数宣言 (など `let value;` ) です。</span><span class="sxs-lookup"><span data-stu-id="91402-127">The most common case on any implicit `any` is in a variable declaration, such as `let value;`.</span></span> <span data-ttu-id="91402-128">これを回避するには、次の 2 つの方法があります。</span><span class="sxs-lookup"><span data-stu-id="91402-128">There are two ways to avoid this:</span></span>

* <span data-ttu-id="91402-129">暗黙的に識別できる型 ( または ) に変数を割り当てます `let value = 5;` `let value = workbook.getWorksheet();` 。</span><span class="sxs-lookup"><span data-stu-id="91402-129">Assign the variable to an implicitly identifiable type (`let value = 5;` or `let value = workbook.getWorksheet();`).</span></span>
* <span data-ttu-id="91402-130">変数を明示的に入力します ( `let value: number;` )</span><span class="sxs-lookup"><span data-stu-id="91402-130">Explicitly type the variable (`let value: number;`)</span></span>

## <a name="no-inheriting-office-script-classes-or-interfaces"></a><span data-ttu-id="91402-131">継承Officeスクリプトクラスまたはインタフェースなし</span><span class="sxs-lookup"><span data-stu-id="91402-131">No inheriting Office Script classes or interfaces</span></span>

<span data-ttu-id="91402-132">Office スクリプトで作成されたクラスおよびインターフェイスは、スクリプトクラスまたはOfficeスクリプト クラスまたはインターフェイスを[拡張または実装](https://www.typescriptlang.org/docs/handbook/classes.html#inheritance)できません。</span><span class="sxs-lookup"><span data-stu-id="91402-132">Classes and interfaces that are created in your Office Script cannot [extend or implement](https://www.typescriptlang.org/docs/handbook/classes.html#inheritance) Office Scripts classes or interfaces.</span></span> <span data-ttu-id="91402-133">つまり、名前空間内のサブ `ExcelScript` クラスまたはサブインターフェイスを持つものはありません。</span><span class="sxs-lookup"><span data-stu-id="91402-133">In other words, nothing in the `ExcelScript` namespace can have subclasses or subinterfaces.</span></span>

## <a name="incompatible-typescript-functions"></a><span data-ttu-id="91402-134">互換性のないタイプスクリプト関数</span><span class="sxs-lookup"><span data-stu-id="91402-134">Incompatible TypeScript functions</span></span>

<span data-ttu-id="91402-135">Officeスクリプト API は、次の場合は使用できません。</span><span class="sxs-lookup"><span data-stu-id="91402-135">Office Scripts APIs cannot be used in the following:</span></span>

* [<span data-ttu-id="91402-136">ジェネレータ関数</span><span class="sxs-lookup"><span data-stu-id="91402-136">Generator functions</span></span>](https://developer.mozilla.org/docs/Web/JavaScript/Guide/Iterators_and_Generators#generator_functions)
* [<span data-ttu-id="91402-137">配列.ソート</span><span class="sxs-lookup"><span data-stu-id="91402-137">Array.sort</span></span>](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Array/sort)

## <a name="eval-is-not-supported"></a><span data-ttu-id="91402-138">`eval` はサポートされていません</span><span class="sxs-lookup"><span data-stu-id="91402-138">`eval` is not supported</span></span>

<span data-ttu-id="91402-139">セキュリティ上の理由から、JavaScript [eval 関数](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/eval) はサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="91402-139">The JavaScript [eval function](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/eval) is not supported for security reasons.</span></span>

## <a name="restricted-identifers"></a><span data-ttu-id="91402-140">制限付き識別コード</span><span class="sxs-lookup"><span data-stu-id="91402-140">Restricted identifers</span></span>

<span data-ttu-id="91402-141">次の単語は、スクリプトの識別子として使用できません。</span><span class="sxs-lookup"><span data-stu-id="91402-141">The following words can't be used as identifiers in a script.</span></span> <span data-ttu-id="91402-142">彼らは予約された用語です。</span><span class="sxs-lookup"><span data-stu-id="91402-142">They are reserved terms.</span></span>

* `Excel`
* `ExcelScript`
* `console`

## <a name="only-arrow-functions-in-array-callbacks"></a><span data-ttu-id="91402-143">配列コールバックの矢印関数のみ</span><span class="sxs-lookup"><span data-stu-id="91402-143">Only arrow functions in array callbacks</span></span>

<span data-ttu-id="91402-144">スクリプトは[、Array](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Array)メソッドにコールバック引数を指定する場合にのみ[、矢印関数](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Functions/Arrow_functions)を使用できます。</span><span class="sxs-lookup"><span data-stu-id="91402-144">Your scripts can only use [arrow functions](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Functions/Arrow_functions) when providing callback arguments for [Array](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Array) methods.</span></span> <span data-ttu-id="91402-145">これらのメソッドに対して、いかなる種類の識別子や「伝統的な」関数も渡すことはできません。</span><span class="sxs-lookup"><span data-stu-id="91402-145">You cannot pass any sort of identifier or "traditional" function to these methods.</span></span>

```TypeScript
const myArray = [1, 2, 3, 4, 5, 6];
let filteredArray = myArray.filter((x) => {
  return x % 2 === 0;
});
/*
  The following code generates a compiler error in the Office Scripts Code Editor.
  filteredArray = myArray.filter(function (x) {
    return x % 2 === 0;
  });
*/
```

## <a name="performance-warnings"></a><span data-ttu-id="91402-146">パフォーマンスの警告</span><span class="sxs-lookup"><span data-stu-id="91402-146">Performance warnings</span></span>

<span data-ttu-id="91402-147">スクリプトにパフォーマンスの問題がある場合、コード エディタの [リンター](https://wikipedia.org/wiki/Lint_(software)) は警告を表示します。</span><span class="sxs-lookup"><span data-stu-id="91402-147">The Code Editor's [linter](https://wikipedia.org/wiki/Lint_(software)) gives warnings if the script might have performance issues.</span></span> <span data-ttu-id="91402-148">ケースとその回避方法については[、「Office スクリプトのパフォーマンスを向上させる](web-client-performance.md)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="91402-148">The cases and how to work around them are documented in [Improve the performance of your Office Scripts](web-client-performance.md).</span></span>

## <a name="external-api-calls"></a><span data-ttu-id="91402-149">外部 API 呼び出し</span><span class="sxs-lookup"><span data-stu-id="91402-149">External API calls</span></span>

<span data-ttu-id="91402-150">詳細については[、「Officeスクリプト」の外部 API コールサポート](external-calls.md)を参照してください。</span><span class="sxs-lookup"><span data-stu-id="91402-150">See [External API call support in Office Scripts](external-calls.md) for more information.</span></span>

## <a name="see-also"></a><span data-ttu-id="91402-151">関連項目</span><span class="sxs-lookup"><span data-stu-id="91402-151">See also</span></span>

* [<span data-ttu-id="91402-152">Excel on the web での Office スクリプトのスクリプトの基本事項</span><span class="sxs-lookup"><span data-stu-id="91402-152">Scripting fundamentals for Office Scripts in Excel on the web</span></span>](scripting-fundamentals.md)
* [<span data-ttu-id="91402-153">Officeスクリプトのパフォーマンスを向上させる</span><span class="sxs-lookup"><span data-stu-id="91402-153">Improve the performance of your Office Scripts</span></span>](web-client-performance.md)
