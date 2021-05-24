---
title: スクリプトの TypeScript の制限Officeスクリプト
description: スクリプト コード エディターで使用される TypeScript コンパイラと linter のOfficeします。
ms.date: 02/05/2021
localization_priority: Normal
ms.openlocfilehash: a4198e0e56224ac5da89e89c43c8d2f3ef44d6d7
ms.sourcegitcommit: 4687693f02fc90a57ba30c461f35046e02e6f5fb
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 05/19/2021
ms.locfileid: "52545020"
---
# <a name="typescript-restrictions-in-office-scripts"></a><span data-ttu-id="212bc-103">スクリプトの TypeScript の制限Officeスクリプト</span><span class="sxs-lookup"><span data-stu-id="212bc-103">TypeScript restrictions in Office Scripts</span></span>

<span data-ttu-id="212bc-104">Officeスクリプトは TypeScript 言語を使用します。</span><span class="sxs-lookup"><span data-stu-id="212bc-104">Office Scripts use the TypeScript language.</span></span> <span data-ttu-id="212bc-105">ほとんどの場合、すべての TypeScript または JavaScript コードは、スクリプトのOfficeされます。</span><span class="sxs-lookup"><span data-stu-id="212bc-105">For the most part, any TypeScript or JavaScript code will work in Office Scripts.</span></span> <span data-ttu-id="212bc-106">ただし、コード エディターによって、スクリプトが一貫して動作し、ブックの目的に合Excelがあります。</span><span class="sxs-lookup"><span data-stu-id="212bc-106">However, there are a few restrictions enforced by the Code Editor to ensure your script works consistently and as intended with your Excel workbook.</span></span>

## <a name="no-any-type-in-office-scripts"></a><span data-ttu-id="212bc-107">スクリプトに 'any' 型Officeはありません</span><span class="sxs-lookup"><span data-stu-id="212bc-107">No 'any' type in Office Scripts</span></span>

<span data-ttu-id="212bc-108">[TypeScript](https://www.typescriptlang.org/docs/handbook/typescript-in-5-minutes.html)では、型を推論できるので、書き込み型は省略可能です。</span><span class="sxs-lookup"><span data-stu-id="212bc-108">Writing [types](https://www.typescriptlang.org/docs/handbook/typescript-in-5-minutes.html) is optional in TypeScript, because the types can be inferred.</span></span> <span data-ttu-id="212bc-109">ただし、Officeスクリプトでは、変数を任意の型に[できない必要があります](https://www.typescriptlang.org/docs/handbook/basic-types.html#any)。</span><span class="sxs-lookup"><span data-stu-id="212bc-109">However, Office Scripts requires that a variable can't be of [type any](https://www.typescriptlang.org/docs/handbook/basic-types.html#any).</span></span> <span data-ttu-id="212bc-110">明示的および暗黙的の両方 `any` は、スクリプトでOfficeされません。</span><span class="sxs-lookup"><span data-stu-id="212bc-110">Both explicit and implicit `any` are not allowed in Office Scripts.</span></span> <span data-ttu-id="212bc-111">これらのケースはエラーとして報告されます。</span><span class="sxs-lookup"><span data-stu-id="212bc-111">These cases are reported as errors.</span></span>

### <a name="explicit-any"></a><span data-ttu-id="212bc-112">明示的 `any`</span><span class="sxs-lookup"><span data-stu-id="212bc-112">Explicit `any`</span></span>

<span data-ttu-id="212bc-113">変数をスクリプト (つまり) の型Office `any` 明示的に宣言することはできません `let someVariable: any;` 。</span><span class="sxs-lookup"><span data-stu-id="212bc-113">You cannot explicitly declare a variable to be of type `any` in Office Scripts (that is, `let someVariable: any;`).</span></span> <span data-ttu-id="212bc-114">この `any` 型は、ユーザーが処理した場合に問題Excel。</span><span class="sxs-lookup"><span data-stu-id="212bc-114">The `any` type causes issues when processed by Excel.</span></span> <span data-ttu-id="212bc-115">たとえば、値が 、 、 または である必要 `Range` `string` `number` があります `boolean` 。</span><span class="sxs-lookup"><span data-stu-id="212bc-115">For example, a `Range` needs to know that a value is a `string`, `number`, or `boolean`.</span></span> <span data-ttu-id="212bc-116">スクリプト内の型として変数が明示的に定義されている場合は、コンパイル時エラー (スクリプトを実行する前のエラー) `any` が表示されます。</span><span class="sxs-lookup"><span data-stu-id="212bc-116">You will receive a compile-time error (an error prior to running the script) if any variable is explicitly defined as the `any` type in the script.</span></span>

:::image type="content" source="../images/explicit-any-editor-message.png" alt-text="コード エディターのホバー テキストの明示的な 'any' メッセージ":::

:::image type="content" source="../images/explicit-any-error-message.png" alt-text="コンソール ウィンドウの明示的な 'any' エラー":::

<span data-ttu-id="212bc-119">前のスクリーンショットでは `[5, 16] Explicit Any is not allowed` 、行の種類を#5列#16示 `any` しています。</span><span class="sxs-lookup"><span data-stu-id="212bc-119">In the previous screenshot `[5, 16] Explicit Any is not allowed` indicates that line #5, column #16 defines `any` type.</span></span> <span data-ttu-id="212bc-120">これにより、エラーを見つけるのに役立ちます。</span><span class="sxs-lookup"><span data-stu-id="212bc-120">This helps you locate the error.</span></span>

<span data-ttu-id="212bc-121">この問題を回避するには、常に変数の種類を定義します。</span><span class="sxs-lookup"><span data-stu-id="212bc-121">To get around this issue, always define the type of the variable.</span></span> <span data-ttu-id="212bc-122">変数の種類が不明な場合は、共用体の型を [使用できます](https://www.typescriptlang.org/docs/handbook/unions-and-intersections.html)。</span><span class="sxs-lookup"><span data-stu-id="212bc-122">If you are uncertain about the type of a variable, you can use a [union type](https://www.typescriptlang.org/docs/handbook/unions-and-intersections.html).</span></span> <span data-ttu-id="212bc-123">これは、型 、または (値の型は、それらの共用体です) の値を保持する変数 `Range` `string` `number` `boolean` `Range` に役立ちます `string | number | boolean` 。</span><span class="sxs-lookup"><span data-stu-id="212bc-123">This can be useful for variables that hold `Range` values, which can be of type `string`, `number`, or `boolean` (the type for `Range` values is a union of those: `string | number | boolean`).</span></span>

### <a name="implicit-any"></a><span data-ttu-id="212bc-124">暗黙的 `any`</span><span class="sxs-lookup"><span data-stu-id="212bc-124">Implicit `any`</span></span>

<span data-ttu-id="212bc-125">TypeScript 変数の型は暗黙的 [に定義](https://www.typescriptlang.org/docs/handbook/type-inference.html) できます。</span><span class="sxs-lookup"><span data-stu-id="212bc-125">TypeScript variable types can be [implicitly](https://www.typescriptlang.org/docs/handbook/type-inference.html) defined.</span></span> <span data-ttu-id="212bc-126">TypeScript コンパイラが変数の種類を特定できない場合 (型が明示的に定義されていないか、型の推論ができない場合)、暗黙的な値であり、コンパイル時エラーが発生します。 `any`</span><span class="sxs-lookup"><span data-stu-id="212bc-126">If the TypeScript compiler is unable to determine the type of a variable (either because type is not defined explicitly or type inference isn't possible), then it's an implicit `any` and you will receive a compilation-time error.</span></span>

<span data-ttu-id="212bc-127">暗黙的な場合の最も一般的 `any` なケースは、 などの変数宣言です `let value;` 。</span><span class="sxs-lookup"><span data-stu-id="212bc-127">The most common case on any implicit `any` is in a variable declaration, such as `let value;`.</span></span> <span data-ttu-id="212bc-128">これを回避するには、次の 2 つの方法があります。</span><span class="sxs-lookup"><span data-stu-id="212bc-128">There are two ways to avoid this:</span></span>

* <span data-ttu-id="212bc-129">変数を暗黙的に識別可能な型 (または) に割り `let value = 5;` 当 `let value = workbook.getWorksheet();` てる。</span><span class="sxs-lookup"><span data-stu-id="212bc-129">Assign the variable to an implicitly identifiable type (`let value = 5;` or `let value = workbook.getWorksheet();`).</span></span>
* <span data-ttu-id="212bc-130">変数 ( ) を明示的に `let value: number;` 入力します。</span><span class="sxs-lookup"><span data-stu-id="212bc-130">Explicitly type the variable (`let value: number;`)</span></span>

## <a name="no-inheriting-office-script-classes-or-interfaces"></a><span data-ttu-id="212bc-131">スクリプト クラスまたはOffice継承なし</span><span class="sxs-lookup"><span data-stu-id="212bc-131">No inheriting Office Script classes or interfaces</span></span>

<span data-ttu-id="212bc-132">スクリプトで作成されたクラスとインターフェイスOffice Script クラスまたはインターフェイスOffice[](https://www.typescriptlang.org/docs/handbook/classes.html#inheritance)拡張または実装できません。</span><span class="sxs-lookup"><span data-stu-id="212bc-132">Classes and interfaces that are created in your Office Script cannot [extend or implement](https://www.typescriptlang.org/docs/handbook/classes.html#inheritance) Office Scripts classes or interfaces.</span></span> <span data-ttu-id="212bc-133">つまり、名前空間にサブクラスやサブインターフェイス `ExcelScript` を含め得るものは何もありません。</span><span class="sxs-lookup"><span data-stu-id="212bc-133">In other words, nothing in the `ExcelScript` namespace can have subclasses or subinterfaces.</span></span>

## <a name="incompatible-typescript-functions"></a><span data-ttu-id="212bc-134">互換性のない TypeScript 関数</span><span class="sxs-lookup"><span data-stu-id="212bc-134">Incompatible TypeScript functions</span></span>

<span data-ttu-id="212bc-135">Officeスクリプト API は、以下では使用できません。</span><span class="sxs-lookup"><span data-stu-id="212bc-135">Office Scripts APIs cannot be used in the following:</span></span>

* [<span data-ttu-id="212bc-136">ジェネレーター関数</span><span class="sxs-lookup"><span data-stu-id="212bc-136">Generator functions</span></span>](https://developer.mozilla.org/docs/Web/JavaScript/Guide/Iterators_and_Generators#generator_functions)
* [<span data-ttu-id="212bc-137">Array.sort</span><span class="sxs-lookup"><span data-stu-id="212bc-137">Array.sort</span></span>](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Array/sort)

## <a name="eval-is-not-supported"></a><span data-ttu-id="212bc-138">`eval` サポートされていません</span><span class="sxs-lookup"><span data-stu-id="212bc-138">`eval` is not supported</span></span>

<span data-ttu-id="212bc-139">JavaScript [eval 関数は](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/eval) 、セキュリティ上の理由からサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="212bc-139">The JavaScript [eval function](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/eval) is not supported for security reasons.</span></span>

## <a name="restricted-identifers"></a><span data-ttu-id="212bc-140">制限付き identifers</span><span class="sxs-lookup"><span data-stu-id="212bc-140">Restricted identifers</span></span>

<span data-ttu-id="212bc-141">次の単語は、スクリプト内の識別子として使用できません。</span><span class="sxs-lookup"><span data-stu-id="212bc-141">The following words can't be used as identifiers in a script.</span></span> <span data-ttu-id="212bc-142">これらは予約済みの用語です。</span><span class="sxs-lookup"><span data-stu-id="212bc-142">They are reserved terms.</span></span>

* `Excel`
* `ExcelScript`
* `console`

## <a name="only-arrow-functions-in-array-callbacks"></a><span data-ttu-id="212bc-143">配列コールバックの矢印関数のみ</span><span class="sxs-lookup"><span data-stu-id="212bc-143">Only arrow functions in array callbacks</span></span>

<span data-ttu-id="212bc-144">スクリプトは、Array メソッド [にコールバック引数](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Functions/Arrow_functions) を指定する場合にのみ矢印関数 [を](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Array) 使用できます。</span><span class="sxs-lookup"><span data-stu-id="212bc-144">Your scripts can only use [arrow functions](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Functions/Arrow_functions) when providing callback arguments for [Array](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Array) methods.</span></span> <span data-ttu-id="212bc-145">これらのメソッドには、任意の種類の識別子または "従来の" 関数を渡す必要があります。</span><span class="sxs-lookup"><span data-stu-id="212bc-145">You cannot pass any sort of identifier or "traditional" function to these methods.</span></span>

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

## <a name="performance-warnings"></a><span data-ttu-id="212bc-146">パフォーマンスに関する警告</span><span class="sxs-lookup"><span data-stu-id="212bc-146">Performance warnings</span></span>

<span data-ttu-id="212bc-147">コード エディターの [linter は、](https://wikipedia.org/wiki/Lint_(software)) スクリプトにパフォーマンスの問題が発生する可能性がある場合に警告を表示します。</span><span class="sxs-lookup"><span data-stu-id="212bc-147">The Code Editor's [linter](https://wikipedia.org/wiki/Lint_(software)) gives warnings if the script might have performance issues.</span></span> <span data-ttu-id="212bc-148">ケースとその回避方法については、「スクリプトのパフォーマンスを向上させる」[にOfficeされています](web-client-performance.md)。</span><span class="sxs-lookup"><span data-stu-id="212bc-148">The cases and how to work around them are documented in [Improve the performance of your Office Scripts](web-client-performance.md).</span></span>

## <a name="external-api-calls"></a><span data-ttu-id="212bc-149">外部 API 呼び出し</span><span class="sxs-lookup"><span data-stu-id="212bc-149">External API calls</span></span>

<span data-ttu-id="212bc-150">詳細[については、「外部 API 呼び出しOfficeスクリプト」](external-calls.md)を参照してください。</span><span class="sxs-lookup"><span data-stu-id="212bc-150">See [External API call support in Office Scripts](external-calls.md) for more information.</span></span>

## <a name="see-also"></a><span data-ttu-id="212bc-151">関連項目</span><span class="sxs-lookup"><span data-stu-id="212bc-151">See also</span></span>

* [<span data-ttu-id="212bc-152">Excel on the web での Office スクリプトのスクリプトの基本事項</span><span class="sxs-lookup"><span data-stu-id="212bc-152">Scripting fundamentals for Office Scripts in Excel on the web</span></span>](scripting-fundamentals.md)
* [<span data-ttu-id="212bc-153">スクリプトのパフォーマンスをOfficeする</span><span class="sxs-lookup"><span data-stu-id="212bc-153">Improve the performance of your Office Scripts</span></span>](web-client-performance.md)
