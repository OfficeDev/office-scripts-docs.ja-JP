---
title: Office スクリプトでの TypeScript の制限
description: Office Scripts コード エディターで使用される TypeScript コンパイラと linter の詳細。
ms.date: 01/29/2021
localization_priority: Normal
ms.openlocfilehash: 41584ff23b333d17b2e267fdb3b0ec8741f3d203
ms.sourcegitcommit: df2b64603f91acb37bf95230efd538db0fbf9206
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 02/04/2021
ms.locfileid: "50099909"
---
# <a name="typescript-restrictions-in-office-scripts"></a><span data-ttu-id="37270-103">Office スクリプトでの TypeScript の制限</span><span class="sxs-lookup"><span data-stu-id="37270-103">TypeScript restrictions in Office Scripts</span></span>

<span data-ttu-id="37270-104">Officeは TypeScript 言語を使用します。</span><span class="sxs-lookup"><span data-stu-id="37270-104">Office Scripts use the TypeScript language.</span></span> <span data-ttu-id="37270-105">ほとんどの場合、TypeScript または JavaScript のコードは、Officeされます。</span><span class="sxs-lookup"><span data-stu-id="37270-105">For the most part, any TypeScript or JavaScript code will work in an Office Script.</span></span> <span data-ttu-id="37270-106">ただし、スクリプトが Excel ブックで一貫して意図した方法で動作することを保証するために、コード エディターによっていくつかの制限が適用されます。</span><span class="sxs-lookup"><span data-stu-id="37270-106">However, there are a few restrictions enforced by the Code Editor to ensure your script works consistently and as intended with your Excel workbook.</span></span>

## <a name="no-any-type-in-office-scripts"></a><span data-ttu-id="37270-107">スクリプトに 'any' 型Officeはありません</span><span class="sxs-lookup"><span data-stu-id="37270-107">No 'any' type in Office Scripts</span></span>

<span data-ttu-id="37270-108">[TypeScript](https://www.typescriptlang.org/docs/handbook/typescript-in-5-minutes.html)では型を推測できるので、書き込み型は省略可能です。</span><span class="sxs-lookup"><span data-stu-id="37270-108">Writing [types](https://www.typescriptlang.org/docs/handbook/typescript-in-5-minutes.html) is optional in TypeScript, because the types can be inferred.</span></span> <span data-ttu-id="37270-109">ただし、Officeスクリプトでは、変数の型を指定できない [必要があります](https://www.typescriptlang.org/docs/handbook/basic-types.html#any)。</span><span class="sxs-lookup"><span data-stu-id="37270-109">However, Office Script requires that a variable can't be of [type any](https://www.typescriptlang.org/docs/handbook/basic-types.html#any).</span></span> <span data-ttu-id="37270-110">スクリプトでは、明示的と暗黙的 `any` の両方Office使用できます。</span><span class="sxs-lookup"><span data-stu-id="37270-110">Both explicit and implicit `any` are not allowed in an Office Script.</span></span> <span data-ttu-id="37270-111">これらのケースは、エラーとして報告されます。</span><span class="sxs-lookup"><span data-stu-id="37270-111">These cases are reported as errors.</span></span>

### <a name="explicit-any"></a><span data-ttu-id="37270-112">Explicit `any`</span><span class="sxs-lookup"><span data-stu-id="37270-112">Explicit `any`</span></span>

<span data-ttu-id="37270-113">スクリプト (つまり) で変数を型として `any` 明示的Office宣言することはできません `let someVariable: any;` 。</span><span class="sxs-lookup"><span data-stu-id="37270-113">You cannot explicitly declare a variable to be of type `any` in Office Scripts (that is, `let someVariable: any;`).</span></span> <span data-ttu-id="37270-114">この `any` 型は、Excel によって処理される際に問題を引き起こします。</span><span class="sxs-lookup"><span data-stu-id="37270-114">The `any` type causes issues when processed by Excel.</span></span> <span data-ttu-id="37270-115">たとえば、値が a , , または `Range` `string` . `number` `boolean`</span><span class="sxs-lookup"><span data-stu-id="37270-115">For example, a `Range` needs to know that a value is a `string`, `number`, or `boolean`.</span></span> <span data-ttu-id="37270-116">スクリプト内の型として変数が明示的に定義されている場合は、コンパイル時エラー (スクリプトを実行する前にエラー) `any` が表示されます。</span><span class="sxs-lookup"><span data-stu-id="37270-116">You will receive a compile-time error (an error prior to running the script) if any variable is explicitly defined as the `any` type in the script.</span></span>

![コード エディターのホバー テキスト内の明示的なメッセージ](../images/explicit-any-editor-message.png)

![コンソール ウィンドウでの明示的なエラー](../images/explicit-any-error-message.png)

<span data-ttu-id="37270-119">上のスクリーンショットでは `[5, 16] Explicit Any is not allowed` 、行の種類が#5列#16示 `any` しています。</span><span class="sxs-lookup"><span data-stu-id="37270-119">In the above screenshot `[5, 16] Explicit Any is not allowed` indicates that line #5, column #16 defines `any` type.</span></span> <span data-ttu-id="37270-120">これにより、エラーを見つけるのに役立ちます。</span><span class="sxs-lookup"><span data-stu-id="37270-120">This helps you locate the error.</span></span>

<span data-ttu-id="37270-121">この問題を回避するには、変数の型を必ず定義してください。</span><span class="sxs-lookup"><span data-stu-id="37270-121">To get around this issue, always define the type of the variable.</span></span> <span data-ttu-id="37270-122">変数の型が不明な場合は、ユニオン型を [使用できます](https://www.typescriptlang.org/docs/handbook/unions-and-intersections.html)。</span><span class="sxs-lookup"><span data-stu-id="37270-122">If you are uncertain about the type of a variable, you can use a [union type](https://www.typescriptlang.org/docs/handbook/unions-and-intersections.html).</span></span> <span data-ttu-id="37270-123">これは、値を保持する変数 (型、または値の型は次の値のユニオン) に `Range` `string` `number` `boolean` `Range` 役立ちます `string | number | boolean` 。</span><span class="sxs-lookup"><span data-stu-id="37270-123">This can be useful for variables that hold `Range` values, which can be of type `string`, `number`, or `boolean` (the type for `Range` values is a union of those: `string | number | boolean`).</span></span>

### <a name="implicit-any"></a><span data-ttu-id="37270-124">暗黙的 `any`</span><span class="sxs-lookup"><span data-stu-id="37270-124">Implicit `any`</span></span>

<span data-ttu-id="37270-125">TypeScript 変数型は[暗黙的](( https://www.typescriptlang.org/docs/handbook/type-inference.html) 定義) することができます。</span><span class="sxs-lookup"><span data-stu-id="37270-125">TypeScript variable types can be [implicitly]((https://www.typescriptlang.org/docs/handbook/type-inference.html) defined.</span></span> <span data-ttu-id="37270-126">TypeScript コンパイラが変数の型を特定できない場合 (型が明示的に定義されていないか、型の推論ができない場合)、暗黙的な値であり、コンパイル時エラーが発生します。 `any`</span><span class="sxs-lookup"><span data-stu-id="37270-126">If the TypeScript compiler is unable to determine the type of a variable (either because type is not defined explicitly or type inference isn't possible), then it's an implicit `any` and you will receive a compilation-time error.</span></span>

<span data-ttu-id="37270-127">暗黙的な宣言で最も一般的な `any` ケースは、次のような変数宣言です `let value;` 。</span><span class="sxs-lookup"><span data-stu-id="37270-127">The most common case on any implicit `any` is in a variable declaration, such as `let value;`.</span></span> <span data-ttu-id="37270-128">これを回避する方法は 2 種類あります。</span><span class="sxs-lookup"><span data-stu-id="37270-128">There are two ways to avoid this:</span></span>

* <span data-ttu-id="37270-129">変数を暗黙的に識別可能な型 (または) に割り当 `let value = 5;` てる `let value = workbook.getWorksheet();` 。</span><span class="sxs-lookup"><span data-stu-id="37270-129">Assign the variable to an implicitly identifiable type (`let value = 5;` or `let value = workbook.getWorksheet();`).</span></span>
* <span data-ttu-id="37270-130">変数 ( ) を明示的に入力 `let value: number;` します。</span><span class="sxs-lookup"><span data-stu-id="37270-130">Explicitly type the variable (`let value: number;`)</span></span>

## <a name="no-inheriting-office-script-classes-or-interfaces"></a><span data-ttu-id="37270-131">スクリプト クラスOffice継承なし</span><span class="sxs-lookup"><span data-stu-id="37270-131">No inheriting Office Script classes or interfaces</span></span>

<span data-ttu-id="37270-132">Office Script で作成されたクラスとインターフェイスは、Office Scripts[](https://www.typescriptlang.org/docs/handbook/classes.html#inheritance)クラスまたはインターフェイスを拡張または実装できません。</span><span class="sxs-lookup"><span data-stu-id="37270-132">Classes and interfaces that are created in your Office Script cannot [extend or implement](https://www.typescriptlang.org/docs/handbook/classes.html#inheritance) Office Scripts classes or interfaces.</span></span> <span data-ttu-id="37270-133">つまり、名前空間内のサブクラス `ExcelScript` やサブインターフェイスを持つものは何もありません。</span><span class="sxs-lookup"><span data-stu-id="37270-133">In other words, nothing in the `ExcelScript` namespace can have subclasses or subinterfaces.</span></span>

## <a name="incompatible-typescript-functions"></a><span data-ttu-id="37270-134">互換性のない TypeScript 関数</span><span class="sxs-lookup"><span data-stu-id="37270-134">Incompatible TypeScript functions</span></span>

<span data-ttu-id="37270-135">Officeスクリプト API は、以下では使用できません。</span><span class="sxs-lookup"><span data-stu-id="37270-135">Office Scripts APIs cannot be used in the following:</span></span>

* [<span data-ttu-id="37270-136">ジェネレーター関数</span><span class="sxs-lookup"><span data-stu-id="37270-136">Generator functions</span></span>](https://developer.mozilla.org/docs/Web/JavaScript/Guide/Iterators_and_Generators#generator_functions)
* [<span data-ttu-id="37270-137">Array.sort</span><span class="sxs-lookup"><span data-stu-id="37270-137">Array.sort</span></span>](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Array/sort)

## <a name="eval-is-not-supported"></a><span data-ttu-id="37270-138">`eval` サポートされていません</span><span class="sxs-lookup"><span data-stu-id="37270-138">`eval` is not supported</span></span>

<span data-ttu-id="37270-139">JavaScript [eval 関数](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/eval) は、セキュリティ上の理由からサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="37270-139">The JavaScript [eval function](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/eval) is not supported for security reasons.</span></span>

## <a name="restricted-identifers"></a><span data-ttu-id="37270-140">制限付き身元</span><span class="sxs-lookup"><span data-stu-id="37270-140">Restricted identifers</span></span>

<span data-ttu-id="37270-141">次の単語は、スクリプト内の識別子として使用できません。</span><span class="sxs-lookup"><span data-stu-id="37270-141">The following words can't be used as identifiers in a script.</span></span> <span data-ttu-id="37270-142">予約された用語です。</span><span class="sxs-lookup"><span data-stu-id="37270-142">They are reserved terms.</span></span>

* `Excel`
* `ExcelScript`
* `console`

## <a name="performance-warnings"></a><span data-ttu-id="37270-143">パフォーマンスの警告</span><span class="sxs-lookup"><span data-stu-id="37270-143">Performance warnings</span></span>

<span data-ttu-id="37270-144">スクリプトにパフォーマンスの問題 [がある可能性](https://wikipedia.org/wiki/Lint_(software)) がある場合は、コード エディターの linter が警告を表示します。</span><span class="sxs-lookup"><span data-stu-id="37270-144">The Code Editor's [linter](https://wikipedia.org/wiki/Lint_(software)) gives warnings if the script might have performance issues.</span></span> <span data-ttu-id="37270-145">ケースとその回避方法については、「スクリプトのパフォーマンスを向上させる [」Officeされています](web-client-performance.md)。</span><span class="sxs-lookup"><span data-stu-id="37270-145">The cases and how to work around them are documented in [Improve the performance of your Office Scripts](web-client-performance.md).</span></span>

## <a name="external-api-calls"></a><span data-ttu-id="37270-146">外部 API 呼び出し</span><span class="sxs-lookup"><span data-stu-id="37270-146">External API calls</span></span>

<span data-ttu-id="37270-147">詳細 [については、「Office Scripts」の](external-calls.md) 「外部 API 呼び出しのサポート」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="37270-147">See [External API call support in Office Scripts](external-calls.md) for more information.</span></span>

## <a name="see-also"></a><span data-ttu-id="37270-148">関連項目</span><span class="sxs-lookup"><span data-stu-id="37270-148">See also</span></span>

* [<span data-ttu-id="37270-149">Excel on the web での Office スクリプトのスクリプトの基本事項</span><span class="sxs-lookup"><span data-stu-id="37270-149">Scripting fundamentals for Office Scripts in Excel on the web</span></span>](scripting-fundamentals.md)
* [<span data-ttu-id="37270-150">スクリプトのパフォーマンスをOfficeする</span><span class="sxs-lookup"><span data-stu-id="37270-150">Improve the performance of your Office Scripts</span></span>](web-client-performance.md)
