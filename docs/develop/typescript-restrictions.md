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
# <a name="typescript-restrictions-in-office-scripts"></a>Officeスクリプトにおける TypeScript の制限

Officeスクリプトは、タイプスクリプト言語を使用します。 ほとんどの場合、TypeScript または JavaScript コードはOfficeスクリプトで動作します。 ただし、スクリプトが一貫して、Excelブックで意図したとおりに動作するように、コード エディターによって適用される制限がいくつかあります。

## <a name="no-any-type-in-office-scripts"></a>Officeスクリプトに 「任意」 タイプがありません

[TypeScript](https://www.typescriptlang.org/docs/handbook/typescript-in-5-minutes.html)では型を推論できるため、型の書き込みはオプションです。 ただし、Officeスクリプトでは、変数を[任意の型](https://www.typescriptlang.org/docs/handbook/basic-types.html#any)にすることはできません。 `any`Office スクリプトでは、明示的なスクリプトと暗黙的な両方を使用できません。 これらのケースはエラーとして報告されます。

### <a name="explicit-any"></a>暁 `any`

`any`Officeスクリプト (つまり) で変数を型として明示的に宣言することはできません `let someVariable: any;` 。 `any`この型は、Excelで処理されるときに問題を引き起こします。 たとえば、 の `Range` 値が `string` `number` 、 、または であることを知る必要があります `boolean` 。 スクリプト内で型として明示的に定義されている変数がある場合、コンパイル エラー (スクリプトの実行前にエラー) `any` が表示されます。

:::image type="content" source="../images/explicit-any-editor-message.png" alt-text="コード エディターのホバー テキスト内の明示的な 'any' メッセージ":::

:::image type="content" source="../images/explicit-any-error-message.png" alt-text="コンソール ウィンドウでの明示的な 'any' エラー":::

前のスクリーンショットでは `[5, 16] Explicit Any is not allowed` 、行#5、列#16が型を定義していることを示 `any` しています。 これにより、エラーを見つけることができます。

この問題を回避するには、常に変数の型を定義します。 変数の型が不明な場合は、 [共用体型](https://www.typescriptlang.org/docs/handbook/unions-and-intersections.html)を使用できます。 これは、値を保持する変数 `Range` (型 `string` `number` 、、または `boolean` 値の型 `Range` がそれらの和集合である) `string | number | boolean` に役立ちます。

### <a name="implicit-any"></a>暗黙の `any`

TypeScript 変数型は [、暗黙的に](https://www.typescriptlang.org/docs/handbook/type-inference.html) 定義できます。 TypeScript コンパイラが変数の型を判断できない場合 (型が明示的に定義されていないか、型の推論が不可能な場合)、暗黙的なエラー `any` が発生します。

暗黙的な場合の最も一般的なケース `any` は、変数宣言 (など `let value;` ) です。 これを回避するには、次の 2 つの方法があります。

* 暗黙的に識別できる型 ( または ) に変数を割り当てます `let value = 5;` `let value = workbook.getWorksheet();` 。
* 変数を明示的に入力します ( `let value: number;` )

## <a name="no-inheriting-office-script-classes-or-interfaces"></a>継承Officeスクリプトクラスまたはインタフェースなし

Office スクリプトで作成されたクラスおよびインターフェイスは、スクリプトクラスまたはOfficeスクリプト クラスまたはインターフェイスを[拡張または実装](https://www.typescriptlang.org/docs/handbook/classes.html#inheritance)できません。 つまり、名前空間内のサブ `ExcelScript` クラスまたはサブインターフェイスを持つものはありません。

## <a name="incompatible-typescript-functions"></a>互換性のないタイプスクリプト関数

Officeスクリプト API は、次の場合は使用できません。

* [ジェネレータ関数](https://developer.mozilla.org/docs/Web/JavaScript/Guide/Iterators_and_Generators#generator_functions)
* [配列.ソート](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Array/sort)

## <a name="eval-is-not-supported"></a>`eval` はサポートされていません

セキュリティ上の理由から、JavaScript [eval 関数](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/eval) はサポートされていません。

## <a name="restricted-identifers"></a>制限付き識別コード

次の単語は、スクリプトの識別子として使用できません。 彼らは予約された用語です。

* `Excel`
* `ExcelScript`
* `console`

## <a name="only-arrow-functions-in-array-callbacks"></a>配列コールバックの矢印関数のみ

スクリプトは[、Array](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Array)メソッドにコールバック引数を指定する場合にのみ[、矢印関数](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Functions/Arrow_functions)を使用できます。 これらのメソッドに対して、いかなる種類の識別子や「伝統的な」関数も渡すことはできません。

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

## <a name="performance-warnings"></a>パフォーマンスの警告

スクリプトにパフォーマンスの問題がある場合、コード エディタの [リンター](https://wikipedia.org/wiki/Lint_(software)) は警告を表示します。 ケースとその回避方法については[、「Office スクリプトのパフォーマンスを向上させる](web-client-performance.md)」を参照してください。

## <a name="external-api-calls"></a>外部 API 呼び出し

詳細については[、「Officeスクリプト」の外部 API コールサポート](external-calls.md)を参照してください。

## <a name="see-also"></a>関連項目

* [Excel on the web での Office スクリプトのスクリプトの基本事項](scripting-fundamentals.md)
* [Officeスクリプトのパフォーマンスを向上させる](web-client-performance.md)
