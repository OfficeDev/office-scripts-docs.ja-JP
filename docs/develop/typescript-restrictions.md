---
title: スクリプトの TypeScript Office
description: スクリプト コード エディターで使用される TypeScript コンパイラと linter Office詳細です。
ms.date: 02/05/2021
localization_priority: Normal
ms.openlocfilehash: 8c9d1beafb236e7ba10dedf00fab944c40fb954d
ms.sourcegitcommit: 5d24e77df70aa2c1c982275d53213c2a9323ff86
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 04/02/2021
ms.locfileid: "51570277"
---
# <a name="typescript-restrictions-in-office-scripts"></a>スクリプトの TypeScript Office

Officeは TypeScript 言語を使用します。 ほとんどの場合、TypeScript または JavaScript のコードは、スクリプトスクリプトでOfficeされます。 ただし、スクリプトが Excel ブックで意図した通り一貫して動作することを確認するために、コード エディターによっていくつかの制限が適用されています。

## <a name="no-any-type-in-office-scripts"></a>スクリプトに 'any' 型Officeはありません

[TypeScript](https://www.typescriptlang.org/docs/handbook/typescript-in-5-minutes.html)では、型を推論できるので、書き込み型は省略可能です。 ただし、Officeスクリプトでは、変数を任意の型に [できない必要があります](https://www.typescriptlang.org/docs/handbook/basic-types.html#any)。 明示的および暗黙的の両方 `any` は、スクリプト内でOfficeされません。 これらのケースはエラーとして報告されます。

### <a name="explicit-any"></a>明示的 `any`

スクリプト (つまり) で変数を明示的に型 `any` Office宣言することはできません `let someVariable: any;` 。 Excel `any` で処理すると、この型によって問題が発生します。 たとえば、値が 、 、 または である必要 `Range` `string` `number` があります `boolean` 。 スクリプト内の型として変数が明示的に定義されている場合は、コンパイル時エラー (スクリプトを実行する前のエラー) `any` が表示されます。

![コード エディターのホバー テキスト内の明示的なメッセージ](../images/explicit-any-editor-message.png)

![コンソール ウィンドウでの明示的なエラー](../images/explicit-any-error-message.png)

上のスクリーンショットでは `[5, 16] Explicit Any is not allowed` 、行の種類を#5列#16示 `any` しています。 これにより、エラーを見つけるのに役立ちます。

この問題を回避するには、常に変数の種類を定義します。 変数の種類が不明な場合は、共用体の型を [使用できます](https://www.typescriptlang.org/docs/handbook/unions-and-intersections.html)。 これは、型 、または (値の型は、それらの共用体です) の値を保持する変数 `Range` `string` `number` `boolean` `Range` に役立ちます `string | number | boolean` 。

### <a name="implicit-any"></a>暗黙的 `any`

TypeScript 変数の型は暗黙的 [に定義](https://www.typescriptlang.org/docs/handbook/type-inference.html) できます。 TypeScript コンパイラが変数の種類を特定できない場合 (型が明示的に定義されていないか、型の推論ができない場合)、暗黙的な値であり、コンパイル時エラーが発生します。 `any`

暗黙的な場合の最も一般的 `any` なケースは、 などの変数宣言です `let value;` 。 これを回避するには、次の 2 つの方法があります。

* 変数を暗黙的に識別可能な型 (または) に割り `let value = 5;` 当 `let value = workbook.getWorksheet();` てる。
* 変数 ( ) を明示的に `let value: number;` 入力します。

## <a name="no-inheriting-office-script-classes-or-interfaces"></a>スクリプト クラスまたはOffice継承なし

スクリプトで作成されたクラスとインターフェイスはOfficeスクリプト クラスまたはインターフェイス [Office](https://www.typescriptlang.org/docs/handbook/classes.html#inheritance) 拡張または実装できません。 つまり、名前空間にサブクラスやサブインターフェイス `ExcelScript` を含め得るものは何もありません。

## <a name="incompatible-typescript-functions"></a>互換性のない TypeScript 関数

Officeスクリプト API は、以下では使用できません。

* [ジェネレーター関数](https://developer.mozilla.org/docs/Web/JavaScript/Guide/Iterators_and_Generators#generator_functions)
* [Array.sort](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Array/sort)

## <a name="eval-is-not-supported"></a>`eval` サポートされていません

JavaScript [eval 関数は](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/eval) 、セキュリティ上の理由からサポートされていません。

## <a name="restricted-identifers"></a>制限付き identifers

次の単語は、スクリプト内の識別子として使用できません。 これらは予約済みの用語です。

* `Excel`
* `ExcelScript`
* `console`

## <a name="only-arrow-functions-in-array-callbacks"></a>配列コールバックの矢印関数のみ

スクリプトは、Array メソッド [にコールバック引数](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Functions/Arrow_functions) を指定する場合にのみ矢印関数 [を](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Array) 使用できます。 これらのメソッドには、任意の種類の識別子または "従来の" 関数を渡す必要があります。

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

## <a name="performance-warnings"></a>パフォーマンスに関する警告

コード エディターの [linter は、](https://wikipedia.org/wiki/Lint_(software)) スクリプトにパフォーマンスの問題が発生する可能性がある場合に警告を表示します。 ケースとその回避方法については、「スクリプトのパフォーマンスを向上させる」 [にOfficeされています](web-client-performance.md)。

## <a name="external-api-calls"></a>外部 API 呼び出し

詳細 [については、「外部 API 呼び出しOfficeスクリプト」](external-calls.md) を参照してください。

## <a name="see-also"></a>関連項目

* [Excel on the web での Office スクリプトのスクリプトの基本事項](scripting-fundamentals.md)
* [スクリプトのパフォーマンスをOfficeする](web-client-performance.md)
