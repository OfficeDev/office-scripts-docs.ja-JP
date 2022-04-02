---
title: スクリプトの TypeScript のOffice
description: スクリプト コード エディターで使用される TypeScript コンパイラと linter のOfficeします。
ms.date: 03/08/2022
ms.localizationpriority: medium
ms.openlocfilehash: b5ba0dfe60081a0bb65dec4e694c7d534cb8df63
ms.sourcegitcommit: 7023b9e23499806901a5ecf8ebc460b76887cca6
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 03/31/2022
ms.locfileid: "64585682"
---
# <a name="typescript-restrictions-in-office-scripts"></a>スクリプトの TypeScript のOffice

Officeは TypeScript 言語を使用します。 ほとんどの場合、TypeScript または JavaScript のコードは、スクリプトのOfficeされます。 ただし、コード エディターによって、スクリプトが一貫して動作し、ブックの目的に合Excelがあります。

## <a name="no-any-type-in-office-scripts"></a>スクリプトに 'any' 型Officeはありません

[TypeScript](https://www.typescriptlang.org/docs/handbook/typescript-in-5-minutes.html) では、型を推論できるので、書き込み型は省略可能です。 ただし、Officeスクリプトでは、変数を any 型にできない[必要があります](https://www.typescriptlang.org/docs/handbook/basic-types.html#any)。 明示的および暗黙的の両方`any`は、スクリプトのOfficeできません。 これらのケースはエラーとして報告されます。

### <a name="explicit-any"></a>明示的 `any`

スクリプト (つまり) で`any`変数を明示的に型Office宣言することはできません`let value: any;`。 この`any`型は、ユーザーが処理した場合に問題Excel。 たとえば、値が `Range` 、 、 または `string``number`である必要があります`boolean`。 スクリプト内の型として変数が明示的に定義されている場合は、コンパイル時エラー (スクリプトを実行する前のエラー) `any` が表示されます。

:::image type="content" source="../images/explicit-any-editor-message.png" alt-text="コード エディターのホバー テキストの明示的な 'any' メッセージ。":::

:::image type="content" source="../images/explicit-any-error-message.png" alt-text="コンソール ウィンドウの明示的な 'any' エラー。":::

前のスクリーンショットでは、行 `[2, 14] Explicit Any is not allowed` #2、列 #14 が型を定義します `any` 。 これにより、エラーを見つけるのに役立ちます。

この問題を回避するには、常に変数の種類を定義します。 変数の種類が不明な場合は、共用体の型を [使用できます](https://www.typescriptlang.org/docs/handbook/unions-and-intersections.html)。 これは、型 、または `Range` `string``boolean` `number`(`Range`値の型は、それらの共用体です) の値を保持する変数に役立ちます。 `string | number | boolean`

### <a name="implicit-any"></a>暗黙的 `any`

TypeScript 変数の型は暗黙的 [に定義](https://www.typescriptlang.org/docs/handbook/type-inference.html) できます。 TypeScript コンパイラが変数の種類を特定できない場合 ( `any` 型が明示的に定義されていないか、型の推論ができない場合)、暗黙的な値であり、コンパイル時エラーが発生します。

:::image type="content" source="../images/implicit-any-editor-message.png" alt-text="コード エディターのホバー テキスト内の暗黙的な 'any' メッセージ。":::

暗黙的な場合の最も一般的なケース `any` は、 などの変数宣言です `let value;`。 これを回避するには、次の 2 つの方法があります。

* 変数を暗黙的に識別可能な型 (または) に割り当`let value = 5;` てる `let value = workbook.getWorksheet();`。
* 変数を明示的に入力する (`let value: number;`)

## <a name="no-inheriting-office-script-classes-or-interfaces"></a>スクリプト クラスまたはOffice継承なし

スクリプトで作成されたクラスとインターフェイスはOfficeスクリプト クラスまたはインターフェイス[Office](https://www.typescriptlang.org/docs/handbook/classes.html#inheritance)拡張または実装できません。 つまり、名前空間にサブクラス `ExcelScript` やサブインターフェイスを含め得るものは何もありません。

## <a name="incompatible-typescript-functions"></a>互換性のない TypeScript 関数

Officeスクリプト API は、次では使用できません。

* [ジェネレーター関数](https://developer.mozilla.org/docs/Web/JavaScript/Guide/Iterators_and_Generators#generator_functions)
* [Array.sort](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Array/sort)

## <a name="eval-is-not-supported"></a>`eval` サポートされていません

JavaScript [eval 関数は](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/eval) 、セキュリティ上の理由からサポートされていません。

## <a name="restricted-identifiers"></a>制限付き識別子

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

## <a name="unions-of-excelscript-types-and-user-defined-types-arent-supported"></a>型とユーザー `ExcelScript` 定義型の共用体はサポートされていません

Officeスクリプトは、実行時に同期コード ブロックから非同期コード ブロックに変換されます。 約束を通じたブックとの通信 [は](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Promise) 、スクリプト作成者から隠されています。 この変換では、型と [ユーザー定義型](https://www.typescriptlang.org/docs/handbook/2/everyday-types.html#union-types) を含む `ExcelScript` 共用体型はサポートされません。 その場合は`Promise`、スクリプトに返されますが、Office スクリプト `Promise`コンパイラはスクリプトを期待し、スクリプト作成者は .

次のコード サンプルは、サポートされていないユニオンとカスタム インターフェイス `ExcelScript.Table` を示 `MyTable` しています。

```TypeScript
function main(workbook: ExcelScript.Workbook) {
  const selectedSheet = workbook.getActiveWorksheet();

  // This union is not supported.
  const tableOrMyTable: ExcelScript.Table | MyTable = selectedSheet.getTables()[0];

  // `getName` returns a promise that can't be resolved by the script.
  const name = tableOrMyTable.getName();

  // This logs "{}" instead of the table name.
  console.log(name);
}

interface MyTable {
  getName(): string
}
```

## <a name="constructors-dont-support-office-scripts-apis-and-console-statements"></a>コンストラクターは、スクリプト API Officeステートメントを`console`サポートしません

`console`ステートメントと多くのスクリプト Office API では、ブックとの同期がExcelされます。 これらの同期は、コンパイル `await` されたランタイム バージョンのスクリプトでステートメントを使用します。 `await` コンストラクターでは [サポートされていません](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Classes/constructor)。 コンストラクターを持つクラスが必要な場合は、Officeスクリプト API `console` またはこれらのコード ブロック内のステートメントを使用しないようにしてください。

次のコード サンプルは、このシナリオを示しています。 というエラーが生成されます `failed to load [code] [library]`。

```TypeScript
function main(workbook: ExcelScript.Workbook) {
  class MyClass {
    constructor() {
      // Console statements and Office Scripts APIs aren't supported in constructors.
      console.log("This won't print.");
    }
  }

  let test = new MyClass();
}
```

## <a name="performance-warnings"></a>パフォーマンスに関する警告

コード エディターの [linter は、](https://wikipedia.org/wiki/Lint_(software)) スクリプトにパフォーマンスの問題が発生する可能性がある場合に警告を表示します。 ケースとその回避方法については、「スクリプトのパフォーマンスを向上させる[」にOfficeされています](web-client-performance.md)。

## <a name="external-api-calls"></a>外部 API 呼び出し

詳細については[、「Office スクリプト」の「外部 API](external-calls.md) 呼び出しのサポート」を参照してください。

## <a name="see-also"></a>関連項目

* [Excel on the web での Office スクリプトのスクリプトの基本事項](scripting-fundamentals.md)
* [スクリプトのパフォーマンスをOfficeする](web-client-performance.md)
