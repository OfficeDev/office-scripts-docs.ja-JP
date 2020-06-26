---
title: 従来のスクリプトをサポートするための Office スクリプト非同期 Api の使用
description: Office スクリプト非同期 Api の入門と、従来のスクリプトでロード/同期パターンを使用する方法について説明します。
ms.date: 06/22/2020
localization_priority: Normal
ms.openlocfilehash: c7b3c1401ecc2b4d0371590e71f61ae6e9ad8a9d
ms.sourcegitcommit: aec3c971c6640429f89b6bb99d2c95ea06725599
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 06/25/2020
ms.locfileid: "44878854"
---
# <a name="using-the-office-scripts-async-apis-to-support-legacy-scripts"></a>従来のスクリプトをサポートするための Office スクリプト非同期 Api の使用

この記事では、従来の非同期の Api を使用してスクリプトを記述する方法について説明します。 これらの Api は、標準の同期された Office スクリプト Api と同じコア機能を備えていますが、スクリプトとブックとの間のデータ同期を制御する必要があります。

> [!IMPORTANT]
> Async モデルは、現在の[API モデル](scripting-fundamentals.md?view=office-scripts)を実装する前に作成されたスクリプトでのみ使用できます。 スクリプトは、作成時に作成した API モデルに完全にロックされます。 これは、レガシスクリプトを新しいモデルに変換する場合は、新しいスクリプトを使用する必要があることも意味します。 現在のモデルは使いやすいため、変更時に古いスクリプトを新しいモデルに更新することをお勧めします。 この移行を実行する方法については、「[従来の非同期スクリプトを現在のモデルに変換](#converting-legacy-async-scripts-to-the-current-model)する」セクションを参照してください。

## <a name="main-function"></a>`main` 関数

非同期 Api を使用するスクリプトは、別の関数を備えてい `main` ます。 これは `async` 、を `Excel.RequestContext` 最初のパラメーターとして持つ関数です。

```TypeScript
async function main(context: Excel.RequestContext) {
    // Your async Office Script
}
```

## <a name="context"></a>コンテキスト

`main` 関数は、`context` という名前の `Excel.RequestContext` パラメーターを受け入れます。 `context` は、スクリプトとブックの間のブリッジと見なすことができます。 スクリプトは、`context` オブジェクトを使用してブックにアクセスし、その `context` を使用してデータをやり取りします。

スクリプトと Excel は異なるプロセスや場所で実行されているため、`context` オブジェクトが必要になります。 スクリプトで、クラウドのブックに変更を加えたり、そのブックからデータをクエリしたりする必要があります。 `context` オブジェクトは、それらのトランザクションを管理します。

## <a name="sync-and-load"></a>同期と読み込み

スクリプトとブックは別の場所で実行されるため、両者の間でデータを転送するには時間がかかります。 非同期 API では、スクリプトとブックを同期する操作をスクリプトが明示的に呼び出すまで、コマンドがキューに登録され `sync` ます。 スクリプトは、次のどちらかを実行することが必要になるまで、独立して動作できます。

- ブックからデータを読み取る (`load` 操作または [ClientResult](/javascript/api/office-scripts/excel/excel.clientresult?view=office-scripts-async) を返すメソッドの後)。
- ブックにデータを書き込む (通常はスクリプトが完了した結果)。

次の図に、スクリプトとブックの間の制御フローの例を示します。

![スクリプトからブックに対して実行される読み取りおよび書き込み操作を示す図。](../images/load-sync.png)

### <a name="sync"></a>同期

非同期スクリプトでブックのデータを読み取る必要がある場合、またはブックにデータを書き込む必要がある場合は、次のようにメソッドを呼び出し `RequestContext.sync` ます。

```TypeScript
await context.sync();
```

> [!NOTE]
> スクリプトが終了すると、`context.sync()` が暗黙的に呼び出されます。

`sync` 操作が完了すると、ブックが更新され、スクリプトが指定した書き込み操作が反映されます。 書き込み操作とは、Excel オブジェクトに任意のプロパティを設定すること (`range.format.fill.color = "red"` など)、またはプロパティを変更するメソッドを呼び出すこと (`range.format.autoFitColumns()` など) を意味します。 また、`sync` 操作では、スクリプトが `load` 操作または `ClientResult` を返すメソッドを使用して要求したブックから任意の値が読み取られます (次のセクションを参照)。

ネットワークによっては、スクリプトとブックを同期するのに時間がかかる場合があります。 `sync`スクリプトの実行速度を速くするために、呼び出しの数を最小限に抑えます。 それ以外の場合、非同期 Api は標準の同期 Api よりも高速ではありません。

### <a name="load"></a>読み込み

非同期スクリプトを読み取る前に、ブックからデータを読み込む必要があります。 ただし、ブック全体からデータを読み込むと、スクリプトの速度が大幅に低下します。 このメソッドを使用すると、 `load` ブックからどのようなデータを取得するかをスクリプトで明示的に指定できます。

`load` メソッドは、すべての Excel オブジェクトで使用できます。 スクリプトでは、オブジェクトのプロパティを読み込んでからでなければ、それらを読み取ることができません。 そうしないと、エラーになります。

次の例では、`Range` オブジェクトを使用して、`load` メソッドでデータを読み込む方法を示します。

|目的 |コマンドの例 | 効果 |
|:--|:--|:--|
|1 つのプロパティを読み込む |`myRange.load("values");` | 単一のプロパティ (この例では、範囲内の値の 2 次元配列) を読み込みます。 |
|複数のプロパティを読み込む |`myRange.load("values, rowCount, columnCount");`| コンマで区切られたリストからすべてのプロパティ (この例では、値、行数、列数) を読み込みます。 |
|すべてを読み込む | `myRange.load();`|範囲のすべてのプロパティを読み込みます。 これは、不要なデータを取得してスクリプトを低速にするために推奨される解決策ではありません。 これは、スクリプトをテストする場合、またはオブジェクトのすべてのプロパティを必要とする場合にのみ使用してください。 |

スクリプトでは、読み込まれた値を読み取る前に、`context.sync()` を呼び出す必要があります。

```TypeScript
/**
 * This script uses the async API to get the row count for a range.
 * It shows how to load a property in the async model.
 */
async function main(context: Excel.RequestContext) {
    let selectedSheet = context.workbook.worksheets.getActiveWorksheet();
    let range = selectedSheet.getRange("A1:B3");

    // Load the property.
    range.load("rowCount");

    // Synchronize with the workbook to get the property.
    await context.sync();

    // Read and log the property value (3).
    console.log(range.rowCount);
}
```

また、コレクション全体のプロパティを読み込むこともできます。 Async API のすべてのコレクションオブジェクトには、 `items` そのコレクション内のオブジェクトを含む配列であるプロパティがあります。 `items` を `load` に対する階層呼び出し (`items\myProperty`) の最初に使用すると、それらの項目それぞれの指定されたプロパティが読み込まれます。 次の例では、ワークシートの `CommentCollection` オブジェクトに含まれる各 `Comment` オブジェクトの `resolved` プロパティが読み込まれます。

```TypeScript
/**
 * This script uses the async API to get resolved property on every comment in the worksheet.
 * It shows how to load a property from every object in a collection.
 */
async function main(context: Excel.RequestContext){
    let selectedSheet = context.workbook.worksheets.getActiveWorksheet();
    let comments = selectedSheet.comments;

    // Load the `resolved` property from every comment in this collection.
    comments.load("items/resolved");

    // Synchronize with the workbook to get the properties.
    await context.sync();
}
```

### <a name="clientresult"></a>ClientResult

ブックから情報を返す非同期 API のメソッドには、パラダイムに似たパターンがあり `load` / `sync` ます。 たとえば、`TableCollection.getCount` はコレクション内のテーブルの数を取得します。 `getCount` は `ClientResult<number>` を返します。つまり、返される `ClientResult` の `value` プロパティは数値になります。 `context.sync()` が呼び出されるまで、スクリプトはその値にアクセスできません。 プロパティの読み込みと同様、`value` は、`sync` が呼び出されるまでは、ローカルの "空の" 値です。

次のスクリプトは、ブック内のテーブルの総数を取得し、その数をコンソールに記録します。

```TypeScript
/**
 * This script uses the async API to get the table count of the workbook.
 * It shows how ClientResult objects return workbook information.
 */
async function main(context: Excel.RequestContext) {
    let tableCount = context.workbook.tables.getCount();

    // This sync call implicitly loads tableCount.value.
    // Any other ClientResult values are loaded too.
    await context.sync();

    // Trying to log the value before calling sync would throw an error.
    console.log(tableCount.value);
}
```

## <a name="converting-legacy-async-scripts-to-the-current-model"></a>従来の非同期スクリプトを現在のモデルに変換する

現在の API モデルでは、、、またはを使用しません `load` `sync` `RequestContext` 。 これにより、スクリプトがより簡単に作成および管理できるようになります。 古いスクリプトを変換するための最善のリソースは、[スタックオーバーフロー](https://stackoverflow.com/questions/tagged/office-scripts)です。 ここでは、特定のシナリオについてコミュニティにサポートを求めることができます。 次のガイダンスは、実行する必要のある一般的な手順の概要を示すために役立ちます。

1. 新しいスクリプトを作成し、それに古い非同期コードをコピーします。 代わりに、現在の方法を使用して、古いメソッド署名を含めないようにしてください `main` `function main(workbook: ExcelScript.Workbook)` 。

2. との呼び出しをすべて削除し `load` `sync` ます。 これらは不要になりました。

3. すべてのプロパティが削除されました。 これらのオブジェクトに and メソッドを使用してアクセスできるようになった `get` `set` ので、これらのプロパティ参照をメソッド呼び出しに切り替える必要があります。 たとえば、次のようなプロパティアクセスを使用してセルの塗りつぶしの色を設定するのではなく、次の `mySheet.getRange("A2:C2").format.fill.color = "blue";` ようなメソッドを使用します。`mySheet.getRange("A2:C2").getFormat().getFill().setColor("blue");`

4. コレクションクラスは、配列に置き換えられました。 `add` `get` これらのコレクションクラスのメソッドとメソッドは、コレクションを所有していたオブジェクトに移動されたので、それに応じて参照を更新する必要があります。 たとえば、ブックの最初のワークシートから "MyChart" という名前のグラフを取得するには、次のコードを使用 `workbook.getWorksheets()[0].getChart("MyChart");` します。 `[0]`で返されるの最初の値にアクセスするには、に注意し `Worksheet[]` て `getWorksheets()` ください。

5. わかりやすくするために名前が変更されたメソッドもあります。 詳細については、「 [Office SCRIPTS API リファレンス](/javascript/api/office-scripts/overview?view=office-scripts)」を参照してください。

## <a name="office-scripts-async-api-reference-documentation"></a>Office スクリプトの非同期 API リファレンスドキュメント

[!INCLUDE [Async reference documentation](../includes/async-reference-documentation-link.md)]
