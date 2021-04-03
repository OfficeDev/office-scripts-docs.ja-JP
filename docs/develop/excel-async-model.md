---
title: 非同期 API をOffice古いスクリプトをサポートする
description: スクリプト非同期 API のOffice、古いスクリプトに対して読み込み/同期パターンを使用する方法について説明します。
ms.date: 02/08/2021
localization_priority: Normal
ms.openlocfilehash: d61a5d8affae2077b23e140645c19dac977ff0d2
ms.sourcegitcommit: 5d24e77df70aa2c1c982275d53213c2a9323ff86
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 04/02/2021
ms.locfileid: "51570284"
---
# <a name="support-older-office-scripts-that-use-the-async-apis"></a>非同期 API をOffice古いスクリプトをサポートする

この記事では、古いモデルの非同期 API を使用するスクリプトを保守および更新する方法について説明します。 これらの API には、現在標準の同期 Office スクリプト API と同じコア機能がありますが、スクリプトとブックの間のデータ同期を制御するには、スクリプトが必要です。

> [!IMPORTANT]
> 非同期モデルは、現在の API モデルの実装前に作成されたスクリプトでのみ [使用できます](scripting-fundamentals.md)。 スクリプトは、作成時に持っている API モデルに完全にロックされます。 つまり、古いスクリプトを新しいモデルに変換する場合は、新しいスクリプトを作成する必要があります。 現在のモデルの方が使いやすいので、変更を加える場合は、古いスクリプトを新しいモデルに更新することをお勧めします。 [ [非同期スクリプトを現在のモデルに変換する](#converting-async-scripts-to-the-current-model) ] セクションには、この移行方法に関するアドバイスがあります。

## <a name="main-function"></a>`main` 関数

非同期 API を使用するスクリプトの機能は異 `main` なります。 これは、最初 `async` のパラメーターを持 `Excel.RequestContext` つ関数です。

```TypeScript
async function main(context: Excel.RequestContext) {
    // Your async Office Script
}
```

## <a name="context"></a>コンテキスト

`main` 関数は、`context` という名前の `Excel.RequestContext` パラメーターを受け入れます。 `context` は、スクリプトとブックの間のブリッジと見なすことができます。 スクリプトは、`context` オブジェクトを使用してブックにアクセスし、その `context` を使用してデータをやり取りします。

スクリプトと Excel は異なるプロセスや場所で実行されているため、`context` オブジェクトが必要になります。 スクリプトで、クラウドのブックに変更を加えたり、そのブックからデータをクエリしたりする必要があります。 `context` オブジェクトは、それらのトランザクションを管理します。

## <a name="sync-and-load"></a>同期と読み込み

スクリプトとブックは別の場所で実行されるため、両者の間でデータを転送するには時間がかかります。 非同期 API では、スクリプトがスクリプトとブックを同期する操作を明示的に呼び出すまで、コマンド `sync` はキューに入れられます。 スクリプトは、次のどちらかを実行することが必要になるまで、独立して動作できます。

- ブックからデータを読み取る (`load` 操作または [ClientResult](/javascript/api/office/officeextension.clientresult?view=excel-js-online&preserve-view=true) を返すメソッドの後)。
- ブックにデータを書き込む (通常はスクリプトが完了した結果)。

次の図に、スクリプトとブックの間の制御フローの例を示します。

![スクリプトからブックに対して実行される読み取りおよび書き込み操作を示す図。](../images/load-sync.png)

### <a name="sync"></a>同期

非同期スクリプトでブックからデータを読み取る必要がある場合やブックにデータを書き込む必要がある場合は、次に示すように `RequestContext.sync` メソッドを呼び出します。

```TypeScript
await context.sync();
```

> [!NOTE]
> スクリプトが終了すると、`context.sync()` が暗黙的に呼び出されます。

`sync` 操作が完了すると、ブックが更新され、スクリプトが指定した書き込み操作が反映されます。 書き込み操作は、Excel オブジェクト上の任意のプロパティ (例: ) を設定するか、プロパティを変更するメソッド (例: ) を呼び `range.format.fill.color = "red"` 出します `range.format.autoFitColumns()` 。 また、`sync` 操作では、スクリプトが `load` 操作または `ClientResult` を返すメソッドを使用して要求したブックから任意の値が読み取られます (次のセクションを参照)。

ネットワークによっては、スクリプトとブックを同期するのに時間がかかる場合があります。 スクリプトの高速実行 `sync` に役立つ呼び出しの数を最小限に抑える。 それ以外の場合、非同期 API は標準の同期 API より高速ではありません。

### <a name="load"></a>読み込み

非同期スクリプトは、ブックを読み取る前にブックからデータを読み込む必要があります。 ただし、ブック全体からデータを読み込むと、スクリプトの速度が大幅に低下します。 この `load` メソッドを使用すると、ブックから取得するデータをスクリプトに具体的に示すことができます。

`load` メソッドは、すべての Excel オブジェクトで使用できます。 スクリプトでは、オブジェクトのプロパティを読み込んでからでなければ、それらを読み取ることができません。 そうしない場合、エラーが発生します。

次の例では、`Range` オブジェクトを使用して、`load` メソッドでデータを読み込む方法を示します。

|目的 |コマンドの例 | 効果 |
|:--|:--|:--|
|1 つのプロパティを読み込む |`myRange.load("values");` | 単一のプロパティ (この例では、範囲内の値の 2 次元配列) を読み込みます。 |
|複数のプロパティを読み込む |`myRange.load("values, rowCount, columnCount");`| コンマで区切られたリストからすべてのプロパティ (この例では、値、行数、列数) を読み込みます。 |
|すべてを読み込む | `myRange.load();`|範囲のすべてのプロパティを読み込みます。 これは、不要なデータを取得することでスクリプトの速度が低下しますので、推奨されるソリューションではありません。 スクリプトのテスト中、またはオブジェクトのすべてのプロパティが必要な場合にのみ、これを使用します。 |

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

また、コレクション全体のプロパティを読み込むこともできます。 非同期 API 内のすべてのコレクション オブジェクトには、そのコレクション内のオブジェクトを含 `items` む配列であるプロパティがあります。 `items` を `load` に対する階層呼び出し (`items\myProperty`) の最初に使用すると、それらの項目それぞれの指定されたプロパティが読み込まれます。 次の例では、ワークシートの `CommentCollection` オブジェクトに含まれる各 `Comment` オブジェクトの `resolved` プロパティが読み込まれます。

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

ブックから情報を返す非同期 API のメソッドは、パラダイムと似たパターン `load` / `sync` を持っています。 たとえば、`TableCollection.getCount` はコレクション内のテーブルの数を取得します。 `getCount` は `ClientResult<number>` を返します。つまり、返される[`ClientResult`](/javascript/api/office/officeextension.clientresult?view=excel-js-online&preserve-view=true)の`value`プロパティは数値になります。 `context.sync()` が呼び出されるまで、スクリプトはその値にアクセスできません。 プロパティの読み込みと同様、`value` は、`sync` が呼び出されるまでは、ローカルの "空の" 値です。

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

## <a name="converting-async-scripts-to-the-current-model"></a>非同期スクリプトを現在のモデルに変換する

現在の API モデルでは、、 `load` 、または `sync` を使用しない `RequestContext` 。 これにより、スクリプトの書き込みと保守が非常に簡単になります。 古いスクリプトを変換するための最適なリソースは、 [スタック オーバーフローです](https://stackoverflow.com/questions/tagged/office-scripts)。 そこで、コミュニティに特定のシナリオに関するヘルプを求めることができます。 次のガイダンスは、実行する必要がある一般的な手順の概要を説明するのに役立ちます。

1. 新しいスクリプトを作成し、古い非同期コードをそのスクリプトにコピーします。 代わりにカレントを使用して、古 `main` いメソッドシグネチャを含め `function main(workbook: ExcelScript.Workbook)` ずにしてください。

2. すべての通話と `load` 呼び出しを `sync` 削除します。 これらはもはや必要ありません。

3. すべてのプロパティが削除されました。 これらのオブジェクトにアクセスするには、メソッドとメソッドを使用します。そのため、これらのプロパティ参照をメソッド呼び出 `get` `set` しに切り替える必要があります。 たとえば、次のようなプロパティ アクセスを使用してセルの塗りつぶしの色を設定する代わりに、次のようなメソッド `mySheet.getRange("A2:C2").format.fill.color = "blue";` を使用します。 `mySheet.getRange("A2:C2").getFormat().getFill().setColor("blue");`

4. コレクション クラスは配列に置き換えされています。 これらのコレクション クラスの and メソッドは、コレクションを所有するオブジェクトに移動されたので、参照を更新 `add` `get` する必要があります。 たとえば、ブックの最初のワークシートから "MyChart" という名前のグラフを取得するには、次のコードを使用します `workbook.getWorksheets()[0].getChart("MyChart");` 。 によって返 `[0]` される値の最初の値に `Worksheet[]` アクセスする点に注意してください `getWorksheets()` 。

5. 一部のメソッドは、わかりやすくするために名前が変更され、便宜上追加されています。 詳細については [、「Officeスクリプト API リファレンス」](/javascript/api/office-scripts/overview) を参照してください。

## <a name="office-scripts-async-api-reference-documentation"></a>Officeスクリプト非同期 API リファレンス ドキュメント

非同期 API は、アドインで使用Officeと同じです。リファレンス ドキュメントは、アドイン JavaScript API リファレンスOffice [Excel セクションに含まれています](/javascript/api/excel?view=excel-js-online&preserve-view=true)。
