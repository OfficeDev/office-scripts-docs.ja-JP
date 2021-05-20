---
title: 非同期 API を使用する古いOfficeスクリプトをサポートする
description: Officeスクリプト非同期 API のプライマーと、古いスクリプトのロード/同期パターンの使用方法。
ms.date: 05/17/2021
localization_priority: Normal
ms.openlocfilehash: 80a1c0dec5393d8882ddb37eea5f81ef23b1ebb1
ms.sourcegitcommit: 4687693f02fc90a57ba30c461f35046e02e6f5fb
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 05/19/2021
ms.locfileid: "52545076"
---
# <a name="support-older-office-scripts-that-use-the-async-apis"></a>非同期 API を使用する古いOfficeスクリプトをサポートする

この記事では、古いモデルの非同期 API を使用するスクリプトを維持および更新する方法について説明します。 これらの API は、現在標準の同期Officeスクリプト API と同じコア機能を持っていますが、スクリプトとワークブック間のデータ同期を制御するスクリプトが必要です。

> [!IMPORTANT]
> 非同期モデルは、現在の [API](scripting-fundamentals.md)モデルの実装前に作成されたスクリプトでのみ使用できます。 スクリプトは、作成時に持っている API モデルに永続的にロックされます。 つまり、古いスクリプトを新しいモデルに変換する場合は、新しいスクリプトを作成する必要があります。 現在のモデルは使いやすいので、変更を加える際には古いスクリプトを新しいモデルに更新することをお勧めします。 「 [非同期スクリプトを現在のモデルに変換する](#convert-async-scripts-to-the-current-model) 」セクションでは、この移行方法に関するアドバイスを提供します。

## <a name="older-main-function-signature"></a>古い `main` 関数シグネチャ

非同期 API を使用するスクリプトは、異なる機能を持ちます `main` 。 これは、最初の `async` パラメータとしてを持つ関数 `Excel.RequestContext` です。

```TypeScript
async function main(context: Excel.RequestContext) {
    // Your async Office Script
}
```

## <a name="context"></a>コンテキスト

`main` 関数は、`context` という名前の `Excel.RequestContext` パラメーターを受け入れます。 `context` は、スクリプトとブックの間のブリッジと見なすことができます。 スクリプトは、`context` オブジェクトを使用してブックにアクセスし、その `context` を使用してデータをやり取りします。

スクリプトと Excel は異なるプロセスや場所で実行されているため、`context` オブジェクトが必要になります。 スクリプトで、クラウドのブックに変更を加えたり、そのブックからデータをクエリしたりする必要があります。 `context` オブジェクトは、それらのトランザクションを管理します。

## <a name="sync-and-load"></a>同期と読み込み

スクリプトとブックは別の場所で実行されるため、両者の間でデータを転送するには時間がかかります。 非同期 API では、スクリプトとブックを同期する操作をスクリプトが明示的に呼び出すまで、コマンド `sync` がキューに入れられます。 スクリプトは、次のどちらかを実行することが必要になるまで、独立して動作できます。

- ブックからデータを読み取る (`load` 操作または [ClientResult](/javascript/api/office/officeextension.clientresult?view=excel-js-online&preserve-view=true) を返すメソッドの後)。
- ブックにデータを書き込む (通常はスクリプトが完了した結果)。

次の図に、スクリプトとブックの間の制御フローの例を示します。

:::image type="content" source="../images/load-sync.png" alt-text="スクリプトからブックに対する読み取りおよび書き込み操作を示す図":::

### <a name="sync"></a>同期

非同期スクリプトでブックからデータを読み取るか、ブックにデータを書き込む必要がある場合は、 `RequestContext.sync` 次のコード スニペットに示すようにメソッドを呼び出します。

```TypeScript
await context.sync();
```

> [!NOTE]
> スクリプトが終了すると、`context.sync()` が暗黙的に呼び出されます。

`sync` 操作が完了すると、ブックが更新され、スクリプトが指定した書き込み操作が反映されます。 書き込み操作は、Excelオブジェクトのプロパティを設定するか (たとえば) プロパティ `range.format.fill.color = "red"` を変更するメソッドを呼び出します (たとえば、 `range.format.autoFitColumns()` )。 また、`sync` 操作では、スクリプトが `load` 操作または `ClientResult` を返すメソッドを使用して要求したブックから任意の値が読み取られます (次のセクションを参照)。

ネットワークによっては、スクリプトとブックを同期するのに時間がかかる場合があります。 スクリプトの高速実行 `sync` を支援する呼び出しの数を最小限に抑えます。 それ以外の場合、非同期 API は標準の同期 API を高速にしません。

### <a name="load"></a>読み込み

非同期スクリプトは、ブックを読み取る前に、ブックからデータを読み込む必要があります。 ただし、ブック全体からデータを読み込むと、スクリプトの速度が大幅に低下します。 `load`このメソッドを使用すると、スクリプトは、ブックから取得する必要があるデータを明確に記述できます。

`load` メソッドは、すべての Excel オブジェクトで使用できます。 スクリプトでは、オブジェクトのプロパティを読み込んでからでなければ、それらを読み取ることができません。 これを行わないとエラーになります。

次の例では、`Range` オブジェクトを使用して、`load` メソッドでデータを読み込む方法を示します。

|目的 |コマンドの例 | 効果 |
|:--|:--|:--|
|1 つのプロパティを読み込む |`myRange.load("values");` | 単一のプロパティ (この例では、範囲内の値の 2 次元配列) を読み込みます。 |
|複数のプロパティを読み込む |`myRange.load("values, rowCount, columnCount");`| コンマで区切られたリストからすべてのプロパティ (この例では、値、行数、列数) を読み込みます。 |
|すべてを読み込む | `myRange.load();`|範囲のすべてのプロパティを読み込みます。 不要なデータを取得してスクリプトの速度を低下させることになるので、これは推奨されるソリューションではありません。 スクリプトのテスト中、またはオブジェクトからのすべてのプロパティが必要な場合にのみ、これを使用します。 |

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

また、コレクション全体のプロパティを読み込むこともできます。 非同期 API 内のすべてのコレクション オブジェクトには `items` 、そのコレクション内のオブジェクトを含む配列であるプロパティがあります。 `items` を `load` に対する階層呼び出し (`items\myProperty`) の最初に使用すると、それらの項目それぞれの指定されたプロパティが読み込まれます。 次の例では、ワークシートの `CommentCollection` オブジェクトに含まれる各 `Comment` オブジェクトの `resolved` プロパティが読み込まれます。

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

ブックから情報を返す非同期 API のメソッドは、 `load` / パラダイムと同様のパターンを持 `sync` ちます。 たとえば、`TableCollection.getCount` はコレクション内のテーブルの数を取得します。 `getCount` は `ClientResult<number>` を返します。つまり、返される[`ClientResult`](/javascript/api/office/officeextension.clientresult?view=excel-js-online&preserve-view=true)の`value`プロパティは数値になります。 `context.sync()` が呼び出されるまで、スクリプトはその値にアクセスできません。 プロパティの読み込みと同様、`value` は、`sync` が呼び出されるまでは、ローカルの "空の" 値です。

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

## <a name="convert-async-scripts-to-the-current-model"></a>非同期スクリプトを現在のモデルに変換する

現在の API モデルでは `load` 、 `sync` 、または を使用しません `RequestContext` 。 これにより、スクリプトの作成と保守が大幅に容易になります。 古いスクリプトを変換するための最良のリソースは、 [Microsoft Q&A](/answers/topics/office-scripts-dev.html)です。 そこで、特定のシナリオに関するヘルプをコミュニティに依頼できます。 次のガイダンスは、必要な一般的な手順の概要を説明するのに役立ちます。

1. 新しいスクリプトを作成し、古い非同期コードをコピーします。 古いメソッドシグネチャを含めるのではなく `main` 、現在のメソッドシグネチャを使用 `function main(workbook: ExcelScript.Workbook)` しないようにしてください。

2. すべての `load` `sync` 呼び出しを削除します。 彼らはもはや必要ではありません。

3. すべてのプロパティが削除されました。 これらのオブジェクトにアクセスするには `get` `set` 、メソッドとメソッドを使用してアクセスするので、これらのプロパティ参照をメソッド呼び出しに切り替える必要があります。 たとえば、次のようにプロパティ アクセスを使用してセルの塗りつぶしの色を設定する代わりに、 `mySheet.getRange("A2:C2").format.fill.color = "blue";` 次のようなメソッドを使用します。 `mySheet.getRange("A2:C2").getFormat().getFill().setColor("blue");`

4. コレクション クラスは配列に置き換えられました。 `add`これらのコレクション クラスの メソッドと メソッド `get` は、コレクションを所有するオブジェクトに移動されたので、参照を更新する必要があります。 たとえば、ブックの最初のワークシートから "MyChart" という名前のグラフを取得するには、次のコードを使用 `workbook.getWorksheets()[0].getChart("MyChart");` します。 `[0]`によって返されるの最初の値にアクセスする `Worksheet[]` に注意 `getWorksheets()` してください。

5. いくつかの方法は、わかりやすくするために名前が変更され、便宜上追加されました。 詳細については[、Officeスクリプト API リファレンス](/javascript/api/office-scripts/overview)を参照してください。

## <a name="office-scripts-async-api-reference-documentation"></a>Officeスクリプト非同期 API リファレンス ドキュメント

非同期 API は、アドインで使用されるOffice API と同等です。リファレンス ドキュメントは、アドインの[JavaScript API のリファレンス のExcelセクションOffice](/javascript/api/excel?view=excel-js-online&preserve-view=true)参照されています。
