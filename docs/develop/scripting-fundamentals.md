---
title: Excel on the web での Office スクリプトのスクリプトの基本事項
description: Office スクリプトを作成する前に理解しておくべきオブジェクト モデルの情報と他の基本事項について説明します。
ms.date: 01/27/2020
localization_priority: Priority
ms.openlocfilehash: 5a709c16e23c00ffc7ee7949a3cb11459dc2d530
ms.sourcegitcommit: d556aaefac80e55f53ac56b7f6ecbc657ebd426f
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 03/26/2020
ms.locfileid: "42978735"
---
# <a name="scripting-fundamentals-for-office-scripts-in-excel-on-the-web-preview"></a>Excel on the web での Office スクリプトのスクリプトの基本事項 (プレビュー)

この記事では、Office スクリプトの技術的な側面について説明します。 Excel オブジェクトどうしが連携する仕組みや、コード エディターがブックと同期する仕組みについて説明します。

[!INCLUDE [Preview note](../includes/preview-note.md)]

## <a name="object-model"></a>オブジェクト モデル

Excel API について理解するには、ブックの構成要素が互いにどのように関連しているかを理解する必要があります。

- **ブック** には、1 つ以上の **ワークシート** が含まれます。
- **ワークシート** では、**Range** オブジェクトを介してセルにアクセスできます。
- **Range** は、連続したセルのグループを表します。
- **Range** は、**表**、**グラフ**、**図形**、およびその他のデータ可視化や組織オブジェクトを作成して配置するために使用されます。
- **ワークシート** には、個々のシートに存在するデータ オブジェクトのコレクションが含まれます。
- **ブック** には、**ブック** 全体のデータ オブジェクト (**表** など) の一部のコレクションが含まれます。

### <a name="ranges"></a>範囲

範囲とは、ブック内の連続したセルのグループのことです。 スクリプトでは、範囲を定義するのに通常 A1 形式の表記が使用されます (例: **B3** は、行 **B**、列 **3** の単一のセルで、**C2:F4** は、行 **C** から **F**、列 **2** から **4** までのセル)。

範囲には `values`、`formulas`、`format` の 3 つの主要なプロパティがあります。 これらのプロパティで、セルの値、評価する数式、およびセルの視覚的な書式設定を取得または設定します。

#### <a name="range-sample"></a>サンプル範囲

次のサンプルで、売上記録の作成方法を示します。 このスクリプトは、`Range` オブジェクトを使用して、値、数式、書式を設定しています。

```TypeScript
async function main(context: Excel.RequestContext) {
  // Get the active worksheet.
  let sheet = context.workbook.worksheets.getActiveWorksheet();

  // Create the headers and format them to stand out.
  let headers = [
    ["Product", "Quantity", "Unit Price", "Totals"]
  ];
  let headerRange = sheet.getRange("B2:E2");
  headerRange.values = headers;
  headerRange.format.fill.color = "#4472C4";
  headerRange.format.font.color = "white";

  // Create the product data rows.
  let productData = [
    ["Almonds", 6, 7.5],
    ["Coffee", 20, 34.5],
    ["Chocolate", 10, 9.56],
  ];
  let dataRange = sheet.getRange("B3:D5");
  dataRange.values = productData;

  // Create the formulas to total the amounts sold.
  let totalFormulas = [
    ["=C3 * D3"],
    ["=C4 * D4"],
    ["=C5 * D5"],
    ["=SUM(E3:E5)"]
  ];
  let totalRange = sheet.getRange("E3:E6");
  totalRange.formulas = totalFormulas;
  totalRange.format.font.bold = true;

  // Display the totals as US dollar amounts.
  totalRange.numberFormat = [["$0.00"]];
}
```

このスクリプトを実行すると、現在のワークシートに次のデータが作成されます。

![値の行、数式の列、書式設定されたヘッダーを示す売上記録。](../images/range-sample.png)

### <a name="charts-tables-and-other-data-objects"></a>グラフ、表、およびその他のデータ オブジェクト

スクリプトを使用することにより、Excel 内でデータ構造やビジュアル化を作成および操作できます。 表とグラフの 2 つのオブジェクトが頻繁に使用されますが、API はピボットテーブル、図形、画像などもサポートしています。

#### <a name="creating-a-table"></a>表の作成

データが入力された範囲を使用することにより、表を作成します。 書式設定とテーブル コントロール (フィルターなど) が自動的に範囲に適用されます。

次のスクリプトでは、前のサンプルの範囲を使用して表を作成します。

```TypeScript
async function main(context: Excel.RequestContext) {
   let sheet = context.workbook.worksheets.getActiveWorksheet();
   sheet.tables.add("B2:E5", true);
}
```

前のデータを含むワークシート上でこのスクリプトを実行すると、次のテーブルが作成されます。

![前の売上記録から作成された表。](../images/table-sample.png)

#### <a name="creating-a-chart"></a>グラフの作成

グラフを作成すると、範囲内のデータを視覚化できます。 スクリプトでさまざまな種類のグラフを作成できます。いずれのグラフも、必要に応じてカスタマイズできます。

次のスクリプトで、3 つの品目の簡単な縦棒グラフが作成され、ワークシートの上端から 100 ピクセル下に配置されます。

```TypeScript
async function main(context: Excel.RequestContext) {
  let sheet = context.workbook.worksheets.getActiveWorksheet();
  let chart = sheet.charts.add(Excel.ChartType.columnStacked, sheet.getRange("B3:C5"));
  chart.top = 100;
}
```

前の表を含むワークシート上でこのスクリプトを実行すると、次のグラフが作成されます。

![前の売上記録の 3 つの品目の数量が表示されている縦棒グラフ。](../images/chart-sample.png)

### <a name="further-reading-on-the-object-model"></a>オブジェクト モデルに関する参考資料

「[Office スクリプト API リファレンス](/javascript/api/office-scripts/overview)」に、Office スクリプトで使用されるオブジェクトが包括的にまとめられています。 目次を使用して、詳細を確認したいクラスに移動できます。 よく参照されているページのいくつかを次に示します。

- [グラフ](/javascript/api/office-scripts/excel/excel.chart)
- [コメント](/javascript/api/office-scripts/excel/excel.comment)
- [PivotTable](/javascript/api/office-scripts/excel/excel.pivottable)
- [Range](/javascript/api/office-scripts/excel/excel.range)
- [範囲の形式](/javascript/api/office-scripts/excel/excel.rangeformat)
- [図形](/javascript/api/office-scripts/excel/excel.shape)
- [表](/javascript/api/office-scripts/excel/excel.table)
- [ブック](/javascript/api/office-scripts/excel/excel.workbook)
- [ワークシート](/javascript/api/office-scripts/excel/excel.worksheet)

## <a name="main-function"></a>`main` 関数

どの Office スクリプトにも、次のシグネチャで、`Excel.RequestContext` 型の定義を含む `main` 関数を含める必要があります。

```TypeScript
async function main(context: Excel.RequestContext) {
    // Your Excel Script
}
```

スクリプトを実行すると、`main` 関数の内部のコードが実行されます。 `main` は、スクリプト内の他の関数を呼び出すことができますが、関数に含まれていないコードは実行されません。

## <a name="context"></a>コンテキスト

`main` 関数は、`context` という名前の `Excel.RequestContext` パラメーターを受け入れます。 `context` は、スクリプトとブックの間のブリッジと見なすことができます。 スクリプトは、`context` オブジェクトを使用してブックにアクセスし、その `context` を使用してデータをやり取りします。

スクリプトと Excel は異なるプロセスや場所で実行されているため、`context` オブジェクトが必要になります。 スクリプトで、クラウドのブックに変更を加えたり、そのブックからデータをクエリしたりする必要があります。 `context` オブジェクトは、それらのトランザクションを管理します。

## <a name="sync-and-load"></a>同期と読み込み

スクリプトとブックは別の場所で実行されるため、両者の間でデータを転送するには時間がかかります。 スクリプトのパフォーマンスを向上させるため、スクリプトが明示的に `sync` 操作を呼び出してスクリプトとブックを同期するまで、コマンドはキューに登録されます。 スクリプトは、次のどちらかを実行することが必要になるまで、独立して動作できます。

- ブックからデータを読み取る (`load` 操作の後)。
- ブックにデータを書き込む (通常はスクリプトが完了した結果)。

次の図に、スクリプトとブックの間の制御フローの例を示します。

![スクリプトからブックに対して実行される読み取りおよび書き込み操作を示す図。](../images/load-sync.png)

### <a name="sync"></a>同期

スクリプトでブックに対するデータの読み取りや書き込みが必要になる場合、次のように `RequestContext.sync` メソッドを呼び出します。

```TypeScript
await context.sync();
```

> [!NOTE]
> スクリプトが終了すると、`context.sync()` が暗黙的に呼び出されます。

`sync` 操作が完了すると、ブックが更新され、スクリプトが指定した書き込み操作が反映されます。 書き込み操作とは、Excel オブジェクトに任意のプロパティを設定すること (`range.format.fill.color = "red"` など)、またはプロパティを変更するメソッドを呼び出すこと (`range.format.autoFitColumns()` など) を意味します。 また、`sync` 操作では、スクリプトが `load` 操作を使用して要求したブックから任意の値が読み取られます (次のセクションを参照)。

ネットワークによっては、スクリプトとブックを同期するのに時間がかかる場合があります。 スクリプトの実行速度を高めるため、`sync` 呼び出しは最小限に抑えることをお勧めします。  

### <a name="load"></a>読み込み

スクリプトでは、ブックからデータを読み込んでから、そのデータを読み取る必要があります。 しかし、ブック全体からデータを読み込むと、スクリプトの速度が大幅に低下します。 代わりに、`load` メソッドを使用すると、どのデータをブックから取得する必要があるかをスクリプトで具体的に指定できます。

`load` メソッドは、すべての Excel オブジェクトで使用できます。 スクリプトでは、オブジェクトのプロパティを読み込んでからでなければ、それらを読み取ることができません。 これに従わないと、エラーが発生します。

次の例では、`Range` オブジェクトを使用して、`load` メソッドでデータを読み込む方法を示します。

|目的 |コマンドの例 | 効果 |
|:--|:--|:--|
|1 つのプロパティを読み込む |`myRange.load("values");` | 単一のプロパティ (この例では、範囲内の値の 2 次元配列) を読み込みます。 |
|複数のプロパティを読み込む |`myRange.load("values, rowCount, columnCount");`| コンマで区切られたリストからすべてのプロパティ (この例では、値、行数、列数) を読み込みます。 |
|すべてを読み込む | `myRange.load();`|範囲のすべてのプロパティを読み込みます。 このソリューションは、不要なデータを取得することによりスクリプトの速度が低下するため、推奨されません。 スクリプトをテストする場合、またはオブジェクトのすべてのプロパティが必要な場合にのみ使用してください。 |

スクリプトでは、読み込まれた値を読み取る前に、`context.sync()` を呼び出す必要があります。

```TypeScript
let range = selectedSheet.getRange("A1:B3");
range.load ("rowCount"); // Load the property.
await context.sync(); // Synchronize with the workbook to get the property.
console.log(range.rowCount); // Read and log the property value (3).
```

また、コレクション全体のプロパティを読み込むこともできます。 どのコレクション オブジェクトにも、`items` プロパティがあります。これは、そのコレクションのオブジェクトを格納する配列です。 `items` を `load` に対する階層呼び出し (`items\myProperty`) の最初に使用すると、それらの項目それぞれの指定されたプロパティが読み込まれます。 次の例では、ワークシートの `CommentCollection` オブジェクトに含まれる各 `Comment` オブジェクトの `resolved` プロパティが読み込まれます。

```TypeScript
let comments = selectedSheet.comments;
comments.load("items/resolved"); // Load the `resolved` property from every comment in this collection.
await context.sync(); // Synchronize with the workbook to get the properties.
```

> [!TIP]
> Office スクリプトでのコレクションの使用方法の詳細については、[「Office スクリプトでの組み込みの JavaScript オブジェクトの使用」の「配列」セクション](javascript-objects.md#array)を参照してください。

## <a name="see-also"></a>関連項目

- [Excel on the web で Office スクリプトを記録、編集、作成する](../tutorials/excel-tutorial.md)
- [Excel on the web で Office スクリプトを使用してブックのデータを読み取る](../tutorials/excel-read-tutorial.md)
- [Office スクリプト API リファレンス](/javascript/api/office-scripts/overview)
- [Office スクリプトでの組み込みの JavaScript オブジェクトの使用](javascript-objects.md)
