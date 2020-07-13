---
title: Excel on the web での Office スクリプトのスクリプトの基本事項
description: Office スクリプトを作成する前に理解しておくべきオブジェクト モデルの情報と他の基本事項について説明します。
ms.date: 06/29/2020
localization_priority: Priority
ms.openlocfilehash: 9ea24f26052877bc70862c8a05321d588f409b11
ms.sourcegitcommit: 30750c4392db3ef057075a5702abb92863c93eda
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 07/01/2020
ms.locfileid: "44999303"
---
# <a name="scripting-fundamentals-for-office-scripts-in-excel-on-the-web-preview"></a>Excel on the web での Office スクリプトのスクリプトの基本事項 (プレビュー)

この記事では、Office スクリプトの技術的な側面について説明します。 Excel オブジェクトどうしが連携する仕組みや、コード エディターがブックと同期する仕組みについて説明します。

[!INCLUDE [Preview note](../includes/preview-note.md)]

## <a name="main-function"></a>`main` 関数

各 Office スクリプトには、最初のパラメーターとして `ExcelScript.Workbook` 型を持つ `main` 関数を含める必要があります。 関数が実行されると、Excel アプリケーションはブックを最初のパラメーターとして指定して、この `main` 関数を呼び出します。 そのため、スクリプトを記録した後、またはコード エディターで新しいスクリプトを作成した後に、`main` 関数の基本シグネチャを変更しないようにすることが重要です。

```typescript
function main(workbook: ExcelScript.Workbook) {
// Your code goes here
}
```

スクリプトを実行すると、`main` 関数の内部のコードが実行されます。 `main` は、スクリプト内の他の関数を呼び出すことができますが、関数に含まれていないコードは実行されません。

> [!CAUTION]
> `main` の関数が `async function main(context: Excel.RequestContext)` のように表示されている場合、スクリプトは従来の非同期 API モデルを使用しています。 前のスクリプトを現在の API モデルに変換する方法など、詳細については、[「Office スクリプトの非同期 API を使用して以前のスクリプトをサポートする」](excel-async-model.md) を参照してください。

## <a name="object-model"></a>オブジェクト モデル

スクリプトを作成するには、Office スクリプト API がどのように連携しているかを理解する必要があります。 ブックのコンポーネントには、相互に特定の関係があります。 多くの点で、これらの関係は Excel UI の関係と一致しています。

- **ブック** には、1 つ以上の **ワークシート** が含まれます。
- **ワークシート** では、**Range** オブジェクトを介してセルにアクセスできます。
- **Range** は、連続したセルのグループを表します。
- **Range** は、**表**、**グラフ**、**図形**、およびその他のデータ可視化や組織オブジェクトを作成して配置するために使用されます。
- **ワークシート** には、個々のシートに存在するデータ オブジェクトのコレクションが含まれます。
- **ブック** には、**ブック** 全体のデータ オブジェクト (**表** など) の一部のコレクションが含まれます。

### <a name="workbook"></a>ブック

すべてのスクリプトには、`main` 関数によって `Workbook` 型の `workbook` オブジェクトが提供されています。 これは、スクリプトが Excel ブックを操作するための最上位レベルのオブジェクトを表します。

次のスクリプトは、アクティブなワークシートをブックから取得し、その名前を記録します。

```typescript
function main(workbook: ExcelScript.Workbook) {
    // Get the active worksheet.
    let sheet = workbook.getActiveWorksheet();

    // Display the current worksheet's name.
    console.log(sheet.getName());
}
```

### <a name="ranges"></a>範囲

範囲とは、ブック内の連続したセルのグループのことです。 スクリプトでは、範囲を定義するのに通常 A1 形式の表記が使用されます (例: **B3** は、列 **B**、行 **3** の単一のセルで、**C2:F4** は、列 **C** から **F**、行 **2** から **4** までのセル)。

範囲には、値、数式、書式の 3 つの主要プロパティがあります。 これらのプロパティで、セルの値、評価する数式、およびセルの視覚的な書式設定を取得または設定します。 `getValues`、`getFormulas`、`getFormat` を介してアクセスします。 値と数式は、`setValues` と `setFormulas` で変更できますが、書式は、個別に設定されている複数の小さなオブジェクトから構成されている `RangeFormat` オブジェクトです。

範囲は、2 次元配列を使用して情報を管理します。 Office スクリプト フレームワークでこれらの配列を処理する方法の詳細については、[「Office スクリプトでの組み込み JavaScript オブジェクトの使用の範囲操作のセクション」](javascript-objects.md#working-with-ranges) を参照してください。

#### <a name="range-sample"></a>サンプル範囲

次のサンプルで、売上記録の作成方法を示します。 このスクリプトは、`Range` オブジェクトを使用して、値、数式、書式の一部を設定しています。

```TypeScript
function main(workbook: ExcelScript.Workbook) {
    // Get the active worksheet.
    let sheet = workbook.getActiveWorksheet();

    // Create the headers and format them to stand out.
    let headers = [["Product", "Quantity", "Unit Price", "Totals"]];
    let headerRange = sheet.getRange("B2:E2");
    headerRange.setValues(headers);
    headerRange.getFormat().getFill().setColor("#4472C4");
    headerRange.getFormat().getFont().setColor("white");

    // Create the product data rows.
    let productData = [
        ["Almonds", 6, 7.5],
        ["Coffee", 20, 34.5],
        ["Chocolate", 10, 9.56],
    ];
    let dataRange = sheet.getRange("B3:D5");
    dataRange.setValues(productData);

    // Create the formulas to total the amounts sold.
    let totalFormulas = [
        ["=C3 * D3"],
        ["=C4 * D4"],
        ["=C5 * D5"],
        ["=SUM(E3:E5)"],
    ];
    let totalRange = sheet.getRange("E3:E6");
    totalRange.setFormulas(totalFormulas);
    totalRange.getFormat().getFont().setBold(true);

    // Display the totals as US dollar amounts.
    totalRange.setNumberFormat("$0.00");
}
```

このスクリプトを実行すると、現在のワークシートに次のデータが作成されます。

![値の行、数式の列、書式設定されたヘッダーを示す売上記録。](../images/range-sample.png)

### <a name="charts-tables-and-other-data-objects"></a>グラフ、表、およびその他のデータ オブジェクト

スクリプトを使用することにより、Excel 内でデータ構造やビジュアル化を作成および操作できます。 表とグラフの 2 つのオブジェクトが頻繁に使用されますが、API はピボットテーブル、図形、画像などもサポートしています。 これらはコレクションに格納され、この記事の後半で説明します。

#### <a name="creating-a-table"></a>表の作成

データが入力された範囲を使用することにより、表を作成します。 書式設定とテーブル コントロール (フィルターなど) が自動的に範囲に適用されます。

次のスクリプトでは、前のサンプルの範囲を使用して表を作成します。

```TypeScript
function main(workbook: ExcelScript.Workbook) {
    // Get the active worksheet.
    let sheet = workbook.getActiveWorksheet();

    // Add a table that has headers using the data from B2:E5.
    sheet.addTable("B2:E5", true);
}
```

前のデータを含むワークシート上でこのスクリプトを実行すると、次のテーブルが作成されます。

![前の売上記録から作成された表。](../images/table-sample.png)

#### <a name="creating-a-chart"></a>グラフの作成

グラフを作成すると、範囲内のデータを視覚化できます。 スクリプトでさまざまな種類のグラフを作成できます。いずれのグラフも、必要に応じてカスタマイズできます。

次のスクリプトで、3 つの品目の簡単な縦棒グラフが作成され、ワークシートの上端から 100 ピクセル下に配置されます。

```TypeScript
function main(workbook: ExcelScript.Workbook) {
    // Get the active worksheet.
    let sheet = workbook.getActiveWorksheet();

    // Create a column chart using the data from B3:C5.
    let chart = sheet.addChart(
        ExcelScript.ChartType.columnStacked,
        sheet.getRange("B3:C5")
    );

    // Set the margin of the chart to be 100 pixels from the top of the screen.
    chart.setTop(100);
}
```

前の表を含むワークシート上でこのスクリプトを実行すると、次のグラフが作成されます。

![前の売上記録の 3 つの品目の数量が表示されている縦棒グラフ。](../images/chart-sample.png)

### <a name="collections-and-other-object-relations"></a>コレクションとその他のオブジェクトの関係

子オブジェクトには、その親オブジェクトを通じてアクセスできます。 たとえば、`Workbook` オブジェクトから `Worksheets` を読み取ることができます。 親クラスには、(`Workbook.getWorksheets()` や `Workbook.getWorksheet(name)` など) 関連する `get` メソッドがあります。 単一の `get` メソッドは、単一のオブジェクトを返し、特定のオブジェクト (ワークシート名など) の ID または名前を要求します。 複数の `get` メソッドは、オブジェクト コレクション全体を配列として返します。 コレクションが空の場合、空の配列 (`[]`) が返されます。

コレクションを取得したら、`length` を取得したり、`for`、`for..of`、`while` ループを使用して反復処理を行ったり、`map`や `forEach` などの TypeScript 配列メソッドを使用したりするなど、通常の配列操作を利用できます。 配列のインデックス値を使用して、コレクション内の個々のオブジェクトにアクセスすることもできます。 たとえば、`workbook.getTables()[0]` はコレクション内の最初のテーブルを返します。 Office スクリプト フレームワークで組み込みの配列機能を使用する方法の詳細については、[「Office スクリプトでの組み込み JavaScript オブジェクトの使用のコレクション操作のセクション」](javascript-objects.md#working-with-collections) を参照してください。

次のスクリプトは、ブック内のすべてのテーブルを取得します。 これにより、ヘッダーが表示され、フィルター ボタンが表示され、テーブル スタイルが「TableStyleLight1」に設定されていることを確認します。

```typescript
function main(workbook: ExcelScript.Workbook) {
  /* Get table collection */
  const tables = workbook.getTables();
  /* Set table formatting properties */
  tables.forEach(table => {
    table.setShowHeaders(true);
    table.setShowFilterButton(true);
    table.setPredefinedTableStyle("TableStyleLight1");
  })
}
```

#### <a name="adding-excel-objects-with-a-script"></a>スクリプトを使用して Excel オブジェクトを追加する

親オブジェクトで使用可能な対応する `add` メソッドを呼び出すことにより、プログラムでテーブルやグラフなどのドキュメント オブジェクトを追加できます。

> [!NOTE]
> コレクション配列にオブジェクトを手動で追加しないでください。 親オブジェクトに `add` メソッドを使用します。たとえば、`Worksheet.addTable` メソッドを使用して、`Worksheet` に `Table` を追加します。

次のスクリプトは、ブック内の最初のワークシートに Excel のテーブルを作成します。 作成されたテーブルは、`addTable` メソッドによって返されます。

```typescript
function main(workbook: ExcelScript.Workbook) {
    // Get the first worksheet.
    let sheet = workbook.getWorksheets()[0];

    // Add a table that uses the data in C3:G10.
    let table = sheet.addTable(
      "C3:G10",
       true /* True because the table has headers. */
    );
}
```

## <a name="removing-excel-objects-with-a-script"></a>スクリプトを使用して Excel オブジェクトを削除する

オブジェクトを削除するには、オブジェクトの `delete` メソッドを呼び出します。

> [!NOTE]
> オブジェクトを追加する場合と同様に、コレクション配列からオブジェクトを手動で削除しないでください。 コレクション型のオブジェクトの `delete` メソッドを使用します。 たとえば、`Table.delete` を使用して `Worksheet` から `Table` を削除します。

次のスクリプトは、ブック内の最初のワークシートを削除します。

```typescript
function main(workbook: ExcelScript.Workbook) {
    // Get first worksheet.
    let sheet = workbook.getWorksheets()[0];

    // Remove that worksheet from the workbook.
    sheet.delete();
}
```

### <a name="further-reading-on-the-object-model"></a>オブジェクト モデルに関する参考資料

「[Office スクリプト API リファレンス](/javascript/api/office-scripts/overview)」に、Office スクリプトで使用されるオブジェクトが包括的にまとめられています。 目次を使用して、詳細を確認したいクラスに移動できます。 よく参照されているページのいくつかを次に示します。

- [グラフ](/javascript/api/office-scripts/excelscript/excelscript.chart)
- [コメント](/javascript/api/office-scripts/excelscript/excelscript.comment)
- [PivotTable](/javascript/api/office-scripts/excelscript/excelscript.pivottable)
- [Range](/javascript/api/office-scripts/excelscript/excelscript.range)
- [範囲の形式](/javascript/api/office-scripts/excelscript/excelscript.rangeformat)
- [図形](/javascript/api/office-scripts/excelscript/excelscript.shape)
- [表](/javascript/api/office-scripts/excelscript/excelscript.table)
- [ブック](/javascript/api/office-scripts/excelscript/excelscript.workbook)
- [ワークシート](/javascript/api/office-scripts/excelscript/excelscript.worksheet)

## <a name="see-also"></a>関連項目

- [Excel on the web で Office スクリプトを記録、編集、作成する](../tutorials/excel-tutorial.md)
- [Excel on the web で Office スクリプトを使用してブックのデータを読み取る](../tutorials/excel-read-tutorial.md)
- [Office スクリプト API リファレンス](/javascript/api/office-scripts/overview)
- [Office スクリプトでの組み込みの JavaScript オブジェクトの使用](javascript-objects.md)
