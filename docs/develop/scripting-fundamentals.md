---
title: Excel on the web での Office スクリプトのスクリプトの基本事項
description: Office スクリプトを作成する前に理解しておくべきオブジェクト モデルの情報と他の基本事項について説明します。
ms.date: 05/24/2021
localization_priority: Priority
ms.openlocfilehash: 629e816ea988d6b8ffe5264c701e3a1eba6c6feb
ms.sourcegitcommit: 90ca8cdf30f2065f63938f6bb6780d024c128467
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 05/25/2021
ms.locfileid: "52639895"
---
# <a name="scripting-fundamentals-for-office-scripts-in-excel-on-the-web"></a>Excel on the web での Office スクリプトのスクリプトの基本事項

この記事では、Office スクリプトの技術的な側面について説明します。 Excel オブジェクトどうしが連携する仕組みや、コード エディターがブックと同期する仕組みについて説明します。

## <a name="typescript-the-language-of-office-scripts"></a>TypeScript: オフィス スクリプトの言語

オフィス スクリプトは [TypeScript](https://www.typescriptlang.org/docs/home.html) で書かれており、[JavaScript](https://developer.mozilla.org/docs/Web/JavaScript) のスーパーセットです。 JavaScript に慣れている場合は、コードの多くが両言語で共通しているため、知識を引き継ぐことができます。 Office スクリプトのコーディング作業を始める前に、初心者レベルのプログラミング知識を身に付けておくことをお勧めします。 以下のリソースは、Office スクリプトのコーディング面を理解するのに役立ちます。

[!INCLUDE [Recommended coding resources](../includes/coding-basics-references.md)]

## <a name="main-function-the-scripts-starting-point"></a>`main` 機能: スクリプトの開始点

各スクリプトには、最初のパラメーターとして `ExcelScript.Workbook` 型の `main` 関数を含める必要があります。 関数を実行すると、Excel アプリケーションはブックを最初のパラメーターとして指定して、この `main` 関数を呼び出します。 `ExcelScript.Workbook` は常に最初のパラメータである必要があります。

```TypeScript
function main(workbook: ExcelScript.Workbook) {
  // Your code goes here
}
```

スクリプトを実行すると、`main` 関数の内部のコードが実行されます。 `main` は、スクリプト内の他の関数を呼び出すことができますが、関数に含まれていないコードは実行されません。 スクリプトは、他の Office スクリプトを呼び出すことはできません。

[Power Automate](https://flow.microsoft.com) では、スクリプトをフローで接続することができます。 スクリプトとフローの間のデータの受け渡しは、`main` メソッドのパラメーターと戻り値を介して行われます。 Office スクリプトと Power Automate を統合する方法については、 [「Power Automate で Office スクリプトを実行する」](power-automate-integration.md)で詳しく説明しています。

## <a name="object-model-overview"></a>オブジェクト モデルの概要

スクリプトを作成するには、Office スクリプト API がどのように連携しているかを理解する必要があります。 ブックのコンポーネントには、相互に特定の関係があります。 多くの点で、これらの関係は Excel UI の関係と一致しています。

- **ブック** には、1 つ以上の **ワークシート** が含まれます。
- **ワークシート** では、**Range** オブジェクトを介してセルにアクセスできます。
- **Range** は、連続したセルのグループを表します。
- **Range** は、**表**、**グラフ**、**図形**、およびその他のデータ可視化や組織オブジェクトを作成して配置するために使用されます。
- **ワークシート** には、個々のシートに存在するデータ オブジェクトのコレクションが含まれます。
- **ブック** には、**ブック** 全体のデータ オブジェクト (**表** など) の一部のコレクションが含まれます。

## <a name="workbook"></a>ブック

すべてのスクリプトには、`main` 関数によって `Workbook` 型の `workbook` オブジェクトが提供されています。 これは、スクリプトが Excel ブックを操作するための最上位レベルのオブジェクトを表します。

次のスクリプトは、アクティブなワークシートをブックから取得し、その名前を記録します。

```TypeScript
function main(workbook: ExcelScript.Workbook) {
    // Get the active worksheet.
    let sheet = workbook.getActiveWorksheet();

    // Display the current worksheet's name.
    console.log(sheet.getName());
}
```

## <a name="ranges"></a>範囲

範囲とは、ブック内の連続したセルのグループのことです。 スクリプトでは、範囲を定義するのに通常 A1 形式の表記が使用されます (例: **B3** は、列 **B**、行 **3** の単一のセルで、**C2:F4** は、列 **C** から **F**、行 **2** から **4** までのセル)。

範囲には、値、数式、書式の 3 つの主要プロパティがあります。 これらのプロパティで、セルの値、評価する数式、およびセルの視覚的な書式設定を取得または設定します。 `getValues`、`getFormulas`、`getFormat` を介してアクセスします。 値と数式は、`setValues` と `setFormulas` で変更できますが、書式は、個別に設定されている複数の小さなオブジェクトから構成されている `RangeFormat` オブジェクトです。

範囲は、2 次元配列を使用して情報を管理します。 Office スクリプト フレームワークでの配列の扱いについては、「[範囲での作業](javascript-objects.md#work-with-ranges)」を参照してください。

### <a name="range-sample"></a>サンプル範囲

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
        ["Chocolate", 10, 9.54],
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

:::image type="content" source="../images/range-sample.png" alt-text="値の行、数式の列、フォーマットされたヘッダーを含む売上記録を含むワークシート":::

### <a name="the-types-of-range-values"></a>レンジ値の種類

各セルには値があります。 この値は、セルに入力された基本的な値であり、Excel で表示されるテキストとは異なる場合があります。 たとえば、セルに日付として "2021 年 5 月 2 日" が表示されていても、実際の値は「44318」であることがあります。 この表示は、数値表示形式で変更できますが、セル内の実際の値やタイプは、新しい値が設定されたときにのみ変更されます。

セルの値を使用する場合には、セルや範囲からどのような値を得ることを期待しているのかを TypeScript に伝達することが重要です。 セルには、次のいずれかのタブの種類を選択します: `string`、`number`、または `boolean`。 スクリプトが返された値をこれらの型の 1 つとして処理するためには、その型を宣言する必要があります。

次のスクリプトは、前のサンプルのテーブルから平均価格を取得します。 コード `priceRange.getValues() as number[][]` を確認します。 この[アサート](https://www.typescriptlang.org/docs/handbook/2/everyday-types.html#type-assertions)は、範囲の値の型が `number[][]` であることを主張します。 この配列のすべての値は、スクリプトで数字として処理されます。

```TypeScript
function main(workbook: ExcelScript.Workbook) {
  // Get the active worksheet.
  let sheet = workbook.getActiveWorksheet();

  // Get the "Unit Price" column. 
  // The result of calling getValues is declared to be a number[][] so that we can perform arithmetic operations.
  let priceRange = sheet.getRange("D3:D5");
  let prices = priceRange.getValues() as number[][];

  // Get the average price.
  let totalPrices = 0;
  prices.forEach((price) => totalPrices += price[0]);
  let averagePrice = totalPrices / prices.length;
  console.log(averagePrice);
}
```

## <a name="charts-tables-and-other-data-objects"></a>グラフ、表、およびその他のデータ オブジェクト

スクリプトを使用することにより、Excel 内でデータ構造やビジュアル化を作成および操作できます。 表とグラフの 2 つのオブジェクトが頻繁に使用されますが、API はピボットテーブル、図形、画像などもサポートしています。 これらはコレクションに格納され、この記事の後半で説明します。

### <a name="create-a-table"></a>テーブルを作成する

データ入力範囲を使ってテーブルを作成します。書式設定とテーブル コントロール (フィルターなど) が自動的に範囲に適用されます。

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

:::image type="content" source="../images/table-sample.png" alt-text="前の売上記録から作成された表を含むワークシート":::

### <a name="create-a-chart"></a>グラフの作成

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

:::image type="content" source="../images/chart-sample.png" alt-text="前の売上記録の 3 つの品目の数量が表示されている縦棒グラフ。":::

## <a name="collections"></a>コレクション

Excel オブジェクトは、1 つ以上の同じ種類のオブジェクトのコレクションがある場合、それらを配列に格納します。 たとえば、`Workbook` オブジェクトには `Worksheet[]` が含まれます。 この配列は `Workbook.getWorksheets()` メソッドでアクセスします。 複数の `get` メソッド (`Worksheet.getCharts()` など) は、オブジェクト コレクション全体を配列として返します。 このパターンは、Office スクリプトの API 全体で見ることができます。たとえば、`Worksheet` オブジェクトには `getTables()` メソッドがあり、`Table[]` を返し、`Table` オブジェクトには `getColumns()` メソッドがあり、`TableColumn[]` を返すといったことです。

返された配列は通常の配列なので、スクリプトでは通常の配列操作がすべて可能です。 配列のインデックス値を使用して、コレクション内の個々のオブジェクトにアクセスすることもできます。 たとえば、`workbook.getTables()[0]` はコレクション内の最初のテーブルを返します。 Office スクリプト フレームワークで組み込みの配列機能を使用する方法については、「[コレクションでの作業](javascript-objects.md#work-with-collections)」を参照してください。 

個々のオブジェクトには、`get` メソッドを通してコレクションからアクセスします。 単一の `get` メソッド (`Worksheet.getTable(name)` など) は、単一のオブジェクトを返し、特定のオブジェクトの ID または名前を要求します。 この ID や名前は通常、スクリプトや Excel の UI で設定します。

次のスクリプトはブック内のすべてのテーブルを取得します。これにより、ヘッダーが表示され、フィルター ボタンが表示され、テーブル スタイルが "TableStyleLight1" に設定されます。

```TypeScript
function main(workbook: ExcelScript.Workbook) {
  // Get the table collection.
  let tables = workbook.getTables();

  // Set the table formatting properties for every table.
  tables.forEach(table => {
    table.setShowHeaders(true);
    table.setShowFilterButton(true);
    table.setPredefinedTableStyle("TableStyleLight1");
  })
}
```

## <a name="add-excel-objects-with-a-script"></a>スクリプトを使用して Excel オブジェクトを追加する

親オブジェクトで使用可能な対応する `add` メソッドを呼び出すことにより、プログラムでテーブルやグラフなどのドキュメント オブジェクトを追加できます。

> [!IMPORTANT]
> コレクション配列にオブジェクトを手動で追加しないでください。 親オブジェクトに `add` メソッドを使用します。たとえば、`Worksheet.addTable` メソッドを使用して、`Worksheet` に `Table` を追加します。

次のスクリプトは、ブック内の最初のワークシートに Excel のテーブルを作成します。 作成されたテーブルは、`addTable` メソッドによって返されます。

```TypeScript
function main(workbook: ExcelScript.Workbook) {
    // Get the first worksheet.
    let sheet = workbook.getWorksheets()[0];

    // Add a table that uses the data in A1:G10.
    let table = sheet.addTable(
      "A1:G10",
       true /* True because the table has headers. */
    );
    
    // Give the table a name for easy reference in other scripts.
    table.setName("MyTable");
}
```

> [!TIP]
> ほとんどの Excel オブジェクトには `setName` メソッドがあります。 これにより、スクリプトの後半や、同じワークブックを扱う他のスクリプトで、Excel オブジェクトに簡単にアクセスできるようになります。

### <a name="verify-an-object-exists-in-the-collection"></a>コレクションにオブジェクトが存在することを確認する

スクリプトでは、続行する前にテーブルなどのオブジェクトが存在するかどうかを確認する必要があります。 スクリプトや Excel の UI で与えられた名前を使って、必要なオブジェクトを特定し、それに応じて行動します。 `get` メソッドは、要求されたオブジェクトがコレクションに存在しない場合、`undefined` を返します。

次のスクリプトは、"MyTable" という名前のテーブルを要求し、`if...else` ステートメントを使用してテーブルが見つかったかどうか確認します。

```TypeScript
function main(workbook: ExcelScript.Workbook) {
  // Get the table named "MyTable".
  let myTable = workbook.getTable("MyTable");

  // If the table is in the workbook, myTable will have a value.
  // Otherwise, the variable will be undefined and go to the else clause.
  if (myTable) {
    let worksheetName = myTable.getWorksheet().getName();
    console.log(`MyTable is on the ${worksheetName} worksheet`);
  } else {
    console.log(`MyTable is not in the workbook.`);
  }
}
```

Office スクリプトで一般的なパターンは、スクリプトを実行するたびに表やグラフなどのオブジェクトを再作成することです。 以前のデータが不要な場合は、新しいオブジェクトを作成する前に以前のオブジェクトを削除するのがよいでしょう。 これにより、他のユーザーによってもたらされた名前の競合やその他の相違を避けることができます。

次のスクリプトは、"MyTable" という名前のテーブルがあればそれを削除し、同じ名前の新しいテーブルを追加します。

```TypeScript
function main(workbook: ExcelScript.Workbook) {
  // Get the table named "MyTable" from the first worksheet.
  let sheet = workbook.getWorksheets()[0];
  let tableName = "MyTable";
  let oldTable = sheet.getTable(tableName);

  // If the table exists, remove it.
  if (oldTable) {
    oldTable.delete();
  }

  // Add a new table with the same name.
  let newTable = sheet.addTable("A1:G10", true);
  newTable.setName(tableName);
}
```

## <a name="remove-excel-objects-with-a-script"></a>スクリプトを使用して Excel オブジェクトを削除する

オブジェクトを削除するには、オブジェクトの `delete` メソッドを呼び出します。

> [!NOTE]
> オブジェクトを追加する場合と同様に、コレクション配列からオブジェクトを手動で削除しないでください。 コレクション型のオブジェクトの `delete` メソッドを使用します。 たとえば、`Table.delete` を使用して `Worksheet` から `Table` を削除します。

次のスクリプトは、ブック内の最初のワークシートを削除します。

```TypeScript
function main(workbook: ExcelScript.Workbook) {
    // Get first worksheet.
    let sheet = workbook.getWorksheets()[0];

    // Remove that worksheet from the workbook.
    sheet.delete();
}
```

## <a name="further-reading-on-the-object-model"></a>オブジェクト モデルに関する参考資料

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
- [Office スクリプトでのベスト プラクティス](best-practices.md)
