---
title: Office スクリプトでのピボットテーブルの操作
description: Office Scripts JavaScript API でピボットテーブルのオブジェクト モデルについて説明します。
ms.date: 04/20/2022
ms.localizationpriority: medium
ms.openlocfilehash: 579f94140214674912c9610e707123924e4aef18
ms.sourcegitcommit: 4e3d3aa25fe4e604b806fbe72310b7a84ee72624
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 04/27/2022
ms.locfileid: "65077091"
---
# <a name="work-with-pivottables-in-office-scripts"></a>Office スクリプトでのピボットテーブルの操作

ピボットテーブルを使用すると、大量のデータコレクションをすばやく分析できます。 その能力には複雑さが伴います。 Office スクリプト API を使用すると、ニーズに合わせてピボットテーブルをカスタマイズできますが、API セットのスコープを使用すると、作業を開始することが困難になります。 この記事では、一般的なピボットテーブル タスクを実行する方法と、重要なクラスとメソッドについて説明します。

> [!NOTE]
> API で使用される用語のコンテキストを理解するには、最初にExcelのピボットテーブルドキュメントを参照してください。 [ワークシート データを分析するピボットテーブルの作成](https://support.microsoft.com/office/a9a84538-bfe9-40a9-a8e9-f99134456576)から始めます。

## <a name="object-model"></a>オブジェクト モデル

:::image type="content" source="../images/pivottable-object-model.png" alt-text="ピボットテーブルを操作するときに使用されるクラス、メソッド、プロパティの簡略化された図。":::

[ピボットテーブル](/javascript/api/office-scripts/excelscript/excelscript.pivottable)は、Office Scripts API のピボットテーブルの中央オブジェクトです。

- [Workbook](/javascript/api/office-scripts/excelscript/excelscript.workbook) オブジェクトには、すべての[ピボットテーブル](/javascript/api/office-scripts/excelscript/excelscript.pivottable)のコレクションがあります。 各 [ワークシート](/javascript/api/office-scripts/excelscript/excelscript.worksheet) には、そのシートのローカルであるピボットテーブル コレクションも含まれています。
- [ピボットテーブル](/javascript/api/office-scripts/excelscript/excelscript.pivottable)には [PivotHierarchies が含まれています](/javascript/api/office-scripts/excelscript/excelscript.pivothierarchy)。 階層は、テーブル内の列と考えることができます。
- [PivotHierarchies](/javascript/api/office-scripts/excelscript/excelscript.pivothierarchy) は、行または列 ([RowColumnPivotHierarchy](/javascript/api/office-scripts/excelscript/excelscript.rowcolumnpivothierarchy))、データ ([DataPivotHierarchy](/javascript/api/office-scripts/excelscript/excelscript.datapivothierarchy))、またはフィルター ([FilterPivotHierarchy](/javascript/api/office-scripts/excelscript/excelscript.filterpivothierarchy)) として追加できます。
- 各 [PivotHierarchy](/javascript/api/office-scripts/excelscript/excelscript.pivothierarchy) には、ピボットフィールドが 1 つだけ含 [まれています](/javascript/api/office-scripts/excelscript/excelscript.pivotfield)。 Excelの外部のピボットテーブル構造には、階層ごとに複数のフィールドが含まれる場合があるため、この設計は将来のオプションをサポートするために存在します。 Office スクリプトの場合、フィールドと階層は同じ情報にマップされます。
- [PivotField](/javascript/api/office-scripts/excelscript/excelscript.pivotfield) には、複数の [PivotItem が含まれています](/javascript/api/office-scripts/excelscript/excelscript.pivotitem)。 各 PivotItem は、フィールド内の一意の値です。 各項目は、テーブル列の値と考えてください。 フィールドがデータに使用されている場合は、項目の集計値 (合計など) を指定することもできます。
- [PivotLayout は](/javascript/api/office-scripts/excelscript/excelscript.pivotlayout)、[PivotFields](/javascript/api/office-scripts/excelscript/excelscript.pivotfield) と [PivotItems](/javascript/api/office-scripts/excelscript/excelscript.pivotitem) の表示方法を定義します。
- [PivotFilter は、](/javascript/api/office-scripts/excelscript/excelscript.pivotfilters) 異なる条件を使用して [ピボットテーブル](/javascript/api/office-scripts/excelscript/excelscript.pivottable) からデータをフィルター処理します。

これらのリレーションシップの実際の動作を確認します。 次のデータでは、さまざまなファームからの果樹の売上について説明します。 この記事のすべての例のベースです。 <a href="pivottable-sample.xlsx">pivottable-sample.xlsx</a>を使用してフォローします。

:::image type="content" source="../images/pivottable-raw-data.png" alt-text="さまざまなファームのさまざまな種類の果樹販売のコレクション。":::

## <a name="create-a-pivottable-with-fields"></a>フィールドを含むピボットテーブルを作成する

ピボットテーブルは、既存のデータへの参照を使用して作成されます。 範囲とテーブルの両方をピボットテーブルのソースにすることができます。 また、ブックに存在する場所も必要です。 ピボットテーブルのサイズは動的であるため、変換先範囲の左上隅のみが指定されます。

次のコード スニペットは、データ範囲に基づいてピボットテーブルを作成します。 ピボットテーブルには階層がないため、データはまだグループ化されていません。

```typescript
  const dataSheet = workbook.getWorksheet("Data");
  const pivotSheet = workbook.getWorksheet("Pivot");

  const farmPivot = pivotSheet.addPivotTable(
    "Farm Pivot", /* The name of the PivotTable. */
    dataSheet.getUsedRange(), /* The source data range. */
    pivotSheet.getRange("A1") /* The location to put the new PivotTable. */);
```

:::image type="content" source="../images/pivottable-empty.png" alt-text="階層のない &quot;ファーム ピボット&quot; という名前のピボットテーブル。":::

### <a name="hierarchies-and-fields"></a>階層とフィールド

ピボットテーブルは階層によって編成されます。 これらの階層は、特定の種類の階層として追加されたときにデータをピボットするために使用されます。 階層には 4 種類あります。

- **行**: 水平方向の行に項目を表示します。
- **列**: 垂直方向の列に項目を表示します。
- **データ**: 行と列に基づいて値の集計を表示します。
- **フィルター**: ピボットテーブルからアイテムを追加または削除します。

ピボットテーブルには、これらの特定の階層に割り当てられているフィールドの数または数を指定できます。 ピボットテーブルには、集計された数値データを表示するために少なくとも 1 つのデータ階層と、その概要をピボットする行または列が少なくとも 1 つ必要です。 次のコード スニペットは、2 つの行階層と 2 つのデータ階層を追加します。

```typescript
  farmPivot.addRowHierarchy(farmPivot.getHierarchy("Farm"));
  farmPivot.addRowHierarchy(farmPivot.getHierarchy("Type"));
  farmPivot.addDataHierarchy(farmPivot.getHierarchy("Crates Sold at Farm"));
  farmPivot.addDataHierarchy(farmPivot.getHierarchy("Crates Sold Wholesale"));
```

:::image type="content" source="../images/pivottable-data-hierarchy.png" alt-text="彼らが出てきたファームに基づいて異なる果樹の総売上を示すピボットテーブル。":::

## <a name="layout-ranges"></a>レイアウト範囲

ピボットテーブルの各部分は、範囲にマップされます。 これにより、スクリプトで後で使用したり、[Power Automate フロー](power-automate-integration.md)で返したりするために、ピボットテーブルからデータを取得できます。 これらの範囲には、. から取得した [PivotLayout](/javascript/api/office-scripts/excelscript/excelscript.pivotlayout) オブジェクトを `PivotTable.getLayout()`介してアクセスされます。 次の図は、次のメソッド `PivotLayout`によって返される範囲を示しています。

:::image type="content" source="../images/pivottable-layout-breakdown.png" alt-text="レイアウトの get range 関数によって返されるピボットテーブルのセクションを示す図。":::

## <a name="filters-and-slicers"></a>フィルターとスライサー

ピボットテーブルをフィルター処理するには、3 つの方法があります。

- [FilterPivotHierarchies](/javascript/api/office-scripts/excelscript/excelscript.filterpivothierarchy)
- [PivotFilters](/javascript/api/office-scripts/excelscript/excelscript.pivotfilters)
- [Slicers](/javascript/api/office-scripts/excelscript/excelscript.slicer)

### <a name="filterpivothierarchies"></a>FilterPivotHierarchies

`FilterPivotHierarchies` 階層を追加して、すべてのデータ行をフィルター処理します。 除外されたアイテムを含む行はすべて、ピボットテーブルとその概要から除外されます。 これらのフィルターは項目に基づいているため、個別の値でのみ機能します。 "分類" がサンプルのフィルター階層である場合、ユーザーはフィルターに対して "Organic" と "従来" の値を選択できます。 同様に、"Crates Sold Wholesale" が選択されている場合、フィルター オプションは数値範囲ではなく、120 や 150 などの個々の数値になります。

`FilterPivotHierarchies` は、すべての値が選択された状態で作成されます。 つまり、ユーザーが手動でフィルター コントロールを操作するか、または .. に属する`FilterPivotHierarchy`フィールドに対して a `PivotManualFilter` が設定されるまで、何もフィルター処理されません。

次のコード スニペットは、フィルター階層として "分類" を追加します。

```typescript
  farmPivot.addFilterHierarchy(farmPivot.getHierarchy("Classification"));
```

:::image type="content" source="../images/pivottable-filter-hierarchy.png" alt-text="ピボットテーブルに &quot;分類&quot; を使用するフィルター コントロール。":::

### <a name="pivotfilters"></a>PivotFilters

オブジェクトは `PivotFilters` 、1 つのフィールドに適用されるフィルターのコレクションです。 各階層には 1 つのフィールドがあるため、フィルターを適用するときは常に最初の `PivotHierarchy.getFields()` フィールドを使用する必要があります。 フィルターの種類は 4 つあります。

- **日付フィルター**: 予定表の日付ベースのフィルター処理。
- **ラベル フィルター**: テキスト比較フィルター。
- **手動フィルター**: カスタム入力フィルター。
- **値フィルター**: 数値比較フィルター。 これにより、関連付けられた階層内の項目と、指定したデータ階層内の値が比較されます。

通常、4 種類のフィルターのうち 1 つだけが作成され、フィールドに適用されます。 スクリプトで互換性のないフィルターを使用しようとすると、"引数が無効か、または形式が正しくありません" というテキストでエラーがスローされます。

次のコード スニペットでは、2 つのフィルターを追加します。 1 つ目は、既存の "分類" フィルター階層内のアイテムを選択する手動フィルターです。 2 番目のフィルターは、"Crates Sold Wholesale" が 300 未満のファームを削除します。 これにより、元のデータの個々の行ではなく、それらのファームの "Sum" が除外されることに注意してください。

```typescript
  const classificationField = farmPivot.getFilterHierarchy("Classification").getFields()[0];
  classificationField.applyFilter({
    manualFilter: { 
      selectedItems: ["Organic"] /* The included items. */
    }
  });

  const farmField = farmPivot.getHierarchy("Farm").getFields()[0];
  farmField.applyFilter({
    valueFilter: {
      condition: ExcelScript.ValueFilterCondition.greaterThan, /* The relationship of the value to the comparator. */
      comparator: 300, /* The value to which items are compared. */
      value: "Sum of Crates Sold Wholesale" /* The name of the data hierarchy. Note the "Sum of" prefix. */
      }
  });
```

:::image type="content" source="../images/pivottable-filters.png" alt-text="値フィルターと手動フィルターが適用された後のピボットテーブル。":::

### <a name="slicers"></a>スライサー

[スライサーは](https://support.microsoft.com/office/249f966b-a9d5-4b0f-b31a-12651785d29d) 、ピボットテーブル (または標準テーブル) のデータをフィルター処理します。 これらは、ワークシート内の移動可能なオブジェクトであり、簡単にフィルター処理を選択できます。 スライサーは、手動フィルターと同様の方法で動作します `PivotFilterHierarchy`。 アイテムをピボットテーブルに `PivotField` 含めるか、ピボットテーブルから除外するかを切り替えます。

次のコード スニペットでは、"Type" フィールドのスライサーを追加します。 選択した項目を "膻舩" と "ライム" に設定し、スライサーを 400 ピクセル左に移動します。

```typescript
  const fruitSlicer = pivotSheet.addSlicer(
    farmPivot, /* The table or PivotTale to be sliced. */
    farmPivot.getHierarchy("Type").getFields()[0] /* What source to use as the slicer options. */
  );
  fruitSlicer.selectItems(["Lemon", "Lime"]);
  fruitSlicer.setLeft(400);
```

:::image type="content" source="../images/slicer.png" alt-text="ピボットテーブル上のデータをフィルター処理するスライサー。":::

## <a name="see-also"></a>関連項目

- [Excel on the web での Office スクリプトのスクリプトの基本事項](scripting-fundamentals.md)
- [Office スクリプト API リファレンス](/javascript/api/office-scripts/overview)
