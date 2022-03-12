---
title: テーブルをExcelし、表示範囲を取得する
description: スクリプトを使用してOfficeテーブルをフィルター処理しExcelオブジェクトの配列として表示範囲を取得する方法について学習します。
ms.date: 03/10/2022
ms.localizationpriority: medium
ms.openlocfilehash: 752566aae1f5e64748e9a7a4c33447129905be22
ms.sourcegitcommit: 79ce4fad6d284b1aa71f5ad6d2938d9ad6a09fee
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 03/12/2022
ms.locfileid: "63459655"
---
# <a name="filter-excel-table-and-get-visible-range-as-a-json-object"></a>テーブルをExcelし、JSON オブジェクトとして表示範囲を取得する

次のサンプルでは、Excelをフィルター処理し、表示範囲を JSON オブジェクトとして返します。 この JSON は、大規模なソリューションの一部Power Automateフローに提供できます。

## <a name="example-scenario"></a>シナリオ例

* テーブル列にフィルターを適用します。
* フィルター処理後に表示範囲を抽出します。
* 特定の JSON 構造を持つオブジェクトを [アセンブルして返します](#sample-json)。

## <a name="sample-excel-file"></a>サンプル Excel ファイル

すぐに <a href="table-filter.xlsx">table-filter.xlsx</a> ブックのダウンロード を行います。 次のスクリプトを追加して、サンプルを自分で試してみてください。

## <a name="sample-code-filter-a-table-and-get-visible-range"></a>サンプル コード: テーブルをフィルター処理し、表示範囲を取得する

```TypeScript
function main(workbook: ExcelScript.Workbook): ReturnTemplate {
  // Get the "Station" column to use as key values in the filter.
  const table1 = workbook.getTable("Table1");
  const keyColumnValues: string [] = table1.getColumnByName('Station').getRangeBetweenHeaderAndTotal().getValues().map(value => value[0] as string);

  // Filter out repeated keys. This call to `filter` only returns the first instance of every unique element in the array.
  const uniqueKeys = keyColumnValues.filter((value, index, array) => array.indexOf(value) === index);
  console.log(uniqueKeys);

  const stationData: ReturnTemplate = {};

  // Filter the table to show only rows corresponding to each key.
  uniqueKeys.forEach((key: string) => {
    table1.getColumnByName('Station').getFilter()
      .applyValuesFilter([key]);
    
    // Get the visible view when a single filter is active.
    const rangeView = table1.getRange().getVisibleView();

    // Create a JSON object with every visible row.
    stationData[key] = returnObjectFromValues(rangeView.getValues() as string[][]);
  });

  // Remove the filters.
  table1.getColumnByName('Station').getFilter().clear();

  // Log the information and return it for a Power Automate flow.
  console.log(JSON.stringify(stationData));
  return stationData;
}

// This function converts a 2D-array of values into a generic JSON object.
function returnObjectFromValues(values: string[][]): BasicObject[] {
  let objectArray: BasicObject[] = [];
  let objectKeys: string[] = [];
  for (let i = 0; i < values.length; i++) {
    if (i === 0) {
      objectKeys = values[i]
      continue;
    }

    let object = {}
    for (let j = 0; j < values[i].length; j++) {
      object[objectKeys[j]] = values[i][j]
    }

    objectArray.push(object);
  }

  return objectArray;
}

interface BasicObject {
  [key: string] : string
}

interface ReturnTemplate {
  [key: string]: BasicObject[]
}
```

### <a name="sample-json"></a>サンプル JSON

各キーは、テーブルの一意の値を表します。 各配列インスタンスは、対応するフィルターを適用するときに表示される行を表します。

```json
{
  "Station-1": [{
    "Station": "Station-1",
    "Shift": "Morning",
    "Date": "27-Oct-20",
    "Responsible": "Debra Berger",
    "Reason&quot;: &quot;"
  }, {
    "Station": "Station-1",
    "Shift": "Afternoon",
    "Date": "27-Oct-20",
    "Responsible": "Delia Dennis",
    "Reason&quot;: &quot;"
  }, {
    "Station": "Station-1",
    "Shift": "Afternoon",
    "Date": "28-Oct-20",
    "Responsible": "Lidia Holloway",
    "Reason&quot;: &quot;"
  }],
  "Station-2": [{
    "Station": "Station-2",
    "Shift": "Morning",
    "Date": "27-Oct-20",
    "Responsible": "Gerhart Moller",
    "Reason&quot;: &quot;"
  }, {
    "Station": "Station-2",
    "Shift": "Afternoon",
    "Date": "28-Oct-20",
    "Responsible": "Grady Archie",
    "Reason&quot;: &quot;"
  }],
  "Station-3": [{
    "Station": "Station-3",
    "Shift": "Morning",
    "Date": "27-Oct-20",
    "Responsible": "Isaiah Langer",
    "Reason&quot;: &quot;"
  }]
}
```

## <a name="training-video-filter-an-excel-table-and-get-the-visible-range"></a>トレーニング ビデオ: テーブルのExcelし、表示範囲を取得する

[Sudhi Ramamurthy が YouTube でこのサンプルを見るのを見る](https://youtu.be/Mv7BrvPq84A)。
