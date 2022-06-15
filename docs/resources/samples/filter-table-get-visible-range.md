---
title: テーブルExcelフィルター処理し、表示範囲を取得する
description: Office スクリプトを使用してExcelテーブルをフィルター処理し、オブジェクトの配列として表示範囲を取得する方法について説明します。
ms.date: 03/10/2022
ms.localizationpriority: medium
ms.openlocfilehash: 103ec97111720ab872c0be843aa0573781d98c44
ms.sourcegitcommit: dd01979d34b3499360d2f79a56f8a8f24f480eed
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 06/15/2022
ms.locfileid: "66088086"
---
# <a name="filter-excel-table-and-get-visible-range-as-a-json-object"></a>テーブルExcelフィルター処理し、JSON オブジェクトとして表示範囲を取得する

このサンプルでは、Excel テーブルをフィルター処理し、表示範囲を [JSON](https://www.w3schools.com/whatis/whatis_json.asp) オブジェクトとして返します。 この JSON は、大規模なソリューションの一部としてPower Automate フローに提供できます。

## <a name="example-scenario"></a>シナリオ例

* テーブル列にフィルターを適用します。
* フィルター処理後に表示範囲を抽出します。
* [特定の JSON 構造](#sample-json)を持つオブジェクトをアセンブルして返します。

## <a name="sample-excel-file"></a>サンプル Excel ファイル

すぐに使用できるブックの <a href="table-filter.xlsx">table-filter.xlsx</a> をダウンロードします。 サンプルを自分で試すには、次のスクリプトを追加します。

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

### <a name="sample-json"></a>JSON のサンプル

各キーは、テーブルの一意の値を表します。 各配列インスタンスは、対応するフィルターが適用されたときに表示される行を表します。 JSON の操作の詳細については、「[JSON を使用して、Office スクリプトとの間でデータを渡す](../../develop/use-json.md)」を参照してください。

```json
{
  "Station-1": [{
    "Station": "Station-1",
    "Shift": "Morning",
    "Date": "27-Oct-20",
    "Responsible": "Debra Berger",
    "Reason": ""
  }, {
    "Station": "Station-1",
    "Shift": "Afternoon",
    "Date": "27-Oct-20",
    "Responsible": "Delia Dennis",
    "Reason": ""
  }, {
    "Station": "Station-1",
    "Shift": "Afternoon",
    "Date": "28-Oct-20",
    "Responsible": "Lidia Holloway",
    "Reason": ""
  }],
  "Station-2": [{
    "Station": "Station-2",
    "Shift": "Morning",
    "Date": "27-Oct-20",
    "Responsible": "Gerhart Moller",
    "Reason": ""
  }, {
    "Station": "Station-2",
    "Shift": "Afternoon",
    "Date": "28-Oct-20",
    "Responsible": "Grady Archie",
    "Reason": ""
  }],
  "Station-3": [{
    "Station": "Station-3",
    "Shift": "Morning",
    "Date": "27-Oct-20",
    "Responsible": "Isaiah Langer",
    "Reason": ""
  }]
}
```

## <a name="training-video-filter-an-excel-table-and-get-the-visible-range"></a>トレーニング ビデオ: Excel テーブルをフィルター処理し、表示範囲を取得する

[YouTube でこのサンプルを見る、スディ Ramamurthy のチュートリアルをご覧ください](https://youtu.be/Mv7BrvPq84A)。
