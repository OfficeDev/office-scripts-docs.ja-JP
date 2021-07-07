---
title: テーブルをExcelし、表示範囲を取得する
description: スクリプトを使用してOfficeテーブルをフィルター処理しExcelオブジェクトの配列として表示範囲を取得する方法について学習します。
ms.date: 06/29/2021
localization_priority: Normal
ms.openlocfilehash: b19b826f95c7e7aeb331130fde05afaafe500c3d
ms.sourcegitcommit: 211c157ca746e266eeb079f5fa1925a1e35ab702
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 07/07/2021
ms.locfileid: "53313954"
---
# <a name="filter-excel-table-and-get-visible-range-as-a-json-object"></a><span data-ttu-id="0f537-103">テーブルExcelし、JSON オブジェクトとして表示範囲を取得する</span><span class="sxs-lookup"><span data-stu-id="0f537-103">Filter Excel table and get visible range as a JSON object</span></span>

<span data-ttu-id="0f537-104">次のサンプルでは、Excelをフィルター処理し、表示範囲を JSON オブジェクトとして返します。</span><span class="sxs-lookup"><span data-stu-id="0f537-104">This sample filters an Excel table and returns the visible range as a JSON object.</span></span> <span data-ttu-id="0f537-105">この JSON は、大規模なソリューションの一部Power Automateフローに提供できます。</span><span class="sxs-lookup"><span data-stu-id="0f537-105">This JSON could be provided to a Power Automate flow as part of a larger solution.</span></span>

## <a name="example-scenario"></a><span data-ttu-id="0f537-106">シナリオ例</span><span class="sxs-lookup"><span data-stu-id="0f537-106">Example scenario</span></span>

* <span data-ttu-id="0f537-107">テーブル列にフィルターを適用します。</span><span class="sxs-lookup"><span data-stu-id="0f537-107">Apply a filter to a table column.</span></span>
* <span data-ttu-id="0f537-108">フィルター処理後に表示範囲を抽出します。</span><span class="sxs-lookup"><span data-stu-id="0f537-108">Extract the visible range after filtering.</span></span>
* <span data-ttu-id="0f537-109">特定の JSON 構造を持つオブジェクトを [アセンブルして返します](#sample-json)。</span><span class="sxs-lookup"><span data-stu-id="0f537-109">Assemble and return an object with a [specific JSON structure](#sample-json).</span></span>

## <a name="sample-excel-file"></a><span data-ttu-id="0f537-110">サンプル Excel ファイル</span><span class="sxs-lookup"><span data-stu-id="0f537-110">Sample Excel file</span></span>

<span data-ttu-id="0f537-111">すぐに <a href="table-filter.xlsx"> 使用table-filter.xlsx</a> ブックのブックをダウンロードします。</span><span class="sxs-lookup"><span data-stu-id="0f537-111">Download <a href="table-filter.xlsx">table-filter.xlsx</a> for a ready-to-use workbook.</span></span> <span data-ttu-id="0f537-112">次のスクリプトを追加して、サンプルを自分で試してみてください。</span><span class="sxs-lookup"><span data-stu-id="0f537-112">Add the following script to try the sample yourself!</span></span>

## <a name="sample-code-filter-a-table-and-get-visible-range"></a><span data-ttu-id="0f537-113">サンプル コード: テーブルをフィルター処理し、表示範囲を取得する</span><span class="sxs-lookup"><span data-stu-id="0f537-113">Sample code: Filter a table and get visible range</span></span>

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
  let objectArray = [];
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

### <a name="sample-json"></a><span data-ttu-id="0f537-114">サンプル JSON</span><span class="sxs-lookup"><span data-stu-id="0f537-114">Sample JSON</span></span>

<span data-ttu-id="0f537-115">各キーは、テーブルの一意の値を表します。</span><span class="sxs-lookup"><span data-stu-id="0f537-115">Each key represents a unique value of a table.</span></span> <span data-ttu-id="0f537-116">各配列インスタンスは、対応するフィルターを適用するときに表示される行を表します。</span><span class="sxs-lookup"><span data-stu-id="0f537-116">Each array instance represents the row that is visible when the corresponding filter is applied.</span></span>

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

## <a name="training-video-filter-an-excel-table-and-get-the-visible-range"></a><span data-ttu-id="0f537-117">トレーニング ビデオ: テーブルのExcelし、表示範囲を取得する</span><span class="sxs-lookup"><span data-stu-id="0f537-117">Training video: Filter an Excel table and get the visible range</span></span>

<span data-ttu-id="0f537-118">[Sudhi Ramamurthy が YouTube でこのサンプルを歩くのを見る](https://youtu.be/Mv7BrvPq84A).</span><span class="sxs-lookup"><span data-stu-id="0f537-118">[Watch Sudhi Ramamurthy walk through this sample on YouTube](https://youtu.be/Mv7BrvPq84A).</span></span>
