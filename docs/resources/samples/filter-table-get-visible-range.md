---
title: Excel テーブルをフィルター処理し、表示範囲を取得する
description: Excel テーブルをフィルター処理しOfficeオブジェクトの配列として表示範囲を取得するには、スクリプトを使用する方法について学習します。
ms.date: 03/16/2021
localization_priority: Normal
ms.openlocfilehash: c0a5842af4a62162225e3fc10203c261b91e010a
ms.sourcegitcommit: 5d24e77df70aa2c1c982275d53213c2a9323ff86
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 04/02/2021
ms.locfileid: "51571468"
---
# <a name="filter-excel-table-and-get-visible-range-as-a-json-object"></a><span data-ttu-id="3dd60-103">Excel テーブルをフィルター処理し、JSON オブジェクトとして表示範囲を取得する</span><span class="sxs-lookup"><span data-stu-id="3dd60-103">Filter Excel table and get visible range as a JSON object</span></span>

<span data-ttu-id="3dd60-104">このサンプルでは、Excel テーブルをフィルター処理し、表示範囲を JSON オブジェクトとして返します。</span><span class="sxs-lookup"><span data-stu-id="3dd60-104">This sample filters an Excel table and returns the visible range as a JSON object.</span></span> <span data-ttu-id="3dd60-105">この JSON は、大規模なソリューションの一部として Power Automate フローに提供できます。</span><span class="sxs-lookup"><span data-stu-id="3dd60-105">This JSON could be provided to a Power Automate flow as part of a larger solution.</span></span>

## <a name="example-scenario"></a><span data-ttu-id="3dd60-106">シナリオ例</span><span class="sxs-lookup"><span data-stu-id="3dd60-106">Example scenario</span></span>

* <span data-ttu-id="3dd60-107">テーブル列にフィルターを適用します。</span><span class="sxs-lookup"><span data-stu-id="3dd60-107">Apply a filter to a table column.</span></span>
* <span data-ttu-id="3dd60-108">フィルター処理後に表示範囲を抽出します。</span><span class="sxs-lookup"><span data-stu-id="3dd60-108">Extract the visible range after filtering.</span></span>
* <span data-ttu-id="3dd60-109">特定の JSON 構造を持つオブジェクトを [アセンブルして返します](#sample-json)。</span><span class="sxs-lookup"><span data-stu-id="3dd60-109">Assemble and return an object with a [specific JSON structure](#sample-json).</span></span>

## <a name="sample-code-filter-a-table-and-get-visible-range"></a><span data-ttu-id="3dd60-110">サンプル コード: テーブルをフィルター処理し、表示範囲を取得する</span><span class="sxs-lookup"><span data-stu-id="3dd60-110">Sample code: Filter a table and get visible range</span></span>

<span data-ttu-id="3dd60-111">次のスクリプトは、テーブルをフィルター処理し、表示範囲を取得します。</span><span class="sxs-lookup"><span data-stu-id="3dd60-111">The following script filters a table and gets the visible range.</span></span>

<span data-ttu-id="3dd60-112">サンプル ファイルをダウンロード <a href="table-filter.xlsx">table-filter.xlsx</a> このスクリプトで使用して、自分で試してみてください。</span><span class="sxs-lookup"><span data-stu-id="3dd60-112">Download the sample file <a href="table-filter.xlsx">table-filter.xlsx</a> and use it with this script to try it out yourself!</span></span>

```TypeScript
function main(workbook: ExcelScript.Workbook): ReturnTemplate {
  const table1 = workbook.getTable("Table1");
  const keyColumnValues: string [] = table1.getColumnByName('Station').getRangeBetweenHeaderAndTotal().getValues().map(v => v[0] as string);
  const uniqueKeys = keyColumnValues.filter((v, i, a) => a.indexOf(v) === i);

  console.log(uniqueKeys);
  const returnObj: ReturnTemplate = {}

  uniqueKeys.forEach((key: string) => {
    table1.getColumnByName('Station').getFilter()
      .applyValuesFilter([key]);
    const rangeView = table1.getRange().getVisibleView();
    returnObj[key] = returnObjectFromValues(rangeView.getValues() as string[][]);
  })
  table1.getColumnByName('Station').getFilter().clear();
  console.log(JSON.stringify(returnObj));
  return returnObj
}

function returnObjectFromValues(values: string[][]): BasicObj[] {
  let objArray = [];
  let objKeys: string[] = [];
  for (let i=0; i < values.length; i++) {
    if (i===0) {
      objKeys = values[i]
      continue;
    }
    let obj = {}
    for (let j=0; j < values[i].length; j++) {
      obj[objKeys[j]] = values[i][j]
    }
    objArray.push(obj);
  }
  return objArray;
}

interface BasicObj {
  [key: string] : string
}

interface ReturnTemplate {
  [key: string]: BasicObj[]
}
```

### <a name="sample-json"></a><span data-ttu-id="3dd60-113">サンプル JSON</span><span class="sxs-lookup"><span data-stu-id="3dd60-113">Sample JSON</span></span>

<span data-ttu-id="3dd60-114">各キーは、テーブルの一意の値を表します。</span><span class="sxs-lookup"><span data-stu-id="3dd60-114">Each key represents a unique value of a table.</span></span> <span data-ttu-id="3dd60-115">各配列インスタンスは、対応するフィルターを適用するときに表示される行を表します。</span><span class="sxs-lookup"><span data-stu-id="3dd60-115">Each array instance represents the row that is visible when the corresponding filter is applied.</span></span>

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

## <a name="training-video-filter-an-excel-table-and-get-the-visible-range"></a><span data-ttu-id="3dd60-116">トレーニング ビデオ: Excel テーブルをフィルター処理し、表示範囲を取得する</span><span class="sxs-lookup"><span data-stu-id="3dd60-116">Training video: Filter an Excel table and get the visible range</span></span>

<span data-ttu-id="3dd60-117">[![Excel テーブルをフィルター処理して表示範囲を取得する方法について、ステップバイステップのビデオを見る](../../images/visible-range-as-objects-vid.jpg)](https://youtu.be/Mv7BrvPq84A "Excel テーブルをフィルター処理し、表示範囲を取得する方法に関するステップバイステップのビデオ")</span><span class="sxs-lookup"><span data-stu-id="3dd60-117">[![Watch step-by-step video on how to filter an Excel table and get the visible range](../../images/visible-range-as-objects-vid.jpg)](https://youtu.be/Mv7BrvPq84A "Step-by-step video on how to filter an Excel table and get the visible range")</span></span>
