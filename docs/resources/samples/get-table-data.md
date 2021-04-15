---
title: Excel データを JSON として出力する
description: Power Automate で使用する EXCEL テーブル データを JSON として出力する方法について説明します。
ms.date: 03/18/2021
localization_priority: Normal
ms.openlocfilehash: db6eb8f8645079eebc369e0a0622539075853953
ms.sourcegitcommit: 45ffe3dbd2c834b78592ad35928cf8096f5e80bc
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 04/14/2021
ms.locfileid: "51754797"
---
# <a name="output-excel-table-data-as-json-for-usage-in-power-automate"></a><span data-ttu-id="c34f3-103">Power Automate で使用するための JSON として Excel テーブル データを出力する</span><span class="sxs-lookup"><span data-stu-id="c34f3-103">Output Excel table data as JSON for usage in Power Automate</span></span>

<span data-ttu-id="c34f3-104">Excel テーブル データは、JSON 形式のオブジェクトの配列として表されます。</span><span class="sxs-lookup"><span data-stu-id="c34f3-104">Excel table data can be represented as an array of objects in the form of JSON.</span></span> <span data-ttu-id="c34f3-105">各オブジェクトは、テーブル内の行を表します。</span><span class="sxs-lookup"><span data-stu-id="c34f3-105">Each object represents a row in the table.</span></span> <span data-ttu-id="c34f3-106">これにより、ユーザーに表示される一貫性のある形式で Excel からデータを抽出できます。</span><span class="sxs-lookup"><span data-stu-id="c34f3-106">This helps extract the data from Excel in a consistent format that is visible to the user.</span></span> <span data-ttu-id="c34f3-107">その後、Power Automate フローを使用して他のシステムにデータを与えできます。</span><span class="sxs-lookup"><span data-stu-id="c34f3-107">The data can then be given to other systems through Power Automate flows.</span></span>

<span data-ttu-id="c34f3-108">_入力テーブル のデータ_</span><span class="sxs-lookup"><span data-stu-id="c34f3-108">_Input table data_</span></span>

:::image type="content" source="../../images/table-input.png" alt-text="入力テーブル データを示すワークシート。":::

<span data-ttu-id="c34f3-110">このサンプルのバリエーションには、表の列の 1 つにもハイパーリンクが含まれています。</span><span class="sxs-lookup"><span data-stu-id="c34f3-110">A variation of this sample also includes the hyperlinks in one of the table columns.</span></span> <span data-ttu-id="c34f3-111">これにより、追加レベルのセル データを JSON に表示できます。</span><span class="sxs-lookup"><span data-stu-id="c34f3-111">This allows additional levels of cell data to be surfaced in the JSON.</span></span>

<span data-ttu-id="c34f3-112">_ハイパーリンクを含む入力テーブル データ_</span><span class="sxs-lookup"><span data-stu-id="c34f3-112">_Input table data that includes hyperlinks_</span></span>

:::image type="content" source="../../images/table-hyperlink-view.png" alt-text="ハイパーリンクとして書式設定されたテーブル データの列を示すワークシート。":::

<span data-ttu-id="c34f3-114">_ハイパーリンクを編集するダイアログ_</span><span class="sxs-lookup"><span data-stu-id="c34f3-114">_Dialog to edit hyperlink_</span></span>

:::image type="content" source="../../images/table-hyperlink-edit.png" alt-text="ハイパーリンクを変更するためのオプションを表示する [ハイパーリンクの編集] ダイアログ ボックス。":::

## <a name="sample-excel-file"></a><span data-ttu-id="c34f3-116">サンプル Excel ファイル</span><span class="sxs-lookup"><span data-stu-id="c34f3-116">Sample Excel file</span></span>

<span data-ttu-id="c34f3-117">これらのサンプルで <a href="table-data-with-hyperlinks.xlsx">table-data-with-hyperlinks.xlsx</a> ファイルをダウンロードして、自分で試してみてください。</span><span class="sxs-lookup"><span data-stu-id="c34f3-117">Download the file <a href="table-data-with-hyperlinks.xlsx">table-data-with-hyperlinks.xlsx</a> used in these samples and try it out yourself!</span></span>

## <a name="sample-code-return-table-data-as-json"></a><span data-ttu-id="c34f3-118">サンプル コード: JSON としてテーブル データを返す</span><span class="sxs-lookup"><span data-stu-id="c34f3-118">Sample code: Return table data as JSON</span></span>

> [!NOTE]
> <span data-ttu-id="c34f3-119">テーブル列に一致 `interface TableData` する構造を変更できます。</span><span class="sxs-lookup"><span data-stu-id="c34f3-119">You can change the `interface TableData` structure to match your table columns.</span></span> <span data-ttu-id="c34f3-120">スペースを含む列名の場合は、サンプルに含まれているなど、キーを二重引用符で囲んでください `"Event ID"` 。</span><span class="sxs-lookup"><span data-stu-id="c34f3-120">Note that for column names with spaces, be sure to place your key in quotation marks, such as with `"Event ID"` in the sample.</span></span>

```TypeScript
function main(workbook: ExcelScript.Workbook): TableData[] {
  const table = workbook.getWorksheet('PlainTable').getTables()[0];
  // If you know the table name, you can also do the following:
  // const table = workbook.getTable('Table13436');
  const texts = table.getRange().getTexts();
  let returnObjects: TableData[] = [];
  if (table.getRowCount() > 0)  {
    returnObjects = returnObjectFromValues(texts);
  } 
  console.log(JSON.stringify(returnObjects));  
  return returnObjects
}

function returnObjectFromValues(values: string[][]): TableData[] {
  let objArray = [];
  let objKeys: string[] = [];
  for (let i = 0; i < values.length; i++) {
    if (i === 0) {
      objKeys = values[i]
      continue;
    }
    let obj = {}
    for (let j = 0; j < values[i].length; j++) {
      obj[objKeys[j]] = values[i][j]
    }
    objArray.push(obj);
  }
  return objArray as TableData[];
}

interface BasicObj {
  [key: string]: string
}

interface TableData {
  "Event ID": string
  Date: string
  Location: string
  Capacity: string
  Speakers: string
}
```

### <a name="sample-output"></a><span data-ttu-id="c34f3-121">サンプル出力</span><span class="sxs-lookup"><span data-stu-id="c34f3-121">Sample output</span></span>

```json
[{
    "Event ID": "E107",
    "Date": "2020-12-10",
    "Location": "Montgomery",
    "Capacity": "10",
    "Speakers&quot;: &quot;Debra Berger"
}, {
    "Event ID": "E108",
    "Date": "2020-12-11",
    "Location": "Montgomery",
    "Capacity": "10",
    "Speakers&quot;: &quot;Delia Dennis"
}, {
    "Event ID": "E109",
    "Date": "2020-12-12",
    "Location": "Montgomery",
    "Capacity": "10",
    "Speakers&quot;: &quot;Diego Siciliani"
}, {
    "Event ID": "E110",
    "Date": "2020-12-13",
    "Location": "Boise",
    "Capacity": "25",
    "Speakers&quot;: &quot;Gerhart Moller"
}, {
    "Event ID": "E111",
    "Date": "2020-12-14",
    "Location": "Salt Lake City",
    "Capacity": "20",
    "Speakers&quot;: &quot;Grady Archie"
}, {
    "Event ID": "E112",
    "Date": "2020-12-15",
    "Location": "Fremont",
    "Capacity": "25",
    "Speakers&quot;: &quot;Irvin Sayers"
}, {
    "Event ID": "E113",
    "Date": "2020-12-16",
    "Location": "Salt Lake City",
    "Capacity": "20",
    "Speakers&quot;: &quot;Isaiah Langer"
}, {
    "Event ID": "E114",
    "Date": "2020-12-17",
    "Location": "Salt Lake City",
    "Capacity": "20",
    "Speakers&quot;: &quot;Johanna Lorenz"
}]
```

## <a name="sample-code-return-table-data-as-json-with-hyperlink-text"></a><span data-ttu-id="c34f3-122">サンプル コード: ハイパーリンク テキストを含む JSON としてテーブル データを返す</span><span class="sxs-lookup"><span data-stu-id="c34f3-122">Sample code: Return table data as JSON with hyperlink text</span></span>

> [!NOTE]
> <span data-ttu-id="c34f3-123">スクリプトは常に、テーブルの 4 列目 (インデックス 0) からハイパーリンクを抽出します。</span><span class="sxs-lookup"><span data-stu-id="c34f3-123">The script always extracts hyperlinks from the 4th column (0 index) of the table.</span></span> <span data-ttu-id="c34f3-124">コメントの下のコードを変更することで、その順序を変更したり、複数の列をハイパーリンク データとして含めすることができます。 `// For the 4th column (0 index), extract the hyperlink and use that instead of text.`</span><span class="sxs-lookup"><span data-stu-id="c34f3-124">You can change that order or include multiple columns as hyperlink data by modifying the code under the comment `// For the 4th column (0 index), extract the hyperlink and use that instead of text.`</span></span>

```TypeScript
function main(workbook: ExcelScript.Workbook): TableData[] {
  const table = workbook.getWorksheet('WithHyperLink').getTables()[0];
  const range = table.getRange();
  // If you know the table name, you can also do the following:
  // const table = workbook.getTable('Table13436');
  const texts = table.getRange().getTexts();
  let returnObjects: TableData[] = [];
  if (table.getRowCount() > 0)  {
    returnObjects = returnObjectFromValues(texts, range);
  } 
  console.log(JSON.stringify(returnObjects));  
  return returnObjects
}

function returnObjectFromValues(values: string[][], range: ExcelScript.Range): TableData[] {
  let objArray = [];
  let objKeys: string[] = [];
  for (let i = 0; i < values.length; i++) {
    if (i === 0) {
      objKeys = values[i]
      continue;
    }
    let obj = {}
    for (let j = 0; j < values[i].length; j++) {
      // For the 4th column (0 index), extract the hyperlink and use that instead of text. 
      if (j === 4) {
        obj[objKeys[j]] = range.getCell(i, j).getHyperlink().address;
      } else {
        obj[objKeys[j]] = values[i][j];
      }
    }
    objArray.push(obj);
  }
  return objArray as TableData[];
}

interface BasicObj {
  [key: string]: string
}

interface TableData {
  "Event ID": string
  Date: string
  Location: string
  Capacity: string
  "Search link": string
  Speakers: string
}
```

### <a name="sample-output"></a><span data-ttu-id="c34f3-125">サンプル出力</span><span class="sxs-lookup"><span data-stu-id="c34f3-125">Sample output</span></span>

```json
[{
    "Event ID": "E107",
    "Date": "2020-12-10",
    "Location": "Montgomery",
    "Capacity": "10",
    "Search link": "https://www.google.com/search?q=Montgomery",
    "Speakers&quot;: &quot;Debra Berger"
}, {
    "Event ID": "E108",
    "Date": "2020-12-11",
    "Location": "Montgomery",
    "Capacity": "10",
    "Search link": "https://www.google.com/search?q=Montgomery",
    "Speakers&quot;: &quot;Delia Dennis"
}, {
    "Event ID": "E109",
    "Date": "2020-12-12",
    "Location": "Montgomery",
    "Capacity": "10",
    "Search link": "https://www.google.com/search?q=Montgomery",
    "Speakers&quot;: &quot;Diego Siciliani"
}, {
    "Event ID": "E110",
    "Date": "2020-12-13",
    "Location": "Boise",
    "Capacity": "25",
    "Search link": "https://www.google.com/search?q=Boise",
    "Speakers&quot;: &quot;Gerhart Moller"
}, {
    "Event ID": "E111",
    "Date": "2020-12-14",
    "Location": "Salt Lake City",
    "Capacity": "20",
    "Search link": "https://www.google.com/search?q=salt+lake+city",
    "Speakers&quot;: &quot;Grady Archie"
}, {
    "Event ID": "E112",
    "Date": "2020-12-15",
    "Location": "Fremont",
    "Capacity": "25",
    "Search link": "https://www.google.com/search?q=Fremont",
    "Speakers&quot;: &quot;Irvin Sayers"
}, {
    "Event ID": "E113",
    "Date": "2020-12-16",
    "Location": "Salt Lake City",
    "Capacity": "20",
    "Search link": "https://www.google.com/search?q=salt+lake+city",
    "Speakers&quot;: &quot;Isaiah Langer"
}, {
    "Event ID": "E114",
    "Date": "2020-12-17",
    "Location": "Salt Lake City",
    "Capacity": "20",
    "Search link": "https://www.google.com/search?q=salt+lake+city",
    "Speakers&quot;: &quot;Johanna Lorenz"
}]
```

## <a name="use-in-power-automate"></a><span data-ttu-id="c34f3-126">Power Automate での使用</span><span class="sxs-lookup"><span data-stu-id="c34f3-126">Use in Power Automate</span></span>

<span data-ttu-id="c34f3-127">Power Automate でこのようなスクリプトを使用する方法については、「Power Automate を使用して自動化されたワークフローを作成する [」を参照してください](../../tutorials/excel-power-automate-returns.md#create-an-automated-workflow-with-power-automate)。</span><span class="sxs-lookup"><span data-stu-id="c34f3-127">For how to use such a script in Power Automate, see [Create an automated workflow with Power Automate](../../tutorials/excel-power-automate-returns.md#create-an-automated-workflow-with-power-automate).</span></span>
