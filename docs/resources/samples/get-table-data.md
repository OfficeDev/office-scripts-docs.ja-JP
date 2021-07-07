---
title: JSON Excelデータを出力する
description: テーブル データを JSON Excelとして出力する方法について説明します。Power Automate。
ms.date: 06/29/2021
localization_priority: Normal
ms.openlocfilehash: 63379d1323f5e2084f4aa39af3f4b6e5e6d7e7bb
ms.sourcegitcommit: 211c157ca746e266eeb079f5fa1925a1e35ab702
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 07/07/2021
ms.locfileid: "53313947"
---
# <a name="output-excel-table-data-as-json-for-usage-in-power-automate"></a><span data-ttu-id="b3c88-103">テーブルExcelを JSON として出力して、テーブルの使用状況をPower Automate</span><span class="sxs-lookup"><span data-stu-id="b3c88-103">Output Excel table data as JSON for usage in Power Automate</span></span>

<span data-ttu-id="b3c88-104">Excelデータは、JSON 形式のオブジェクトの配列として表されます。</span><span class="sxs-lookup"><span data-stu-id="b3c88-104">Excel table data can be represented as an array of objects in the form of JSON.</span></span> <span data-ttu-id="b3c88-105">各オブジェクトは、テーブル内の行を表します。</span><span class="sxs-lookup"><span data-stu-id="b3c88-105">Each object represents a row in the table.</span></span> <span data-ttu-id="b3c88-106">これにより、ユーザーに表示されるExcel形式でデータを抽出できます。</span><span class="sxs-lookup"><span data-stu-id="b3c88-106">This helps extract the data from Excel in a consistent format that is visible to the user.</span></span> <span data-ttu-id="b3c88-107">その後、データを他のシステムに与え、Power Automateできます。</span><span class="sxs-lookup"><span data-stu-id="b3c88-107">The data can then be given to other systems through Power Automate flows.</span></span>

<span data-ttu-id="b3c88-108">_入力テーブル のデータ_</span><span class="sxs-lookup"><span data-stu-id="b3c88-108">_Input table data_</span></span>

:::image type="content" source="../../images/table-input.png" alt-text="入力テーブル データを示すワークシート。":::

<span data-ttu-id="b3c88-110">このサンプルのバリエーションには、表の列の 1 つにもハイパーリンクが含まれています。</span><span class="sxs-lookup"><span data-stu-id="b3c88-110">A variation of this sample also includes the hyperlinks in one of the table columns.</span></span> <span data-ttu-id="b3c88-111">これにより、追加レベルのセル データを JSON に表示できます。</span><span class="sxs-lookup"><span data-stu-id="b3c88-111">This allows additional levels of cell data to be surfaced in the JSON.</span></span>

<span data-ttu-id="b3c88-112">_ハイパーリンクを含む入力テーブル データ_</span><span class="sxs-lookup"><span data-stu-id="b3c88-112">_Input table data that includes hyperlinks_</span></span>

:::image type="content" source="../../images/table-hyperlink-view.png" alt-text="ハイパーリンクとして書式設定されたテーブル データの列を示すワークシート。":::

<span data-ttu-id="b3c88-114">_ハイパーリンクを編集するダイアログ_</span><span class="sxs-lookup"><span data-stu-id="b3c88-114">_Dialog to edit hyperlink_</span></span>

:::image type="content" source="../../images/table-hyperlink-edit.png" alt-text="ハイパーリンクを変更するためのオプションを表示する [ハイパーリンクの編集] ダイアログ ボックス。":::

## <a name="sample-excel-file"></a><span data-ttu-id="b3c88-116">サンプル Excel ファイル</span><span class="sxs-lookup"><span data-stu-id="b3c88-116">Sample Excel file</span></span>

<span data-ttu-id="b3c88-117">すぐに使用できる <a href="table-data-with-hyperlinks.xlsx"> ブックtable-data-with-hyperlinks.xlsx</a> ファイル をダウンロードします。</span><span class="sxs-lookup"><span data-stu-id="b3c88-117">Download the file <a href="table-data-with-hyperlinks.xlsx">table-data-with-hyperlinks.xlsx</a> for a ready-to-use workbook.</span></span> <span data-ttu-id="b3c88-118">次のスクリプトを追加して、サンプルを自分で試してみてください。</span><span class="sxs-lookup"><span data-stu-id="b3c88-118">Add the following script to try the sample yourself!</span></span>

## <a name="sample-code-return-table-data-as-json"></a><span data-ttu-id="b3c88-119">サンプル コード: JSON としてテーブル データを返す</span><span class="sxs-lookup"><span data-stu-id="b3c88-119">Sample code: Return table data as JSON</span></span>

> [!NOTE]
> <span data-ttu-id="b3c88-120">テーブル列に一致 `interface TableData` する構造を変更できます。</span><span class="sxs-lookup"><span data-stu-id="b3c88-120">You can change the `interface TableData` structure to match your table columns.</span></span> <span data-ttu-id="b3c88-121">スペースを含む列名の場合は、サンプルに含まれているなど、キーを二重引用符で囲んでください `"Event ID"` 。</span><span class="sxs-lookup"><span data-stu-id="b3c88-121">Note that for column names with spaces, be sure to place your key in quotation marks, such as with `"Event ID"` in the sample.</span></span>

```TypeScript
function main(workbook: ExcelScript.Workbook): TableData[] {
  // Get the first table in the "PlainTable" worksheet.
  // If you know the table name, use `workbook.getTable('TableName')` instead.
  const table = workbook.getWorksheet('PlainTable').getTables()[0];

  // Get all the values from the table as text.
  const texts = table.getRange().getTexts();

  // Create an array of JSON objects that match the row structure.
  let returnObjects: TableData[] = [];
  if (table.getRowCount() > 0)  {
    returnObjects = returnObjectFromValues(texts);
  }

  // Log the information and return it for a Power Automate flow.
  console.log(JSON.stringify(returnObjects));  
  return returnObjects
}

// This function converts a 2D-array of values into a generic JSON object.
// In this case, we have defined the TableData object, but any similar interface would work.
function returnObjectFromValues(values: string[][]): TableData[] {
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

  return objectArray as TableData[];
}

interface TableData {
  "Event ID": string
  Date: string
  Location: string
  Capacity: string
  Speakers: string
}
```

### <a name="sample-output-from-the-plaintable-worksheet"></a><span data-ttu-id="b3c88-122">"PlainTable" ワークシートからの出力例</span><span class="sxs-lookup"><span data-stu-id="b3c88-122">Sample output from the "PlainTable" worksheet</span></span>

```json
[{
    "Event ID": "E107",
    "Date": "2020-12-10",
    "Location": "Montgomery",
    "Capacity": "10",
    "Speakers": "Debra Berger"
}, {
    "Event ID": "E108",
    "Date": "2020-12-11",
    "Location": "Montgomery",
    "Capacity": "10",
    "Speakers": "Delia Dennis"
}, {
    "Event ID": "E109",
    "Date": "2020-12-12",
    "Location": "Montgomery",
    "Capacity": "10",
    "Speakers": "Diego Siciliani"
}, {
    "Event ID": "E110",
    "Date": "2020-12-13",
    "Location": "Boise",
    "Capacity": "25",
    "Speakers": "Gerhart Moller"
}, {
    "Event ID": "E111",
    "Date": "2020-12-14",
    "Location": "Salt Lake City",
    "Capacity": "20",
    "Speakers": "Grady Archie"
}, {
    "Event ID": "E112",
    "Date": "2020-12-15",
    "Location": "Fremont",
    "Capacity": "25",
    "Speakers": "Irvin Sayers"
}, {
    "Event ID": "E113",
    "Date": "2020-12-16",
    "Location": "Salt Lake City",
    "Capacity": "20",
    "Speakers": "Isaiah Langer"
}, {
    "Event ID": "E114",
    "Date": "2020-12-17",
    "Location": "Salt Lake City",
    "Capacity": "20",
    "Speakers": "Johanna Lorenz"
}]
```

## <a name="sample-code-return-table-data-as-json-with-hyperlink-text"></a><span data-ttu-id="b3c88-123">サンプル コード: ハイパーリンク テキストを含む JSON としてテーブル データを返す</span><span class="sxs-lookup"><span data-stu-id="b3c88-123">Sample code: Return table data as JSON with hyperlink text</span></span>

> [!NOTE]
> <span data-ttu-id="b3c88-124">スクリプトは常に、テーブルの 4 列目 (インデックス 0) からハイパーリンクを抽出します。</span><span class="sxs-lookup"><span data-stu-id="b3c88-124">The script always extracts hyperlinks from the 4th column (0 index) of the table.</span></span> <span data-ttu-id="b3c88-125">コメントの下のコードを変更することで、その順序を変更したり、複数の列をハイパーリンク データとして含めすることができます。 `// For the 4th column (0 index), extract the hyperlink and use that instead of text.`</span><span class="sxs-lookup"><span data-stu-id="b3c88-125">You can change that order or include multiple columns as hyperlink data by modifying the code under the comment `// For the 4th column (0 index), extract the hyperlink and use that instead of text.`</span></span>

```TypeScript
function main(workbook: ExcelScript.Workbook): TableData[] {
  // Get the first table in the "WithHyperLink" worksheet.
  // If you know the table name, use `workbook.getTable('TableName')` instead.
  const table = workbook.getWorksheet('WithHyperLink').getTables()[0];

  // Get all the values from the table as text.
  const range = table.getRange();

  // Create an array of JSON objects that match the row structure.
  let returnObjects: TableData[] = [];
  if (table.getRowCount() > 0)  {
    returnObjects = returnObjectFromValues(range);
  }

  // Log the information and return it for a Power Automate flow.
  console.log(JSON.stringify(returnObjects));  
  return returnObjects
}

function returnObjectFromValues(range: ExcelScript.Range): TableData[] {
  let values = range.getTexts();
  let objectArray = [];
  let objectKeys: string[] = [];
  for (let i = 0; i < values.length; i++) {
    if (i === 0) {
      objectKeys = values[i]
      continue;
    }

    let object = {}
    for (let j = 0; j < values[i].length; j++) {
      // For the 4th column (0 index), extract the hyperlink and use that instead of text. 
      if (j === 4) {
        object[objectKeys[j]] = range.getCell(i, j).getHyperlink().address;
      } else {
        object[objectKeys[j]] = values[i][j];
      }
    }

    objectArray.push(object);
  }
  return objectArray as TableData[];
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

### <a name="sample-output-from-the-withhyperlink-worksheet"></a><span data-ttu-id="b3c88-126">"WithHyperLink" ワークシートからの出力例</span><span class="sxs-lookup"><span data-stu-id="b3c88-126">Sample output from the "WithHyperLink" worksheet</span></span>

```json
[{
    "Event ID": "E107",
    "Date": "2020-12-10",
    "Location": "Montgomery",
    "Capacity": "10",
    "Search link": "https://www.google.com/search?q=Montgomery",
    "Speakers": "Debra Berger"
}, {
    "Event ID": "E108",
    "Date": "2020-12-11",
    "Location": "Montgomery",
    "Capacity": "10",
    "Search link": "https://www.google.com/search?q=Montgomery",
    "Speakers": "Delia Dennis"
}, {
    "Event ID": "E109",
    "Date": "2020-12-12",
    "Location": "Montgomery",
    "Capacity": "10",
    "Search link": "https://www.google.com/search?q=Montgomery",
    "Speakers": "Diego Siciliani"
}, {
    "Event ID": "E110",
    "Date": "2020-12-13",
    "Location": "Boise",
    "Capacity": "25",
    "Search link": "https://www.google.com/search?q=Boise",
    "Speakers": "Gerhart Moller"
}, {
    "Event ID": "E111",
    "Date": "2020-12-14",
    "Location": "Salt Lake City",
    "Capacity": "20",
    "Search link": "https://www.google.com/search?q=salt+lake+city",
    "Speakers": "Grady Archie"
}, {
    "Event ID": "E112",
    "Date": "2020-12-15",
    "Location": "Fremont",
    "Capacity": "25",
    "Search link": "https://www.google.com/search?q=Fremont",
    "Speakers": "Irvin Sayers"
}, {
    "Event ID": "E113",
    "Date": "2020-12-16",
    "Location": "Salt Lake City",
    "Capacity": "20",
    "Search link": "https://www.google.com/search?q=salt+lake+city",
    "Speakers": "Isaiah Langer"
}, {
    "Event ID": "E114",
    "Date": "2020-12-17",
    "Location": "Salt Lake City",
    "Capacity": "20",
    "Search link": "https://www.google.com/search?q=salt+lake+city",
    "Speakers": "Johanna Lorenz"
}]
```

## <a name="use-in-power-automate"></a><span data-ttu-id="b3c88-127">[Power Automate</span><span class="sxs-lookup"><span data-stu-id="b3c88-127">Use in Power Automate</span></span>

<span data-ttu-id="b3c88-128">このようなスクリプトを使用する方法については、「Power Automateを使用して自動化されたワークフローを作成する」[を参照Power Automate。](../../tutorials/excel-power-automate-returns.md#create-an-automated-workflow-with-power-automate)</span><span class="sxs-lookup"><span data-stu-id="b3c88-128">For how to use such a script in Power Automate, see [Create an automated workflow with Power Automate](../../tutorials/excel-power-automate-returns.md#create-an-automated-workflow-with-power-automate).</span></span>
