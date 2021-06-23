---
title: シートの空白行を数える
description: Office スクリプトを使用して、ワークシート内のデータの代わりに空白行が含まれていますを検出し、空白の行数をレポートして、Power Automate フローで使用する方法について説明します。
ms.date: 05/04/2021
localization_priority: Normal
ms.openlocfilehash: 73fe0f995ee6ccaa1328b68983f0ec6887d96a09
ms.sourcegitcommit: 4693c8f79428ec74695328275703af0ba1bfea8f
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 06/23/2021
ms.locfileid: "53074579"
---
# <a name="count-blank-rows-on-sheets"></a><span data-ttu-id="5fa9a-103">シートの空白行を数える</span><span class="sxs-lookup"><span data-stu-id="5fa9a-103">Count blank rows on sheets</span></span>

<span data-ttu-id="5fa9a-104">このプロジェクトには、次の 2 つのスクリプトが含まれています。</span><span class="sxs-lookup"><span data-stu-id="5fa9a-104">This project includes two scripts:</span></span>

* <span data-ttu-id="5fa9a-105">[指定したシートの空白行を](#sample-code-count-blank-rows-on-a-given-sheet)数える: 指定したワークシートの使用範囲を走査し、空白の行数を返します。</span><span class="sxs-lookup"><span data-stu-id="5fa9a-105">[Count blank rows on a given sheet](#sample-code-count-blank-rows-on-a-given-sheet): Traverses the used range on a given worksheet and returns a blank row count.</span></span>
* <span data-ttu-id="5fa9a-106">[すべてのシートで空白行](#sample-code-count-blank-rows-on-all-sheets)を数える : すべてのワークシートの使用範囲を走査し、空白の行数を返します。</span><span class="sxs-lookup"><span data-stu-id="5fa9a-106">[Count blank rows on all sheets](#sample-code-count-blank-rows-on-all-sheets): Traverses the used range on _all of the worksheets_ and returns a blank row count.</span></span>

> [!NOTE]
> <span data-ttu-id="5fa9a-107">スクリプトの場合、空白の行はデータがない任意の行です。</span><span class="sxs-lookup"><span data-stu-id="5fa9a-107">For our script, a blank row is any row where there's no data.</span></span> <span data-ttu-id="5fa9a-108">行には書式設定を指定できます。</span><span class="sxs-lookup"><span data-stu-id="5fa9a-108">The row can have formatting.</span></span>

<span data-ttu-id="5fa9a-109">_このシートは、4 つの空白行の数を返します_</span><span class="sxs-lookup"><span data-stu-id="5fa9a-109">_This sheet returns count of 4 blank rows_</span></span>

:::image type="content" source="../../images/blank-rows.png" alt-text="空白行を含むデータを示すワークシート。":::

<span data-ttu-id="5fa9a-111">_このシートは、0 行の数を返します (すべての行にいくつかのデータがあります)_</span><span class="sxs-lookup"><span data-stu-id="5fa9a-111">_This sheet returns count of 0 blank rows (all rows have some data)_</span></span>

:::image type="content" source="../../images/no-blank-rows.png" alt-text="空白行のないデータを示すワークシート。":::

## <a name="sample-code-count-blank-rows-on-a-given-sheet"></a><span data-ttu-id="5fa9a-113">サンプル コード: 特定のシートの空白行を数える</span><span class="sxs-lookup"><span data-stu-id="5fa9a-113">Sample code: Count blank rows on a given sheet</span></span>

```TypeScript
function main(workbook: ExcelScript.Workbook): number
{
  // Get the worksheet named "Sheet1".
  const sheet = workbook.getWorksheet('Sheet1'); 
  
  // Get the entire data range.
  const range = sheet.getUsedRange(true);

  // If the used range is empty, end the script.
  if (!range) {
    console.log(`No data on this sheet.`);
    return;
  }
  
  // Log the address of the used range.
  console.log(`Used range for the worksheet: ${range.getAddress()}`);
    
  // Look through the values in the range for blank rows.
  const values = range.getValues();
  let emptyRows = 0;
  for (let row of values) {
    let emptyRow = true;
    
    // Look at every cell in the row for one with a value.
    for (let cell of row) {
      if (cell.toString().length > 0) {
        emptyRow = false
      }
    }

    // If no cell had a value, the row is empty.
    if (emptyRow) {
      emptyRows++;
    }
  }

  // Log the number of empty rows.
  console.log(`Total empty rows: ${emptyRows}`);

  // Return the number of empty rows for use in a Power Automate flow.
  return emptyRows;
}
```

## <a name="sample-code-count-blank-rows-on-all-sheets"></a><span data-ttu-id="5fa9a-114">サンプル コード: すべてのシートで空白行をカウントする</span><span class="sxs-lookup"><span data-stu-id="5fa9a-114">Sample code: Count blank rows on all sheets</span></span>

```TypeScript
function main(workbook: ExcelScript.Workbook): number
{
  // Loop through every worksheet in the workbook.
  const sheets = workbook.getWorksheets();
  let emptyRows = 0;
  for (let sheet of sheets) {     
    // Get the entire data range.
    const range = sheet.getUsedRange(true);
  
    // If the used range is empty, skip to the next worksheet.
    if (!range) {
      console.log(`No data on this sheet.`);
      continue;
    }
    
    // Log the address of the used range.
    console.log(`Used range for the worksheet: ${range.getAddress()}`);
      
    // Look through the values in the range for blank rows.
    const values = range.getValues();
    for (let row of values) {
      let emptyRow = true;
      
      // Look at every cell in the row for one with a value.
      for (let cell of row) {
        if (cell.toString().length > 0) {
          emptyRow = false
        }
      }
  
      // If no cell had a value, the row is empty.
      if (emptyRow) {
        emptyRows++;
      }
    }
  }

  // Log the number of empty rows.
  console.log(`Total empty rows: ${emptyRows}`);

  // Return the number of empty rows for use in a Power Automate flow.
  return emptyRows;
}
```

## <a name="use-with-power-automate"></a><span data-ttu-id="5fa9a-115">[ユーザーと一緒にPower Automate</span><span class="sxs-lookup"><span data-stu-id="5fa9a-115">Use with Power Automate</span></span>

:::image type="content" source="../../images/use-in-power-automate.png" alt-text="スクリプトPower Automateを実行するためにセットアップする方法を示すOfficeフロー。":::
