---
title: シートの空白行を数える
description: Office スクリプトを使用して、ワークシート内のデータの代わりに空白行が含まれていますを検出し、空白の行数をレポートして、Power Automate フローで使用する方法について説明します。
ms.date: 03/31/2021
localization_priority: Normal
ms.openlocfilehash: db84f2446c168f867c325a05129fe982c9645731
ms.sourcegitcommit: f7a7aebfb687f2a35dbed07ed62ff352a114525a
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 05/06/2021
ms.locfileid: "52232586"
---
# <a name="count-blank-rows-on-sheets"></a><span data-ttu-id="71a0d-103">シートの空白行を数える</span><span class="sxs-lookup"><span data-stu-id="71a0d-103">Count blank rows on sheets</span></span>

<span data-ttu-id="71a0d-104">このプロジェクトには、次の 2 つのスクリプトが含まれています。</span><span class="sxs-lookup"><span data-stu-id="71a0d-104">This project includes two scripts:</span></span>

* <span data-ttu-id="71a0d-105">[指定したシートの空白行を](#sample-code-count-blank-rows-on-a-given-sheet)数える: 指定したワークシートの使用範囲を走査し、空白の行数を返します。</span><span class="sxs-lookup"><span data-stu-id="71a0d-105">[Count blank rows on a given sheet](#sample-code-count-blank-rows-on-a-given-sheet): Traverses the used range on a given worksheet and returns a blank row count.</span></span>
* <span data-ttu-id="71a0d-106">[すべてのシートで空白行](#sample-code-count-blank-rows-on-all-sheets)を数える : すべてのワークシートの使用範囲を走査し、空白の行数を返します。</span><span class="sxs-lookup"><span data-stu-id="71a0d-106">[Count blank rows on all sheets](#sample-code-count-blank-rows-on-all-sheets): Traverses the used range on _all of the worksheets_ and returns a blank row count.</span></span>

> [!NOTE]
> <span data-ttu-id="71a0d-107">スクリプトの場合、空白の行はデータがない任意の行です。</span><span class="sxs-lookup"><span data-stu-id="71a0d-107">For our script, a blank row is any row where there's no data.</span></span> <span data-ttu-id="71a0d-108">行には書式設定を指定できます。</span><span class="sxs-lookup"><span data-stu-id="71a0d-108">The row can have formatting.</span></span>

<span data-ttu-id="71a0d-109">_このシートは、4 つの空白行の数を返します_</span><span class="sxs-lookup"><span data-stu-id="71a0d-109">_This sheet returns count of 4 blank rows_</span></span>

:::image type="content" source="../../images/blank-rows.png" alt-text="空白行を含むデータを示すワークシート":::

<span data-ttu-id="71a0d-111">_このシートは、0 行の数を返します (すべての行にいくつかのデータがあります)_</span><span class="sxs-lookup"><span data-stu-id="71a0d-111">_This sheet returns count of 0 blank rows (all rows have some data)_</span></span>

:::image type="content" source="../../images/no-blank-rows.png" alt-text="空白行のないデータを示すワークシート":::

## <a name="sample-code-count-blank-rows-on-a-given-sheet"></a><span data-ttu-id="71a0d-113">サンプル コード: 特定のシートの空白行を数える</span><span class="sxs-lookup"><span data-stu-id="71a0d-113">Sample code: Count blank rows on a given sheet</span></span>

```TypeScript
function main(workbook: ExcelScript.Workbook): number
{
  const sheet = workbook.getWorksheet('Sheet1'); 
  // Getting the active worksheet is not suitable for a script used by Power Automate.
  // const sheet = workbook.getActiveWorksheet();
  
  const range = sheet.getUsedRange(true); // Get value only.
  if (!range) {
    console.log(`No data on this sheet. `);
    return;
  }
  console.log(`Used range for the worksheet: ${range.getAddress()}`);
  const values = range.getValues();
  let emptyRows = 0;
  for (let row of values) {
    let len = 0; 
    for (let cell of row) {
      len = len + cell.toString().length;
    }
    if (len === 0) { 
      emptyRows++;
    }
  }
  console.log(`Total empty row: ` + emptyRows);
  return emptyRows;
}
```

## <a name="sample-code-count-blank-rows-on-all-sheets"></a><span data-ttu-id="71a0d-114">サンプル コード: すべてのシートで空白行をカウントする</span><span class="sxs-lookup"><span data-stu-id="71a0d-114">Sample code: Count blank rows on all sheets</span></span>

```TypeScript
function main(workbook: ExcelScript.Workbook): number
{
  const sheets = workbook.getWorksheets();
  let emptyRows = 0;
  for (let sheet of sheets) { 
    const range = sheet.getUsedRange(true); // Get value only.
    if (!range) {
      console.log(`No data on this sheet. `);
      continue;
    }
    console.log(`Used range for the worksheet ${sheet.getName()}: ${range.getAddress()}`);
    const values = range.getValues();

    for (let row of values) {
      let len = 0;
      for (let cell of row) {
        len = len + cell.toString().length;
      }
      if (len === 0) {
        emptyRows++;
      }
    }
  }
  console.log(`Total empty row: ` + emptyRows);
  return emptyRows;
}
```

## <a name="use-with-power-automate"></a><span data-ttu-id="71a0d-115">[ユーザーと一緒にPower Automate</span><span class="sxs-lookup"><span data-stu-id="71a0d-115">Use with Power Automate</span></span>

:::image type="content" source="../../images/use-in-power-automate.png" alt-text="スクリプトPower Automate実行をセットアップする方法を示すOfficeフロー":::
