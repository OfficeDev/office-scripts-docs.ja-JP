---
title: シートの空白行を数える
description: Office スクリプトを使用して、ワークシート内のデータの代わりに空白行が含まれていますを検出し、空の行数をレポートして、Power Automate フローで使用する方法について説明します。
ms.date: 06/29/2021
ms.localizationpriority: medium
ms.openlocfilehash: 1ae513928b885994dc7f6d1b8ad66d694b61e7b7
ms.sourcegitcommit: 7023b9e23499806901a5ecf8ebc460b76887cca6
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 03/31/2022
ms.locfileid: "64585563"
---
# <a name="count-blank-rows-on-sheets"></a>シートの空白行を数える

このプロジェクトには、次の 2 つのスクリプトが含まれています。

* [指定したシートの空白行](#sample-code-count-blank-rows-on-a-given-sheet)を数える: 指定したワークシートの使用範囲を走査し、空白の行数を返します。
* [すべてのシートで空白行を](#sample-code-count-blank-rows-on-all-sheets)数える: すべてのワークシートの使用範囲を走査し、空白の行数を返します。

> [!NOTE]
> スクリプトの場合、空白の行はデータがない任意の行です。 行には書式設定を指定できます。

_このシートは、4 つの空白行の数を返します_

:::image type="content" source="../../images/blank-rows.png" alt-text="空白行を含むデータを示すワークシート。":::

_このシートは、0 行の数を返します (すべての行にいくつかのデータがあります)_

:::image type="content" source="../../images/no-blank-rows.png" alt-text="空白行のないデータを示すワークシート。":::

## <a name="sample-code-count-blank-rows-on-a-given-sheet"></a>サンプル コード: 特定のシートの空白行を数える

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

## <a name="sample-code-count-blank-rows-on-all-sheets"></a>サンプル コード: すべてのシートで空白行をカウントする

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
