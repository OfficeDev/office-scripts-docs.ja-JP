---
title: 計算モードを管理Excel
description: スクリプトを使用してOfficeモードを管理する方法について説明Excel on the web。
ms.date: 04/28/2021
localization_priority: Normal
ms.openlocfilehash: 34a14874197ffda8487df5e450e3dcab980f7ed5
ms.sourcegitcommit: f7a7aebfb687f2a35dbed07ed62ff352a114525a
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 05/06/2021
ms.locfileid: "52232453"
---
# <a name="manage-calculation-mode-in-excel"></a>計算モードを管理Excel

このサンプルでは、スクリプトを使用[して計算](/javascript/api/office-scripts/excelscript/excelscript.calculationmode)モードを使用し、計算モードでメソッドExcel on the web Officeします。 スクリプトは、任意のファイルでExcelできます。

## <a name="scenario"></a>シナリオ

このExcel on the web、API を使用してファイルの計算モードをプログラムで制御できます。 次のアクションは、スクリプトを使用Officeできます。

1. 計算モードを取得します。
1. 計算モードを設定します。
1. 手動Excel (再計算とも呼ばれます) に設定されているファイルの数式を計算します。

## <a name="sample-code-control-calculation-mode"></a>サンプル コード: コントロールの計算モード

```TypeScript
function main(workbook: ExcelScript.Workbook) {
    // Set calculation mode.
    workbook.getApplication().setCalculationMode(ExcelScript.CalculationMode.manual);
    // Get calculation mode.
    const calcMode = workbook.getApplication().getCalculationMode();    
    console.log(calcMode);
    // Calculate (for manual mode files).
    workbook.getApplication().calculate(ExcelScript.CalculationType.full);
}
```

## <a name="training-video-manage-calculation-mode"></a>トレーニング ビデオ: 計算モードの管理

[Sudhi Ramamurthy が YouTube でこのサンプルを歩くのを見る](https://youtu.be/iw6O8QH01CI).
