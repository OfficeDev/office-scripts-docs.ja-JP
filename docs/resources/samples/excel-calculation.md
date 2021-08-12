---
title: 計算モードを管理Excel
description: スクリプトを使用してOfficeモードを管理する方法について説明Excel on the web。
ms.date: 05/06/2021
localization_priority: Normal
ms.openlocfilehash: d33c4f21b21333ccefe26effc3df70235978b480a999364793e9a45d21dfba7f
ms.sourcegitcommit: 75f7ed8c2d23a104acc293f8ce29ea580b4fcdc5
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 08/11/2021
ms.locfileid: "57846709"
---
# <a name="manage-calculation-mode-in-excel"></a>計算モードを管理Excel

このサンプルでは、スクリプトを使用[して計算](/javascript/api/office-scripts/excelscript/excelscript.calculationmode)モードを使用し、計算モードでメソッドExcel on the web Officeします。 スクリプトは、任意のファイルでExcelできます。

## <a name="scenario"></a>シナリオ

数式の数が多いブックは、再計算に時間がかかる場合があります。 計算が発生Excel制御するのではなく、スクリプトの一部として管理できます。 これは、特定のシナリオでのパフォーマンスに役立ちます。

サンプル スクリプトは、計算モードを手動に設定します。 つまり、スクリプトが数式に指示した場合 (または UI を使用して手動で計算した場合) にのみ、ブックが数式 [を再計算します](https://support.microsoft.com/office/change-formula-recalculation-iteration-or-precision-in-excel-73fc7dac-91cf-4d36-86e8-67124f6bcce4)。 次に、スクリプトは現在の計算モードを表示し、ブック全体を完全に再計算します。

## <a name="sample-code-control-calculation-mode"></a>サンプル コード: コントロールの計算モード

```TypeScript
function main(workbook: ExcelScript.Workbook) {
    // Set the calculation mode to manual.
    workbook.getApplication().setCalculationMode(ExcelScript.CalculationMode.manual);
    // Get and log the calculation mode.
    const calcMode = workbook.getApplication().getCalculationMode();    
    console.log(calcMode);
    // Manually calculate the file.
    workbook.getApplication().calculate(ExcelScript.CalculationType.full);
}
```

## <a name="training-video-manage-calculation-mode"></a>トレーニング ビデオ: 計算モードの管理

[Sudhi Ramamurthy が YouTube でこのサンプルを歩くのを見る](https://youtu.be/iw6O8QH01CI).
