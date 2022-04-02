---
title: 計算モードを管理Excel
description: スクリプトを使用してOfficeモードを管理する方法について説明Excel on the web。
ms.date: 05/06/2021
ms.localizationpriority: medium
ms.openlocfilehash: fec88c904d95bfdab1514d44921f7fb1c6e9dd35
ms.sourcegitcommit: 7023b9e23499806901a5ecf8ebc460b76887cca6
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 03/31/2022
ms.locfileid: "64585514"
---
# <a name="manage-calculation-mode-in-excel"></a>計算モードを管理Excel

このサンプルでは、スクリプトを[使用して計算](/javascript/api/office-scripts/excelscript/excelscript.calculationmode)モードを使用し、Excel on the webメソッドOffice示します。 スクリプトは、任意のファイルでExcelできます。

## <a name="scenario"></a>シナリオ

数式の数が多いブックは、再計算に時間がかかる場合があります。 計算が行Excel制御するのではなく、スクリプトの一部として管理できます。 これは、特定のシナリオでのパフォーマンスに役立ちます。

サンプル スクリプトは、計算モードを手動に設定します。 つまり、スクリプトが数式に指示した場合 (または UI を使用して手動で計算した場合) にのみ、ブックが数式 [を再計算します](https://support.microsoft.com/office/73fc7dac-91cf-4d36-86e8-67124f6bcce4)。 次に、スクリプトは現在の計算モードを表示し、ブック全体を完全に再計算します。

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

[Sudhi Ramamurthy が YouTube でこのサンプルを見るのを見る](https://youtu.be/iw6O8QH01CI)。
