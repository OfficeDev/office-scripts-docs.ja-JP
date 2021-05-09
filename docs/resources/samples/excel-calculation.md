---
title: 計算モードを管理Excel
description: スクリプトを使用してOfficeモードを管理する方法について説明Excel on the web。
ms.date: 05/06/2021
localization_priority: Normal
ms.openlocfilehash: a60fddc91b3a8f124a44722d0d75e6e9f239351d
ms.sourcegitcommit: 763d341857bcb209b2f2c278a82fdb63d0e18f0a
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 05/08/2021
ms.locfileid: "52285914"
---
# <a name="manage-calculation-mode-in-excel"></a><span data-ttu-id="f4544-103">計算モードを管理Excel</span><span class="sxs-lookup"><span data-stu-id="f4544-103">Manage calculation mode in Excel</span></span>

<span data-ttu-id="f4544-104">このサンプルでは、スクリプトを使用[して計算](/javascript/api/office-scripts/excelscript/excelscript.calculationmode)モードを使用し、計算モードでメソッドExcel on the web Officeします。</span><span class="sxs-lookup"><span data-stu-id="f4544-104">This sample shows how to use the [calculation mode](/javascript/api/office-scripts/excelscript/excelscript.calculationmode) and calculate methods in Excel on the web using Office Scripts.</span></span> <span data-ttu-id="f4544-105">スクリプトは、任意のファイルでExcelできます。</span><span class="sxs-lookup"><span data-stu-id="f4544-105">You can try the script on any Excel file.</span></span>

## <a name="scenario"></a><span data-ttu-id="f4544-106">シナリオ</span><span class="sxs-lookup"><span data-stu-id="f4544-106">Scenario</span></span>

<span data-ttu-id="f4544-107">数式の数が多いブックは、再計算に時間がかかる場合があります。</span><span class="sxs-lookup"><span data-stu-id="f4544-107">Workbooks with large numbers of formulas can take a while to recalculate.</span></span> <span data-ttu-id="f4544-108">計算が発生Excel制御するのではなく、スクリプトの一部として管理できます。</span><span class="sxs-lookup"><span data-stu-id="f4544-108">Rather than letting Excel control when calculations happen, you can manage them as part of your script.</span></span> <span data-ttu-id="f4544-109">これは、特定のシナリオでのパフォーマンスに役立ちます。</span><span class="sxs-lookup"><span data-stu-id="f4544-109">This will help with performance in certain scenarios.</span></span>

<span data-ttu-id="f4544-110">サンプル スクリプトは、計算モードを手動に設定します。</span><span class="sxs-lookup"><span data-stu-id="f4544-110">The sample script sets the calculation mode to manual.</span></span> <span data-ttu-id="f4544-111">つまり、スクリプトが数式に指示した場合 (または UI を使用して手動で計算した場合) にのみ、ブックが数式 [を再計算します](https://support.microsoft.com/office/change-formula-recalculation-iteration-or-precision-in-excel-73fc7dac-91cf-4d36-86e8-67124f6bcce4)。</span><span class="sxs-lookup"><span data-stu-id="f4544-111">This means that the workbook will only recalculate formulas when the script tells it to (or you [manually calculate through the UI](https://support.microsoft.com/office/change-formula-recalculation-iteration-or-precision-in-excel-73fc7dac-91cf-4d36-86e8-67124f6bcce4)).</span></span> <span data-ttu-id="f4544-112">次に、スクリプトは現在の計算モードを表示し、ブック全体を完全に再計算します。</span><span class="sxs-lookup"><span data-stu-id="f4544-112">The script then displays the current calculation mode and fully recalculates the entire workbook.</span></span>

## <a name="sample-code-control-calculation-mode"></a><span data-ttu-id="f4544-113">サンプル コード: コントロールの計算モード</span><span class="sxs-lookup"><span data-stu-id="f4544-113">Sample code: Control calculation mode</span></span>

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

## <a name="training-video-manage-calculation-mode"></a><span data-ttu-id="f4544-114">トレーニング ビデオ: 計算モードの管理</span><span class="sxs-lookup"><span data-stu-id="f4544-114">Training video: Manage calculation mode</span></span>

<span data-ttu-id="f4544-115">[Sudhi Ramamurthy が YouTube でこのサンプルを歩くのを見る](https://youtu.be/iw6O8QH01CI).</span><span class="sxs-lookup"><span data-stu-id="f4544-115">[Watch Sudhi Ramamurthy walk through this sample on YouTube](https://youtu.be/iw6O8QH01CI).</span></span>
