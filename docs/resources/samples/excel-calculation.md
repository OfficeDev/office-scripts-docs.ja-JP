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
# <a name="manage-calculation-mode-in-excel"></a><span data-ttu-id="cefe6-103">計算モードを管理Excel</span><span class="sxs-lookup"><span data-stu-id="cefe6-103">Manage calculation mode in Excel</span></span>

<span data-ttu-id="cefe6-104">このサンプルでは、スクリプトを使用[して計算](/javascript/api/office-scripts/excelscript/excelscript.calculationmode)モードを使用し、計算モードでメソッドExcel on the web Officeします。</span><span class="sxs-lookup"><span data-stu-id="cefe6-104">This sample shows how to use the [calculation mode](/javascript/api/office-scripts/excelscript/excelscript.calculationmode) and calculate methods in Excel on the web using Office Scripts.</span></span> <span data-ttu-id="cefe6-105">スクリプトは、任意のファイルでExcelできます。</span><span class="sxs-lookup"><span data-stu-id="cefe6-105">You can try the script on any Excel file.</span></span>

## <a name="scenario"></a><span data-ttu-id="cefe6-106">シナリオ</span><span class="sxs-lookup"><span data-stu-id="cefe6-106">Scenario</span></span>

<span data-ttu-id="cefe6-107">このExcel on the web、API を使用してファイルの計算モードをプログラムで制御できます。</span><span class="sxs-lookup"><span data-stu-id="cefe6-107">In Excel on the web, a file's calculation mode can be controlled programmatically using APIs.</span></span> <span data-ttu-id="cefe6-108">次のアクションは、スクリプトを使用Officeできます。</span><span class="sxs-lookup"><span data-stu-id="cefe6-108">The following actions are possible using Office Scripts.</span></span>

1. <span data-ttu-id="cefe6-109">計算モードを取得します。</span><span class="sxs-lookup"><span data-stu-id="cefe6-109">Get the calculation mode.</span></span>
1. <span data-ttu-id="cefe6-110">計算モードを設定します。</span><span class="sxs-lookup"><span data-stu-id="cefe6-110">Set the calculation mode.</span></span>
1. <span data-ttu-id="cefe6-111">手動Excel (再計算とも呼ばれます) に設定されているファイルの数式を計算します。</span><span class="sxs-lookup"><span data-stu-id="cefe6-111">Calculate Excel formulas for files that are set to the manual mode (also referred to as recalculate).</span></span>

## <a name="sample-code-control-calculation-mode"></a><span data-ttu-id="cefe6-112">サンプル コード: コントロールの計算モード</span><span class="sxs-lookup"><span data-stu-id="cefe6-112">Sample code: Control calculation mode</span></span>

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

## <a name="training-video-manage-calculation-mode"></a><span data-ttu-id="cefe6-113">トレーニング ビデオ: 計算モードの管理</span><span class="sxs-lookup"><span data-stu-id="cefe6-113">Training video: Manage calculation mode</span></span>

<span data-ttu-id="cefe6-114">[Sudhi Ramamurthy が YouTube でこのサンプルを歩くのを見る](https://youtu.be/iw6O8QH01CI).</span><span class="sxs-lookup"><span data-stu-id="cefe6-114">[Watch Sudhi Ramamurthy walk through this sample on YouTube](https://youtu.be/iw6O8QH01CI).</span></span>
