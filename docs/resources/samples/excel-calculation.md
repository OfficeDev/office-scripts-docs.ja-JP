---
title: Excel で計算モードを管理する
description: Web 上の Excel でOfficeスクリプトを使用して計算モードを管理する方法について説明します。
ms.date: 03/18/2021
localization_priority: Normal
ms.openlocfilehash: 0239437c7b52dca1fd8d1a4fc66bab7965cbd91a
ms.sourcegitcommit: 5d24e77df70aa2c1c982275d53213c2a9323ff86
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 04/02/2021
ms.locfileid: "51571528"
---
# <a name="manage-calculation-mode-in-excel"></a><span data-ttu-id="7433f-103">Excel で計算モードを管理する</span><span class="sxs-lookup"><span data-stu-id="7433f-103">Manage calculation mode in Excel</span></span>

<span data-ttu-id="7433f-104">このサンプルでは、スクリプトを使用[](/javascript/api/office-scripts/excelscript/excelscript.calculationmode)して、計算モードを使用し、Web 上の Excel でメソッドOfficeします。</span><span class="sxs-lookup"><span data-stu-id="7433f-104">This sample shows how to use the [calculation mode](/javascript/api/office-scripts/excelscript/excelscript.calculationmode) and calculate methods in Excel on the web using Office Scripts.</span></span> <span data-ttu-id="7433f-105">任意の Excel ファイルでスクリプトを試してみてください。</span><span class="sxs-lookup"><span data-stu-id="7433f-105">You can try the script on any Excel file.</span></span>

## <a name="scenario"></a><span data-ttu-id="7433f-106">シナリオ</span><span class="sxs-lookup"><span data-stu-id="7433f-106">Scenario</span></span>

<span data-ttu-id="7433f-107">Web 上の Excel では、API を使用してファイルの計算モードをプログラムで制御できます。</span><span class="sxs-lookup"><span data-stu-id="7433f-107">In Excel on the web, a file's calculation mode can be controlled programmatically using APIs.</span></span> <span data-ttu-id="7433f-108">次のアクションは、スクリプトを使用Officeできます。</span><span class="sxs-lookup"><span data-stu-id="7433f-108">The following actions are possible using Office Scripts.</span></span>

1. <span data-ttu-id="7433f-109">計算モードを取得します。</span><span class="sxs-lookup"><span data-stu-id="7433f-109">Get the calculation mode.</span></span>
1. <span data-ttu-id="7433f-110">計算モードを設定します。</span><span class="sxs-lookup"><span data-stu-id="7433f-110">Set the calculation mode.</span></span>
1. <span data-ttu-id="7433f-111">手動モード (再計算とも呼ばれます) に設定されているファイルの Excel 数式を計算します。</span><span class="sxs-lookup"><span data-stu-id="7433f-111">Calculate Excel formulas for files that are set to the manual mode (also referred to as recalculate).</span></span>

## <a name="sample-code-control-calculation-mode"></a><span data-ttu-id="7433f-112">サンプル コード: コントロールの計算モード</span><span class="sxs-lookup"><span data-stu-id="7433f-112">Sample code: Control calculation mode</span></span>

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

## <a name="training-video-manage-calculation-mode"></a><span data-ttu-id="7433f-113">トレーニング ビデオ: 計算モードの管理</span><span class="sxs-lookup"><span data-stu-id="7433f-113">Training video: Manage calculation mode</span></span>

<span data-ttu-id="7433f-114">[![Web 上の Excel で計算モードを管理する方法について、ステップバイステップのビデオを見る](../../images/calc-mode-vid.jpg)](https://youtu.be/iw6O8QH01CI "Web 上の Excel で計算モードを管理する方法に関するステップバイステップのビデオ")</span><span class="sxs-lookup"><span data-stu-id="7433f-114">[![Watch step-by-step video on how to manage calculation mode in Excel on the web](../../images/calc-mode-vid.jpg)](https://youtu.be/iw6O8QH01CI "Step-by-step video on how to manage calculation mode in Excel on the web")</span></span>
