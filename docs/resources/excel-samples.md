---
title: Excel on the web の Office スクリプトのサンプルスクリプト
description: Web 上の Excel の Office スクリプトで使用するコードサンプルのコレクションです。
ms.date: 02/19/2020
localization_priority: Normal
ms.openlocfilehash: abb4064dfde8b644035e725832e481e6463e979e
ms.sourcegitcommit: b075eed5a6f275274fbbf6d62633219eac416f26
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 03/10/2020
ms.locfileid: "42700418"
---
# <a name="sample-scripts-for-office-scripts-in-excel-on-the-web-preview"></a><span data-ttu-id="e4e6d-103">Web 上の Excel での Office スクリプトのサンプルスクリプト (プレビュー)</span><span class="sxs-lookup"><span data-stu-id="e4e6d-103">Sample scripts for Office Scripts in Excel on the web (preview)</span></span>

<span data-ttu-id="e4e6d-104">次のサンプルは、独自のブックで試すことができる簡単なスクリプトです。</span><span class="sxs-lookup"><span data-stu-id="e4e6d-104">The following samples are simple scripts for you to try on your own workbooks.</span></span> <span data-ttu-id="e4e6d-105">Web 上の Excel で使用するには、次のようにします。</span><span class="sxs-lookup"><span data-stu-id="e4e6d-105">To use them in Excel on the web:</span></span>

1. <span data-ttu-id="e4e6d-106">[**自動化**] タブを開きます。</span><span class="sxs-lookup"><span data-stu-id="e4e6d-106">Open the **Automate** tab.</span></span>
2. <span data-ttu-id="e4e6d-107">**コードエディター**を押します。</span><span class="sxs-lookup"><span data-stu-id="e4e6d-107">Press **Code Editor**.</span></span>
3. <span data-ttu-id="e4e6d-108">コードエディターの作業ウィンドウで、[**新しいスクリプト**] をクリックします。</span><span class="sxs-lookup"><span data-stu-id="e4e6d-108">Press **New Script** in the Code Editor's task pane.</span></span>
4. <span data-ttu-id="e4e6d-109">スクリプト全体を、選択したサンプルに置き換えます。</span><span class="sxs-lookup"><span data-stu-id="e4e6d-109">Replace the entire script with the sample of your choice.</span></span>
5. <span data-ttu-id="e4e6d-110">コードエディターの作業ウィンドウで、[**実行**] をクリックします。</span><span class="sxs-lookup"><span data-stu-id="e4e6d-110">Press **Run** in the Code Editor's task pane.</span></span>

[!INCLUDE [Preview note](../includes/preview-note.md)]

## <a name="scripting-basics"></a><span data-ttu-id="e4e6d-111">スクリプトの基礎</span><span class="sxs-lookup"><span data-stu-id="e4e6d-111">Scripting basics</span></span>

<span data-ttu-id="e4e6d-112">これらのサンプルでは、Office スクリプトの基本的な構成要素を示します。</span><span class="sxs-lookup"><span data-stu-id="e4e6d-112">These samples demonstrate fundamental building blocks for Office Scripts.</span></span> <span data-ttu-id="e4e6d-113">これらをスクリプトに追加して、ソリューションを拡張し、一般的な問題を解決します。</span><span class="sxs-lookup"><span data-stu-id="e4e6d-113">Add these to your scripts to extend your solution and solve common problems.</span></span>

### <a name="read-and-log-one-cell"></a><span data-ttu-id="e4e6d-114">1つのセルを読み取り、ログに記録する</span><span class="sxs-lookup"><span data-stu-id="e4e6d-114">Read and log one cell</span></span>

<span data-ttu-id="e4e6d-115">この例では、 **A1**の値を読み取り、コンソールに出力します。</span><span class="sxs-lookup"><span data-stu-id="e4e6d-115">This sample reads the value of **A1** and prints it to the console.</span></span>

``` TypeScript
async function main(context: Excel.RequestContext) {
  // Get the current worksheet.
  let selectedSheet = context.workbook.worksheets.getActiveWorksheet();

  // Get the value of cell A1.
  let range = selectedSheet.getRange("A1");
  range.load("values");
  await context.sync();

  // Print the value of A1.
  console.log(range.values);
}
```

### <a name="work-with-dates"></a><span data-ttu-id="e4e6d-116">日付の操作</span><span class="sxs-lookup"><span data-stu-id="e4e6d-116">Work with dates</span></span>

<span data-ttu-id="e4e6d-117">この例では、JavaScript の[date](https://developer.mozilla.org/docs/web/javascript/reference/global_objects/date)オブジェクトを使用して、現在の日付と時刻を取得し、その値を作業中のワークシート内の2つのセルに書き込みます。</span><span class="sxs-lookup"><span data-stu-id="e4e6d-117">This sample uses the JavaScript [Date](https://developer.mozilla.org/docs/web/javascript/reference/global_objects/date) object to get the current date and time and then writes those values to two cells in the active worksheet.</span></span>

```TypeScript
async function main(context: Excel.RequestContext) {
  // Get the cells at A1 and B1.
  let dateRange = context.workbook.worksheets.getActiveWorksheet().getRange("A1");
  let timeRange = context.workbook.worksheets.getActiveWorksheet().getRange("B1");

  // Get the current date and time with the JavaScript Date object.
  let date = new Date(Date.now());

  // Add the date string to A1.
  dateRange.values = [[date.toLocaleDateString()]];
  
  // Add the time string to B1.
  timeRange.values = [[date.toLocaleTimeString()]];
}
```

## <a name="display-data"></a><span data-ttu-id="e4e6d-118">データの表示</span><span class="sxs-lookup"><span data-stu-id="e4e6d-118">Display data</span></span>

<span data-ttu-id="e4e6d-119">これらのサンプルは、ワークシートデータを操作し、ユーザーにより良い表示や組織を提供する方法を示しています。</span><span class="sxs-lookup"><span data-stu-id="e4e6d-119">These samples demonstrate how to work with worksheet data and provide users with a better view or organization.</span></span>

### <a name="apply-conditional-formatting"></a><span data-ttu-id="e4e6d-120">条件付き書式の適用</span><span class="sxs-lookup"><span data-stu-id="e4e6d-120">Apply conditional formatting</span></span>

<span data-ttu-id="e4e6d-121">この例では、ワークシートで現在使用されている範囲に条件付き書式を適用します。</span><span class="sxs-lookup"><span data-stu-id="e4e6d-121">This sample applies conditional formatting to the currently used range in the worksheet.</span></span> <span data-ttu-id="e4e6d-122">条件付き書式は、値の上位10% の緑の塗りつぶしです。</span><span class="sxs-lookup"><span data-stu-id="e4e6d-122">The conditional formatting is a green fill for the top 10% of values.</span></span>

```TypeScript
async function main(context: Excel.RequestContext) {
  // Get the current worksheet.
  let selectedSheet = context.workbook.worksheets.getActiveWorksheet();

  // Get the used range in the worksheet.
  let range = selectedSheet.getUsedRange();

  // Set the fill color to green for the top 10% of values in the range.
  let conditionalFormat = range.conditionalFormats.add(Excel.ConditionalFormatType.topBottom);
  conditionalFormat.topBottom.format.fill.color = "green";
  conditionalFormat.topBottom.rule = {
    rank: 10, // The percentage threshold.
    type: Excel.ConditionalTopBottomCriterionType.topPercent // The type of the top/bottom condition.
  };
}
```

### <a name="create-a-sorted-table"></a><span data-ttu-id="e4e6d-123">並べ替えられたテーブルを作成する</span><span class="sxs-lookup"><span data-stu-id="e4e6d-123">Create a sorted table</span></span>

<span data-ttu-id="e4e6d-124">次の使用例は、現在のワークシートの使用範囲から表を作成し、最初の列に基づいて並べ替えます。</span><span class="sxs-lookup"><span data-stu-id="e4e6d-124">This sample creates a table from the current worksheet's used range, then sorts it based on the first column.</span></span>

```TypeScript
async function main(context: Excel.RequestContext) {
  // Get the current worksheet.
  let selectedSheet = context.workbook.worksheets.getActiveWorksheet();

  // Create a table with the used cells.
  let usedRange = selectedSheet.getUsedRange();
  let newTable = selectedSheet.tables.add(usedRange, true);

  // Sort the table using the first column.
  newTable.sort.apply([{ key: 0, ascending: true }]);
}
```

## <a name="collaboration"></a><span data-ttu-id="e4e6d-125">グループ作業</span><span class="sxs-lookup"><span data-stu-id="e4e6d-125">Collaboration</span></span>

<span data-ttu-id="e4e6d-126">これらのサンプルは、コメントなど、Excel のグループ作業関連機能を操作する方法を示しています。</span><span class="sxs-lookup"><span data-stu-id="e4e6d-126">These samples demonstrate how to work with collaboration-related features of Excel, such as comments.</span></span>

### <a name="delete-resolved-comments"></a><span data-ttu-id="e4e6d-127">解決されたコメントの削除</span><span class="sxs-lookup"><span data-stu-id="e4e6d-127">Delete resolved comments</span></span>

<span data-ttu-id="e4e6d-128">この例では、現在のワークシートから解決されたすべてのコメントを削除します。</span><span class="sxs-lookup"><span data-stu-id="e4e6d-128">This sample deletes all resolved comments from the current worksheet.</span></span>

```TypeScript
async function main(context: Excel.RequestContext) {
  // Get the current worksheet.
  let selectedSheet = context.workbook.worksheets.getActiveWorksheet();

  // Get the comments on this worksheet.
  let comments = selectedSheet.comments;
  comments.load("items/resolved");
  await context.sync();

  // Delete the resolved comments.
  comments.items.forEach((comment) => {
      if (comment.resolved) {
          comment.delete();
      }
  });
}
```

## <a name="scenario-samples"></a><span data-ttu-id="e4e6d-129">シナリオサンプル</span><span class="sxs-lookup"><span data-stu-id="e4e6d-129">Scenario samples</span></span>

<span data-ttu-id="e4e6d-130">大規模な現実世界のソリューションを紹介するサンプルについては、「 [Office スクリプトのサンプルシナリオ](scenarios/sample-scenario-overview.md)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="e4e6d-130">For samples showcasing larger, real-world solutions, visit [Sample scenarios for Office Scripts](scenarios/sample-scenario-overview.md).</span></span>

## <a name="suggest-new-samples"></a><span data-ttu-id="e4e6d-131">新しいサンプルを提案する</span><span class="sxs-lookup"><span data-stu-id="e4e6d-131">Suggest new samples</span></span>

<span data-ttu-id="e4e6d-132">新しいサンプルの提案を歓迎します。</span><span class="sxs-lookup"><span data-stu-id="e4e6d-132">We welcome suggestions for new samples.</span></span> <span data-ttu-id="e4e6d-133">他のスクリプト開発者を支援する一般的なシナリオがある場合は、以下のフィードバックセクションでご連絡ください。</span><span class="sxs-lookup"><span data-stu-id="e4e6d-133">If there is a common scenario that would help other script developers, please tell us in the feedback section below.</span></span>
