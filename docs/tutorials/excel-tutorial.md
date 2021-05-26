---
title: Excel on the web で Office スクリプトを記録、編集、作成する
description: 操作レコーダーを使用したスクリプトの記録、ブックへのデータの書き込みなど、Office スクリプトの基本について説明したチュートリアル。
ms.date: 05/23/2021
localization_priority: Priority
ms.openlocfilehash: 19cd7bf6c3120d674553d37a36f45d36f46ee852
ms.sourcegitcommit: 0343e4a9843f7ab6ec99d6ddf955050271b061c7
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 05/26/2021
ms.locfileid: "52655906"
---
# <a name="record-edit-and-create-office-scripts-in-excel-on-the-web"></a><span data-ttu-id="855fb-103">Excel on the web で Office スクリプトを記録、編集、作成する</span><span class="sxs-lookup"><span data-stu-id="855fb-103">Record, edit, and create Office Scripts in Excel on the web</span></span>

<span data-ttu-id="855fb-104">このチュートリアルでは、Excel on the web の Office スクリプトの基本となる記録、編集、書き込みについて説明します。</span><span class="sxs-lookup"><span data-stu-id="855fb-104">This tutorial teaches you the basics of recording, editing, and writing an Office Script for Excel on the web.</span></span> <span data-ttu-id="855fb-105">売上記録ワークシートにいくつか書式設定を適用するスクリプトを記録します。</span><span class="sxs-lookup"><span data-stu-id="855fb-105">You'll record a script that applies some formatting to a sales record worksheet.</span></span> <span data-ttu-id="855fb-106">記録されたスクリプトを編集して、より多くの書式設定を適用し、テーブルを作成して、そのテーブルを並べ替えます。</span><span class="sxs-lookup"><span data-stu-id="855fb-106">You'll then edit the recorded script to apply more formatting, create a table, and sort that table.</span></span> <span data-ttu-id="855fb-107">記録して編集するこのパターンは、Excel のアクションがコードとしてどのように表示されるか確認するための重要なツールです。</span><span class="sxs-lookup"><span data-stu-id="855fb-107">This record-then-edit pattern is an important tool to see what your Excel actions look like as code.</span></span>

## <a name="prerequisites"></a><span data-ttu-id="855fb-108">前提条件</span><span class="sxs-lookup"><span data-stu-id="855fb-108">Prerequisites</span></span>

[!INCLUDE [Tutorial prerequisites](../includes/tutorial-prerequisites.md)]

> [!IMPORTANT]
> <span data-ttu-id="855fb-109">このチュートリアルは、JavaScript や TypeScript について初級から中級レベルの知識を持つユーザーを対象としています。</span><span class="sxs-lookup"><span data-stu-id="855fb-109">This tutorial is intended for people with beginner to intermediate-level knowledge of JavaScript or TypeScript.</span></span> <span data-ttu-id="855fb-110">JavaScript を使い慣れていない場合は、「[Mozilla の JavaScript チュートリアル](https://developer.mozilla.org/docs/Web/JavaScript/Guide/Introduction)」から始めることをお勧めします。</span><span class="sxs-lookup"><span data-stu-id="855fb-110">If you're new to JavaScript, we recommend starting with the [Mozilla JavaScript tutorial](https://developer.mozilla.org/docs/Web/JavaScript/Guide/Introduction).</span></span> <span data-ttu-id="855fb-111">スクリプト環境の詳細については、「[Office スクリプト コード エディターの環境](../overview/code-editor-environment.md)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="855fb-111">Visit [Office Scripts Code Editor environment](../overview/code-editor-environment.md) to learn more about the script environment.</span></span>

## <a name="add-data-and-record-a-basic-script"></a><span data-ttu-id="855fb-112">データを追加し、基本スクリプトを記録する</span><span class="sxs-lookup"><span data-stu-id="855fb-112">Add data and record a basic script</span></span>

<span data-ttu-id="855fb-113">まず、いくらかのデータと、最初の小さなスクリプトが必要です。</span><span class="sxs-lookup"><span data-stu-id="855fb-113">First, we'll need some data and a small starting script.</span></span>

1. <span data-ttu-id="855fb-114">Excel for the Web で新しいブックを作成します。</span><span class="sxs-lookup"><span data-stu-id="855fb-114">Create a new workbook in Excel for the Web.</span></span>
2. <span data-ttu-id="855fb-115">次の果物売上データをコピーし、ワークシートのセル **A1** から始まるセル範囲に貼り付けます。</span><span class="sxs-lookup"><span data-stu-id="855fb-115">Copy the following fruit sales data and paste it into the worksheet, starting at cell **A1**.</span></span>

    |<span data-ttu-id="855fb-116">果物</span><span class="sxs-lookup"><span data-stu-id="855fb-116">Fruit</span></span> |<span data-ttu-id="855fb-117">2018</span><span class="sxs-lookup"><span data-stu-id="855fb-117">2018</span></span> |<span data-ttu-id="855fb-118">2019</span><span class="sxs-lookup"><span data-stu-id="855fb-118">2019</span></span> |
    |:---|:---|:---|
    |<span data-ttu-id="855fb-119">オレンジ</span><span class="sxs-lookup"><span data-stu-id="855fb-119">Oranges</span></span> |<span data-ttu-id="855fb-120">1000</span><span class="sxs-lookup"><span data-stu-id="855fb-120">1000</span></span> |<span data-ttu-id="855fb-121">1200</span><span class="sxs-lookup"><span data-stu-id="855fb-121">1200</span></span> |
    |<span data-ttu-id="855fb-122">レモン</span><span class="sxs-lookup"><span data-stu-id="855fb-122">Lemons</span></span> |<span data-ttu-id="855fb-123">800</span><span class="sxs-lookup"><span data-stu-id="855fb-123">800</span></span> |<span data-ttu-id="855fb-124">900</span><span class="sxs-lookup"><span data-stu-id="855fb-124">900</span></span> |
    |<span data-ttu-id="855fb-125">ライム</span><span class="sxs-lookup"><span data-stu-id="855fb-125">Limes</span></span> |<span data-ttu-id="855fb-126">600</span><span class="sxs-lookup"><span data-stu-id="855fb-126">600</span></span> |<span data-ttu-id="855fb-127">500</span><span class="sxs-lookup"><span data-stu-id="855fb-127">500</span></span> |
    |<span data-ttu-id="855fb-128">グレープフルーツ</span><span class="sxs-lookup"><span data-stu-id="855fb-128">Grapefruits</span></span> |<span data-ttu-id="855fb-129">900</span><span class="sxs-lookup"><span data-stu-id="855fb-129">900</span></span> |<span data-ttu-id="855fb-130">700</span><span class="sxs-lookup"><span data-stu-id="855fb-130">700</span></span> |

3. <span data-ttu-id="855fb-131">**[自動化]** タブを開きます。**[自動化]** タブが表示されていない場合は、ドロップダウン矢印を押して、リボンのオーバーフローを確認します。</span><span class="sxs-lookup"><span data-stu-id="855fb-131">Open the **Automate** tab. If you don't see the **Automate** tab, check the ribbon overflow by pressing the drop-down arrow.</span></span> <span data-ttu-id="855fb-132">それでも表示されない場合は、[「Office スクリプトのトラブルシューティング」](../testing/troubleshooting.md#automate-tab-not-appearing-or-office-scripts-unavailable)の記事の説明に従います。</span><span class="sxs-lookup"><span data-stu-id="855fb-132">If it's still not there, follow the advice in the article [Troubleshoot Office Scripts](../testing/troubleshooting.md#automate-tab-not-appearing-or-office-scripts-unavailable).</span></span>
4. <span data-ttu-id="855fb-133">**[操作を記録する]** ボタンを押します。</span><span class="sxs-lookup"><span data-stu-id="855fb-133">Press the **Record Actions** button.</span></span>
5. <span data-ttu-id="855fb-134">セル **A2:C2** ("オレンジ" 行) を選択し、塗りつぶしの色をオレンジ色に設定します。</span><span class="sxs-lookup"><span data-stu-id="855fb-134">Select cells **A2:C2** (the "Oranges" row) and set the fill color to orange.</span></span>
6. <span data-ttu-id="855fb-135">**[停止]** ボタンを押して、記録を停止します。</span><span class="sxs-lookup"><span data-stu-id="855fb-135">Stop the recording by pressing the **Stop** button.</span></span>

    <span data-ttu-id="855fb-136">ワークシートは次のようになります (色が違っていても問題ありません)。</span><span class="sxs-lookup"><span data-stu-id="855fb-136">Your worksheet should look like this (don't worry if the color is different):</span></span>

    :::image type="content" source="../images/tutorial-1.png" alt-text="&quot;オレンジ&quot; を含む行がオレンジ色で強調表示された、フルーツの売上データ行を示すワークシート":::

## <a name="edit-an-existing-script"></a><span data-ttu-id="855fb-138">既存のスクリプトを編集する</span><span class="sxs-lookup"><span data-stu-id="855fb-138">Edit an existing script</span></span>

<span data-ttu-id="855fb-139">前のスクリプトでは、"オレンジ" の行がオレンジ色になります。</span><span class="sxs-lookup"><span data-stu-id="855fb-139">The previous script colored the "Oranges" row to be orange.</span></span> <span data-ttu-id="855fb-140">"レモン" の行に黄色を追加しましょう。</span><span class="sxs-lookup"><span data-stu-id="855fb-140">Let's add a yellow row for the "Lemons".</span></span>

1. <span data-ttu-id="855fb-141">[**詳細**] ウィンドウを開き、[**編集**] ボタンを押します。</span><span class="sxs-lookup"><span data-stu-id="855fb-141">From the now-open **Details** pane, press the **Edit** button.</span></span>
2. <span data-ttu-id="855fb-142">次のようなコードが表示されるはずです。</span><span class="sxs-lookup"><span data-stu-id="855fb-142">You should see something similar to this code:</span></span>

    ```TypeScript
    function main(workbook: ExcelScript.Workbook) {
      // Set fill color to FFC000 for range Sheet1!A2:C2
      let selectedSheet = workbook.getActiveWorksheet();
      selectedSheet.getRange("A2:C2").getFormat().getFill().setColor("FFC000");
    }
    ```

    <span data-ttu-id="855fb-143">このコードは、ブックから現在のワークシートを取得します。</span><span class="sxs-lookup"><span data-stu-id="855fb-143">This code gets the current worksheet from the workbook.</span></span> <span data-ttu-id="855fb-144">次に、**A2:C2** の範囲の塗りつぶしの色を設定します。</span><span class="sxs-lookup"><span data-stu-id="855fb-144">Then, it sets the fill color of the range **A2:C2**.</span></span>

    <span data-ttu-id="855fb-145">範囲は、Excel on the web の Office スクリプトの基本となる部分です。</span><span class="sxs-lookup"><span data-stu-id="855fb-145">Ranges are a fundamental part of Office Scripts in Excel on the web.</span></span> <span data-ttu-id="855fb-146">範囲とは、隣接するセルからなる四角形のブロックで、値、数式、書式設定が含まれます。</span><span class="sxs-lookup"><span data-stu-id="855fb-146">A range is a contiguous, rectangular block of cells that contains values, formula, and formatting.</span></span> <span data-ttu-id="855fb-147">範囲はセルの基本構造であり、スクリプト タスクの大部分は範囲を指定することにより実行します。</span><span class="sxs-lookup"><span data-stu-id="855fb-147">They are the basic structure of cells through which you'll perform most of your scripting tasks.</span></span>

3. <span data-ttu-id="855fb-148">次の行をスクリプトの最後 (`color` の設定箇所と末尾の `}` の間) に追加します。</span><span class="sxs-lookup"><span data-stu-id="855fb-148">Add the following line to the end of the script (between where the `color` is set and the closing `}`):</span></span>

    ```TypeScript
    selectedSheet.getRange("A3:C3").getFormat().getFill().setColor("yellow");
    ```

4. <span data-ttu-id="855fb-149">**[実行]** を押して、スクリプトをテストします。</span><span class="sxs-lookup"><span data-stu-id="855fb-149">Test the script by pressing **Run**.</span></span> <span data-ttu-id="855fb-150">ブックは次のように表示されるはずです。</span><span class="sxs-lookup"><span data-stu-id="855fb-150">Your workbook should now look like this:</span></span>

    :::image type="content" source="../images/tutorial-2.png" alt-text="&quot;オレンジ&quot; の行はオレンジ色、&quot;レモン&quot; の行は黄色で強調表示されているフルーツの売上データ行を示すワークシート":::

## <a name="create-a-table"></a><span data-ttu-id="855fb-152">テーブルを作成する</span><span class="sxs-lookup"><span data-stu-id="855fb-152">Create a table</span></span>

<span data-ttu-id="855fb-153">この果物売上データをテーブルに変換しましょう。</span><span class="sxs-lookup"><span data-stu-id="855fb-153">Let's convert this fruit sales data into a table.</span></span> <span data-ttu-id="855fb-154">プロセス全体でスクリプトを使用します。</span><span class="sxs-lookup"><span data-stu-id="855fb-154">We'll use our script for the entire process.</span></span>

1. <span data-ttu-id="855fb-155">次の行をスクリプトの最後 (末尾の `}` の前) に追加します。</span><span class="sxs-lookup"><span data-stu-id="855fb-155">Add the following line to the end of the script (before the closing `}`):</span></span>

    ```TypeScript
    let table = selectedSheet.addTable("A1:C5", true);
    ```

2. <span data-ttu-id="855fb-156">この呼び出しは `Table` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="855fb-156">That call returns a `Table` object.</span></span> <span data-ttu-id="855fb-157">そのテーブルを使用して、データを並べ替えましょう。</span><span class="sxs-lookup"><span data-stu-id="855fb-157">Let's use that table to sort the data.</span></span> <span data-ttu-id="855fb-158">"果物" 列の値に基づいて、データを昇順で並べ替えます。</span><span class="sxs-lookup"><span data-stu-id="855fb-158">We'll sort the data in ascending order based on the values in the "Fruit" column.</span></span> <span data-ttu-id="855fb-159">次の行を、テーブル作成の後に追加します。</span><span class="sxs-lookup"><span data-stu-id="855fb-159">Add the following line after the table creation:</span></span>

    ```TypeScript
    table.getSort().apply([{ key: 0, ascending: true }]);
    ```

    <span data-ttu-id="855fb-160">スクリプトは次のようになります。</span><span class="sxs-lookup"><span data-stu-id="855fb-160">Your script should look like this:</span></span>

    ```TypeScript
    function main(workbook: ExcelScript.Workbook) {
        // Set fill color to FFC000 for range Sheet1!A2:C2
        let selectedSheet = workbook.getActiveWorksheet();
        selectedSheet.getRange("A2:C2").getFormat().getFill().setColor("FFC000");
        selectedSheet.getRange("A3:C3").getFormat().getFill().setColor("yellow");
        let table = selectedSheet.addTable("A1:C5", true);
        table.getSort().apply([{ key: 0, ascending: true }]);
    }
    ```

    <span data-ttu-id="855fb-161">テーブルには `TableSort` オブジェクトがあり、`Table.getSort` メソッドを使用してアクセスできます。</span><span class="sxs-lookup"><span data-stu-id="855fb-161">Tables have a `TableSort` object, accessed through the `Table.getSort` method.</span></span> <span data-ttu-id="855fb-162">そのオブジェクトに並べ替え条件を適用できます。</span><span class="sxs-lookup"><span data-stu-id="855fb-162">You can apply sorting criteria to that object.</span></span> <span data-ttu-id="855fb-163">`apply` メソッドは、`SortField` オブジェクトの配列を受け取ります。</span><span class="sxs-lookup"><span data-stu-id="855fb-163">The `apply` method takes in an array of `SortField` objects.</span></span> <span data-ttu-id="855fb-164">今回は、並べ替え条件が 1 つだけなので、`SortField` を 1 つだけ使用します。</span><span class="sxs-lookup"><span data-stu-id="855fb-164">In this case, we only have one sorting criteria, so we only use one `SortField`.</span></span> <span data-ttu-id="855fb-165">`key: 0` は、並べ替えを定義する値を含む列を "0" (テーブルの 1 列目。この例では **A**) に設定します。</span><span class="sxs-lookup"><span data-stu-id="855fb-165">`key: 0` sets the column with the sort-defining values to "0" (which is the first column on the table, **A** in this case).</span></span> <span data-ttu-id="855fb-166">`ascending: true` は、昇順 (降順ではなく) にデータを並べ替えます。</span><span class="sxs-lookup"><span data-stu-id="855fb-166">`ascending: true` sorts the data in ascending order (instead of descending order).</span></span>

3. <span data-ttu-id="855fb-p111">スクリプトを実行します。テーブルが次のように表示されます。</span><span class="sxs-lookup"><span data-stu-id="855fb-p111">Run the script. You should see a table like this:</span></span>

    :::image type="content" source="../images/tutorial-3.png" alt-text="並べ替えされたフルーツの売上テーブルを示すワークシート":::

    > [!NOTE]
    > <span data-ttu-id="855fb-170">スクリプトを再実行すると、エラーが表示されます。</span><span class="sxs-lookup"><span data-stu-id="855fb-170">If you re-run the script, you'll get an error.</span></span> <span data-ttu-id="855fb-171">これは、テーブルの上に別のテーブルを重ねて作成することはできないためです。</span><span class="sxs-lookup"><span data-stu-id="855fb-171">This is because you cannot create a table on top of another table.</span></span> <span data-ttu-id="855fb-172">ただし、別のワークシートやブックでスクリプトを実行することはできます。</span><span class="sxs-lookup"><span data-stu-id="855fb-172">However, you can run the script on a different worksheet or workbook.</span></span>

### <a name="re-run-the-script"></a><span data-ttu-id="855fb-173">スクリプトを再実行する</span><span class="sxs-lookup"><span data-stu-id="855fb-173">Re-run the script</span></span>

1. <span data-ttu-id="855fb-174">現在のブックに新しいワークシートを作成します。</span><span class="sxs-lookup"><span data-stu-id="855fb-174">Create a new worksheet in the current workbook.</span></span>
2. <span data-ttu-id="855fb-175">このチュートリアルの最初にある果物のデータをコピーし、新しいワークシートのセル **A1** から始まるセル範囲に貼り付けます。</span><span class="sxs-lookup"><span data-stu-id="855fb-175">Copy the fruit data from the beginning of the tutorial and paste it into the new worksheet, starting at cell **A1**.</span></span>
3. <span data-ttu-id="855fb-176">スクリプトを実行します。</span><span class="sxs-lookup"><span data-stu-id="855fb-176">Run the script.</span></span>

## <a name="next-steps"></a><span data-ttu-id="855fb-177">次の手順</span><span class="sxs-lookup"><span data-stu-id="855fb-177">Next steps</span></span>

<span data-ttu-id="855fb-178">チュートリアルの「[Excel on the web で Office スクリプトを使用してブックのデータを読み取る](excel-read-tutorial.md)」を完了します。</span><span class="sxs-lookup"><span data-stu-id="855fb-178">Complete the [Read workbook data with Office Scripts in Excel on the web](excel-read-tutorial.md) tutorial.</span></span> <span data-ttu-id="855fb-179">このチュートリアルでは、Office スクリプトを使用してブックのデータを読み取る方法について説明します。</span><span class="sxs-lookup"><span data-stu-id="855fb-179">It teaches you how to read data from a workbook with an Office Script.</span></span>
