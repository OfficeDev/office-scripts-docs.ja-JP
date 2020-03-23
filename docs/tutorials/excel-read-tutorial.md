---
title: Excel on the web で Office スクリプトを使用してブックのデータを読み取る
description: ブックのデータを読み取り、スクリプトでそのデータを評価する方法について説明した Office スクリプトのチュートリアル。
ms.date: 01/27/2020
localization_priority: Priority
ms.openlocfilehash: 42ed0fe5843a78692f9660b873211e3668702164
ms.sourcegitcommit: b075eed5a6f275274fbbf6d62633219eac416f26
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 03/10/2020
ms.locfileid: "42700323"
---
# <a name="read-workbook-data-with-office-scripts-in-excel-on-the-web"></a><span data-ttu-id="2b1c9-103">Excel on the web で Office スクリプトを使用してブックのデータを読み取る</span><span class="sxs-lookup"><span data-stu-id="2b1c9-103">Read workbook data with Office Scripts in Excel on the web</span></span>

<span data-ttu-id="2b1c9-104">このチュートリアルでは、Excel on the web 用の Office スクリプトを使用してブックのデータを読み取る方法について説明します。</span><span class="sxs-lookup"><span data-stu-id="2b1c9-104">This tutorial will teach you how to read data from a workbook with an Office Script for Excel on the web.</span></span> <span data-ttu-id="2b1c9-105">その後、読み取ったデータを編集し、ブックに戻します。</span><span class="sxs-lookup"><span data-stu-id="2b1c9-105">You'll then edit the data you read and put it back in the workbook.</span></span>

> [!TIP]
> <span data-ttu-id="2b1c9-106">Office スクリプトを初めて使用する場合は、チュートリアルの「[Excel on the web で Office スクリプトを記録、編集、作成する](excel-tutorial.md)」から始めることをお勧めします。</span><span class="sxs-lookup"><span data-stu-id="2b1c9-106">If you are new to Office Scripts, we recommend starting with the [Record, edit, and create Office Scripts in Excel on the web](excel-tutorial.md) tutorial.</span></span>

## <a name="prerequisites"></a><span data-ttu-id="2b1c9-107">前提条件</span><span class="sxs-lookup"><span data-stu-id="2b1c9-107">Prerequisites</span></span>

[!INCLUDE [Preview note](../includes/preview-note.md)]

<span data-ttu-id="2b1c9-108">このチュートリアルを開始するには、Office スクリプトへのアクセスが必要です。これには次のものが必要です。</span><span class="sxs-lookup"><span data-stu-id="2b1c9-108">Before starting this tutorial, you'll need access to Office Scripts, which requires the following:</span></span>

- <span data-ttu-id="2b1c9-109">[Excel on the web](https://www.office.com/launch/excel)。</span><span class="sxs-lookup"><span data-stu-id="2b1c9-109">[Excel on the web](https://www.office.com/launch/excel).</span></span>
- <span data-ttu-id="2b1c9-110">[組織に対して Office スクリプトを許可する](https://support.office.com/article/office-scripts-settings-in-m365-19d3c51a-6ca2-40ab-978d-60fa49554dcf)よう管理者に依頼します。これにより、リボンに **[自動化]** タブが追加されます。</span><span class="sxs-lookup"><span data-stu-id="2b1c9-110">Ask your administrator to [enable Office Scripts for your organization](https://support.office.com/article/office-scripts-settings-in-m365-19d3c51a-6ca2-40ab-978d-60fa49554dcf), which adds the **Automate** tab to the ribbon.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="2b1c9-111">このチュートリアルは、JavaScript や TypeScript について初級から中級レベルの知識を持つユーザーを対象としています。</span><span class="sxs-lookup"><span data-stu-id="2b1c9-111">This tutorial is intended for people with beginner to intermediate-level knowledge of JavaScript or TypeScript.</span></span> <span data-ttu-id="2b1c9-112">JavaScript を使い慣れていない場合は、[Mozilla の JavaScript チュートリアル](https://developer.mozilla.org/docs/Web/JavaScript/Guide/Introduction)をご覧になることをお勧めします。</span><span class="sxs-lookup"><span data-stu-id="2b1c9-112">If you're new to JavaScript, we recommend reviewing the [Mozilla JavaScript tutorial](https://developer.mozilla.org/docs/Web/JavaScript/Guide/Introduction).</span></span> <span data-ttu-id="2b1c9-113">スクリプト環境の詳細については、「[Excel on the web の Office スクリプト](../overview/excel.md)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="2b1c9-113">Visit [Office Scripts in Excel on the web](../overview/excel.md) to learn more about the script environment.</span></span>

## <a name="read-a-cell"></a><span data-ttu-id="2b1c9-114">セルを読み取る。</span><span class="sxs-lookup"><span data-stu-id="2b1c9-114">Read a cell</span></span>

<span data-ttu-id="2b1c9-115">操作レコーダーで作成したスクリプトは、ブックに情報を書き込む操作のみを実行できます。</span><span class="sxs-lookup"><span data-stu-id="2b1c9-115">Scripts made with the Action Recorder can only write information to the workbook.</span></span> <span data-ttu-id="2b1c9-116">コード エディターを使用すると、ブックのデータを読み取ることも可能なスクリプトの編集と作成ができます。</span><span class="sxs-lookup"><span data-stu-id="2b1c9-116">With the Code Editor, you can edit and make scripts that also read data from a workbook.</span></span>

<span data-ttu-id="2b1c9-117">データを読み取り、読み取った内容に基づいて動作するスクリプトを作成しましょう。</span><span class="sxs-lookup"><span data-stu-id="2b1c9-117">Let's make a script that reads data and acts based on what was read.</span></span> <span data-ttu-id="2b1c9-118">今回は、サンプルの銀行取引明細書を使用します。</span><span class="sxs-lookup"><span data-stu-id="2b1c9-118">We're going to work with a sample banking statement.</span></span> <span data-ttu-id="2b1c9-119">この明細書は、支払いと貸方がまとまった明細書です。</span><span class="sxs-lookup"><span data-stu-id="2b1c9-119">This statement is a combined checking and credit statement.</span></span> <span data-ttu-id="2b1c9-120">残念ながら、残高の変化が異なる仕方で報告されています。</span><span class="sxs-lookup"><span data-stu-id="2b1c9-120">Unfortunately, they report balance changes differently.</span></span> <span data-ttu-id="2b1c9-121">支払い明細では、収入を負の貸方として記録し、支出を負の借方として記録しています。</span><span class="sxs-lookup"><span data-stu-id="2b1c9-121">The checking statement gives income as positive credit and costs as negative debit.</span></span> <span data-ttu-id="2b1c9-122">貸方明細ではその逆になっています。</span><span class="sxs-lookup"><span data-stu-id="2b1c9-122">The credit statement does the opposite.</span></span>

<span data-ttu-id="2b1c9-123">チュートリアルの残りの部分で、スクリプトを使用してこのデータを正規化します。</span><span class="sxs-lookup"><span data-stu-id="2b1c9-123">Over the rest of the tutorial, we will normalize this data using a script.</span></span> <span data-ttu-id="2b1c9-124">まず、ブックからデータを読み取る方法について説明します。</span><span class="sxs-lookup"><span data-stu-id="2b1c9-124">First, let's learn how to read data from the workbook.</span></span>

1. <span data-ttu-id="2b1c9-125">チュートリアルの残りの部分で使用したブックに新しいワークシートを作成します。</span><span class="sxs-lookup"><span data-stu-id="2b1c9-125">Create a new worksheet in the workbook you've used for the rest of the tutorial.</span></span>
2. <span data-ttu-id="2b1c9-126">次のデータをコピーし、新しいワークシートのセル **A1** から始まるセル範囲に貼り付けます。</span><span class="sxs-lookup"><span data-stu-id="2b1c9-126">Copy the following data and paste it into the new worksheet, starting at cell **A1**.</span></span>

    |<span data-ttu-id="2b1c9-127">日付</span><span class="sxs-lookup"><span data-stu-id="2b1c9-127">Date</span></span> |<span data-ttu-id="2b1c9-128">取引</span><span class="sxs-lookup"><span data-stu-id="2b1c9-128">Account</span></span> |<span data-ttu-id="2b1c9-129">説明</span><span class="sxs-lookup"><span data-stu-id="2b1c9-129">Description</span></span> |<span data-ttu-id="2b1c9-130">借方</span><span class="sxs-lookup"><span data-stu-id="2b1c9-130">Debit</span></span> |<span data-ttu-id="2b1c9-131">貸方</span><span class="sxs-lookup"><span data-stu-id="2b1c9-131">Credit</span></span> |
    |:--|:--|:--|:--|:--|
    |<span data-ttu-id="2b1c9-132">2019/10/10</span><span class="sxs-lookup"><span data-stu-id="2b1c9-132">10/10/2019</span></span> |<span data-ttu-id="2b1c9-133">支払い</span><span class="sxs-lookup"><span data-stu-id="2b1c9-133">Checking</span></span> |<span data-ttu-id="2b1c9-134">Coho Vineyard</span><span class="sxs-lookup"><span data-stu-id="2b1c9-134">Coho Vineyard</span></span> |<span data-ttu-id="2b1c9-135">-20.05</span><span class="sxs-lookup"><span data-stu-id="2b1c9-135">-20.05</span></span> | |
    |<span data-ttu-id="2b1c9-136">2019/10/11</span><span class="sxs-lookup"><span data-stu-id="2b1c9-136">10/11/2019</span></span> |<span data-ttu-id="2b1c9-137">貸方</span><span class="sxs-lookup"><span data-stu-id="2b1c9-137">Credit</span></span> |<span data-ttu-id="2b1c9-138">The Phone Company</span><span class="sxs-lookup"><span data-stu-id="2b1c9-138">The Phone Company</span></span> |<span data-ttu-id="2b1c9-139">99.95</span><span class="sxs-lookup"><span data-stu-id="2b1c9-139">99.95</span></span> | |
    |<span data-ttu-id="2b1c9-140">2019/10/13</span><span class="sxs-lookup"><span data-stu-id="2b1c9-140">10/13/2019</span></span> |<span data-ttu-id="2b1c9-141">貸方</span><span class="sxs-lookup"><span data-stu-id="2b1c9-141">Credit</span></span> |<span data-ttu-id="2b1c9-142">Coho Vineyard</span><span class="sxs-lookup"><span data-stu-id="2b1c9-142">Coho Vineyard</span></span> |<span data-ttu-id="2b1c9-143">154.43</span><span class="sxs-lookup"><span data-stu-id="2b1c9-143">154.43</span></span> | |
    |<span data-ttu-id="2b1c9-144">2019/10/15</span><span class="sxs-lookup"><span data-stu-id="2b1c9-144">10/15/2019</span></span> |<span data-ttu-id="2b1c9-145">支払い</span><span class="sxs-lookup"><span data-stu-id="2b1c9-145">Checking</span></span> |<span data-ttu-id="2b1c9-146">外部預金</span><span class="sxs-lookup"><span data-stu-id="2b1c9-146">External Deposit</span></span> | |<span data-ttu-id="2b1c9-147">1000</span><span class="sxs-lookup"><span data-stu-id="2b1c9-147">1000</span></span> |
    |<span data-ttu-id="2b1c9-148">2019/10/20</span><span class="sxs-lookup"><span data-stu-id="2b1c9-148">10/20/2019</span></span> |<span data-ttu-id="2b1c9-149">貸方</span><span class="sxs-lookup"><span data-stu-id="2b1c9-149">Credit</span></span> |<span data-ttu-id="2b1c9-150">Coho Vineyard - 返金</span><span class="sxs-lookup"><span data-stu-id="2b1c9-150">Coho Vineyard - Refund</span></span> | |<span data-ttu-id="2b1c9-151">-35.45</span><span class="sxs-lookup"><span data-stu-id="2b1c9-151">-35.45</span></span> |
    |<span data-ttu-id="2b1c9-152">2019/10/25</span><span class="sxs-lookup"><span data-stu-id="2b1c9-152">10/25/2019</span></span> |<span data-ttu-id="2b1c9-153">支払い</span><span class="sxs-lookup"><span data-stu-id="2b1c9-153">Checking</span></span> |<span data-ttu-id="2b1c9-154">Best For You Organics Company</span><span class="sxs-lookup"><span data-stu-id="2b1c9-154">Best For You Organics Company</span></span> | <span data-ttu-id="2b1c9-155">-85.64</span><span class="sxs-lookup"><span data-stu-id="2b1c9-155">-85.64</span></span> | |
    |<span data-ttu-id="2b1c9-156">2019/11/01</span><span class="sxs-lookup"><span data-stu-id="2b1c9-156">11/01/2019</span></span> |<span data-ttu-id="2b1c9-157">支払い</span><span class="sxs-lookup"><span data-stu-id="2b1c9-157">Checking</span></span> |<span data-ttu-id="2b1c9-158">外部預金</span><span class="sxs-lookup"><span data-stu-id="2b1c9-158">External Deposit</span></span> | |<span data-ttu-id="2b1c9-159">1000</span><span class="sxs-lookup"><span data-stu-id="2b1c9-159">1000</span></span> |

3. <span data-ttu-id="2b1c9-160">**[コード エディター]** を開き、**[新しいスクリプト]** を選択します。</span><span class="sxs-lookup"><span data-stu-id="2b1c9-160">Open the **Code Editor** and select **New Script**.</span></span>
4. <span data-ttu-id="2b1c9-161">書式設定をクリーンアップします。</span><span class="sxs-lookup"><span data-stu-id="2b1c9-161">Let's clean up the formatting.</span></span> <span data-ttu-id="2b1c9-162">これは財務ドキュメントなので、**[借方]** 列と **[貸方]** 列の数値の書式設定を変更して、値がドル金額として表示されるようにします。</span><span class="sxs-lookup"><span data-stu-id="2b1c9-162">This is a financial document, so let's change the number formatting in the **Debit** and **Credit** columns to show values as dollar amounts.</span></span> <span data-ttu-id="2b1c9-163">さらに、列幅をデータに合わせます。</span><span class="sxs-lookup"><span data-stu-id="2b1c9-163">Let's also fit the column width to the data.</span></span>

    <span data-ttu-id="2b1c9-164">スクリプトの内容を次のコードで置き換えます。</span><span class="sxs-lookup"><span data-stu-id="2b1c9-164">Replace the script contents with the following code:</span></span>

    ```TypeScript
    async function main(context: Excel.RequestContext) {
      // Get the current worksheet.
      let workbook = context.workbook;
      let worksheets = workbook.worksheets;
      let selectedSheet = worksheets.getActiveWorksheet();

      // Format the range to display numerical dollar amounts.
      selectedSheet.getRange("D2:E8").numberFormat = [["$#,##0.00"]];

      // Fit the width of all the used columns to the data.
      selectedSheet.getUsedRange().format.autofitColumns();
    }
    ```

5. <span data-ttu-id="2b1c9-165">では、いずれかの数値列の値を読み取ってみましょう。</span><span class="sxs-lookup"><span data-stu-id="2b1c9-165">Now let's read a value from one of the number columns.</span></span> <span data-ttu-id="2b1c9-166">次のコードをスクリプトの最後に追加します。</span><span class="sxs-lookup"><span data-stu-id="2b1c9-166">Add the following code to the end of the script:</span></span>

    ```TypeScript
    // Get the value of cell D2.
    let range = selectedSheet.getRange("D2");
    range.load("values");
    await context.sync();
  
    // Print the value of D2.
    console.log(range.values);
    ```

    <span data-ttu-id="2b1c9-167">`load` と `sync` への呼び出しに注目してください。</span><span class="sxs-lookup"><span data-stu-id="2b1c9-167">Note the calls to `load` and `sync`.</span></span> <span data-ttu-id="2b1c9-168">これらのメソッドの詳細については、「[Excel on the web での Office スクリプトのスクリプトの基本事項](../develop/scripting-fundamentals.md#sync-and-load)」で説明します。</span><span class="sxs-lookup"><span data-stu-id="2b1c9-168">You can learn the details of those methods in [Scripting Fundamentals for Office Scripts in Excel on the web](../develop/scripting-fundamentals.md#sync-and-load).</span></span> <span data-ttu-id="2b1c9-169">ここでは、データの読み取りを要求し、スクリプトとブックを同期してそのデータを読み取る必要があることを覚えておいてください。</span><span class="sxs-lookup"><span data-stu-id="2b1c9-169">For now, know that you must request data to be read and then sync your script with the workbook to read that data.</span></span>

6. <span data-ttu-id="2b1c9-170">スクリプトを実行します。</span><span class="sxs-lookup"><span data-stu-id="2b1c9-170">Run the script.</span></span>
7. <span data-ttu-id="2b1c9-171">コンソールを開きます。</span><span class="sxs-lookup"><span data-stu-id="2b1c9-171">Open the console.</span></span> <span data-ttu-id="2b1c9-172">**省略記号**のメニューを選択し、**[Logs...](ログ...)** を押します。</span><span class="sxs-lookup"><span data-stu-id="2b1c9-172">Go to the **Ellipses** menu and press **Logs...**.</span></span>
8. <span data-ttu-id="2b1c9-173">コンソールに `[Array[1]]` が表示されます。</span><span class="sxs-lookup"><span data-stu-id="2b1c9-173">You should see `[Array[1]]` in the console.</span></span> <span data-ttu-id="2b1c9-174">範囲は 2 次元のデータ配列であるため、これは数値ではありません。</span><span class="sxs-lookup"><span data-stu-id="2b1c9-174">This is not a number because ranges are two-dimensional arrays of data.</span></span> <span data-ttu-id="2b1c9-175">この 2 次元の範囲は、コンソールに直接ログ記録されます。</span><span class="sxs-lookup"><span data-stu-id="2b1c9-175">That two-dimensional range is being logged to the console directly.</span></span> <span data-ttu-id="2b1c9-176">コード エディターを使用すると、この配列の内容を表示できます。</span><span class="sxs-lookup"><span data-stu-id="2b1c9-176">Luckily, the Code Editor does let you see the contents of the array.</span></span>
9. <span data-ttu-id="2b1c9-177">2 次元の配列がコンソールにログ記録すると、各行の下に列の値がグループ化されます。</span><span class="sxs-lookup"><span data-stu-id="2b1c9-177">When a two-dimensional array is logged to the console, it groups column values under each row.</span></span> <span data-ttu-id="2b1c9-178">青い三角形を押して、配列のログを展開します。</span><span class="sxs-lookup"><span data-stu-id="2b1c9-178">Expand the array log by pressing the blue triangle.</span></span>
10. <span data-ttu-id="2b1c9-179">新たに表示された青い三角形を押して、配列の第 2 レベルを展開します。</span><span class="sxs-lookup"><span data-stu-id="2b1c9-179">Expand the second level of the array by pressing the newly revealed blue triangle.</span></span> <span data-ttu-id="2b1c9-180">次のように表示されるはずです。</span><span class="sxs-lookup"><span data-stu-id="2b1c9-180">You should now see this:</span></span>

    ![出力 "-20.05" が 2 つの配列の下に入れ子になって表示されているコンソール ログ。](../images/tutorial-4.png)

## <a name="modify-the-value-of-a-cell"></a><span data-ttu-id="2b1c9-182">セルの値を変更する</span><span class="sxs-lookup"><span data-stu-id="2b1c9-182">Modify the value of a cell</span></span>

<span data-ttu-id="2b1c9-183">データを読み取れたので、そのデータを使用してブックを変更しましょう。</span><span class="sxs-lookup"><span data-stu-id="2b1c9-183">Now that we can read data, let's use that data to modify the workbook.</span></span> <span data-ttu-id="2b1c9-184">セル **D2** の値を、`Math.abs` 関数を使用して正の値にします。</span><span class="sxs-lookup"><span data-stu-id="2b1c9-184">We'll make the value of the cell **D2** positive with the `Math.abs` function.</span></span> <span data-ttu-id="2b1c9-185">[Math](https://developer.mozilla.org/docs/web/javascript/reference/global_objects/math) オブジェクトには、スクリプトでアクセスできる多くの関数が含まれています。</span><span class="sxs-lookup"><span data-stu-id="2b1c9-185">The [Math](https://developer.mozilla.org/docs/web/javascript/reference/global_objects/math) object contains many functions to which your scripts have access.</span></span> <span data-ttu-id="2b1c9-186">`Math` および他の組み込みオブジェクトの詳細については、「[Office スクリプトでの組み込みの JavaScript オブジェクトの使用](../develop/javascript-objects.md)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="2b1c9-186">More information about `Math` and other built-in objects can be found at [Using built-in JavaScript objects in Office Scripts](../develop/javascript-objects.md).</span></span>

1. <span data-ttu-id="2b1c9-187">次のコードをスクリプトの最後に追加します。</span><span class="sxs-lookup"><span data-stu-id="2b1c9-187">Add the following code to the end of the script:</span></span>

    ```TypeScript
    // Run the `Math.abs` function with the value at D2 and apply that value back to D2.
    let positiveValue = Math.abs(range.values[0][0]);
    range.values = [[positiveValue]];
    ```

2. <span data-ttu-id="2b1c9-188">セル **D2** の値が正の値になります。</span><span class="sxs-lookup"><span data-stu-id="2b1c9-188">The value of cell **D2** should now be positive.</span></span>

## <a name="modify-the-values-of-a-column"></a><span data-ttu-id="2b1c9-189">列の値を変更する</span><span class="sxs-lookup"><span data-stu-id="2b1c9-189">Modify the values of a column</span></span>

<span data-ttu-id="2b1c9-190">1 つのセルの読み取り方法と書き込み方法がわかったので、スクリプトを一般化して、**[借方]** 列と **[貸方]** 列全体を操作できるようにしましょう。</span><span class="sxs-lookup"><span data-stu-id="2b1c9-190">Now that we know how to read and write to a single cell, let's generalize the script to work on the entire **Debit** and **Credit** columns.</span></span>

1. <span data-ttu-id="2b1c9-191">1 つのセルにのみ影響するコード (前述の絶対値コード) を削除します。すると、スクリプトは次のようになります。</span><span class="sxs-lookup"><span data-stu-id="2b1c9-191">Remove the code that affects only a single cell (the previous absolute value code), such that your script now looks like this:</span></span>

    ```TypeScript
    async function main(context: Excel.RequestContext) {
      // Get the current worksheet.
      let workbook = context.workbook;
      let worksheets = workbook.worksheets;
      let selectedSheet = worksheets.getActiveWorksheet();

      // Format the range to display numerical dollar amounts.
      selectedSheet.getRange("D2:E8").numberFormat = [["$#,##0.00"]];

      // Fit the width of all the used columns to the data.
      selectedSheet.getUsedRange().format.autofitColumns();
    }
    ```

2. <span data-ttu-id="2b1c9-192">最後の 2 つの列の行を反復処理するループを追加します。</span><span class="sxs-lookup"><span data-stu-id="2b1c9-192">Add a loop that iterates through the rows in the last two columns.</span></span> <span data-ttu-id="2b1c9-193">スクリプトにより、各セルの値が現在の値の絶対値に設定されます。</span><span class="sxs-lookup"><span data-stu-id="2b1c9-193">For each cell, the script sets the value to the current value's absolute value.</span></span>

    <span data-ttu-id="2b1c9-194">セルの位置を定義する配列は 0 から始まることにご注意ください。</span><span class="sxs-lookup"><span data-stu-id="2b1c9-194">Note that the array defining cell locations is zero-based.</span></span> <span data-ttu-id="2b1c9-195">したがって、セル **A1** は `range[0][0]` になります。</span><span class="sxs-lookup"><span data-stu-id="2b1c9-195">That means cell **A1** is `range[0][0]`.</span></span>

    ```TypeScript
    // Get the values of the used range.
    let range = selectedSheet.getUsedRange();
    range.load("rowCount,values");
    await context.sync();

    // Iterate over the fourth and fifth columns and set their values to their absolute value.
    for (let i = 1; i < range.rowCount; i++) {
      // The column at index 3 is column "4" in the worksheet.
      if (range.values[i][3] != 0) {
        let positiveValue = Math.abs(range.values[i][3]);
        selectedSheet.getCell(i, 3).values = [[positiveValue]];
      }

      // The column at index 4 is column "5" in the worksheet.
      if (range.values[i][4] != 0) {
        let positiveValue = Math.abs(range.values[i][4]);
        selectedSheet.getCell(i, 4).values = [[positiveValue]];
      }
    }
    ```

    <span data-ttu-id="2b1c9-196">スクリプトのこの部分は、いくつかの重要なタスクを実行します。</span><span class="sxs-lookup"><span data-stu-id="2b1c9-196">This portion of the script does several important tasks.</span></span> <span data-ttu-id="2b1c9-197">まず、指定された範囲の値と行数を読み込みます。</span><span class="sxs-lookup"><span data-stu-id="2b1c9-197">First, it loads the values and row count of the used range.</span></span> <span data-ttu-id="2b1c9-198">これにより、値が表示され、いつ停止すればよいかを確認できます。</span><span class="sxs-lookup"><span data-stu-id="2b1c9-198">This lets us look at values and know when to stop.</span></span> <span data-ttu-id="2b1c9-199">次に、指定された範囲を反復処理し、**[借方]** 列と **[貸方]** 列の各セルをチェックします。</span><span class="sxs-lookup"><span data-stu-id="2b1c9-199">Second, it iterates through the used range, checking each cell in the **Debit** or **Credit** columns.</span></span> <span data-ttu-id="2b1c9-200">最後に、セルの値が 0 ではない場合、その値が絶対値で置き換えられます。</span><span class="sxs-lookup"><span data-stu-id="2b1c9-200">Finally, if the value in the cell is not 0, it is replaced by its absolute value.</span></span> <span data-ttu-id="2b1c9-201">0 は使用しないので、空のセルはそのままにしておきます。</span><span class="sxs-lookup"><span data-stu-id="2b1c9-201">We're avoiding zeroes so we can leave the blank cells as they were.</span></span>

3. <span data-ttu-id="2b1c9-202">スクリプトを実行します。</span><span class="sxs-lookup"><span data-stu-id="2b1c9-202">Run the script.</span></span>

    <span data-ttu-id="2b1c9-203">銀行取引明細書は次のように表示されるはずです。</span><span class="sxs-lookup"><span data-stu-id="2b1c9-203">Your banking statement should now look like this:</span></span>

    ![書式設定された正の値のみを含むテーブル形式の銀行取引明細書。](../images/tutorial-5.png)

## <a name="next-steps"></a><span data-ttu-id="2b1c9-205">次の手順</span><span class="sxs-lookup"><span data-stu-id="2b1c9-205">Next steps</span></span>

<span data-ttu-id="2b1c9-206">コード エディターを開き、「[Excel on the web での Office スクリプトのサンプル スクリプト](../resources/excel-samples.md)」をいくつか試してみます。</span><span class="sxs-lookup"><span data-stu-id="2b1c9-206">Open the Code Editor and try out some of our [Sample scripts for Office Scripts in Excel on the web](../resources/excel-samples.md).</span></span> <span data-ttu-id="2b1c9-207">Office スクリプトの作成について詳しくは、「[Excel on the web での Office スクリプトのスクリプトの基本事項](../develop/scripting-fundamentals.md)」も参照してください。</span><span class="sxs-lookup"><span data-stu-id="2b1c9-207">You can also visit [Scripting Fundamentals for Office Scripts in Excel on the web](../develop/scripting-fundamentals.md) to learn more about creating Office Scripts.</span></span>
