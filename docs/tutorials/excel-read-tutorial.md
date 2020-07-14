---
title: Excel on the web で Office スクリプトを使用してブックのデータを読み取る
description: ブックのデータを読み取り、スクリプトでそのデータを評価する方法について説明した Office スクリプトのチュートリアル。
ms.date: 04/23/2020
localization_priority: Priority
ms.openlocfilehash: 93204184d4b5947b2a67107b1fd73c178a73c32e
ms.sourcegitcommit: aec3c971c6640429f89b6bb99d2c95ea06725599
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 06/25/2020
ms.locfileid: "44878688"
---
# <a name="read-workbook-data-with-office-scripts-in-excel-on-the-web"></a><span data-ttu-id="4df3c-103">Excel on the web で Office スクリプトを使用してブックのデータを読み取る</span><span class="sxs-lookup"><span data-stu-id="4df3c-103">Read workbook data with Office Scripts in Excel on the web</span></span>

<span data-ttu-id="4df3c-104">このチュートリアルでは、Excel on the web 用の Office スクリプトを使用してブックのデータを読み取る方法について説明します。</span><span class="sxs-lookup"><span data-stu-id="4df3c-104">This tutorial teaches you how to read data from a workbook with an Office Script for Excel on the web.</span></span> <span data-ttu-id="4df3c-105">その後、読み取ったデータを編集し、ブックに戻します。</span><span class="sxs-lookup"><span data-stu-id="4df3c-105">You'll then edit the data you read and put it back in the workbook.</span></span>

> [!TIP]
> <span data-ttu-id="4df3c-106">Office スクリプトを初めて使用する場合は、チュートリアルの「[Excel on the web で Office スクリプトを記録、編集、作成する](excel-tutorial.md)」から始めることをお勧めします。</span><span class="sxs-lookup"><span data-stu-id="4df3c-106">If you are new to Office Scripts, we recommend starting with the [Record, edit, and create Office Scripts in Excel on the web](excel-tutorial.md) tutorial.</span></span>

## <a name="prerequisites"></a><span data-ttu-id="4df3c-107">前提条件</span><span class="sxs-lookup"><span data-stu-id="4df3c-107">Prerequisites</span></span>

[!INCLUDE [Tutorial prerequisites](../includes/tutorial-prerequisites.md)]

> [!IMPORTANT]
> <span data-ttu-id="4df3c-108">このチュートリアルは、JavaScript や TypeScript について初級から中級レベルの知識を持つユーザーを対象としています。</span><span class="sxs-lookup"><span data-stu-id="4df3c-108">This tutorial is intended for people with beginner to intermediate-level knowledge of JavaScript or TypeScript.</span></span> <span data-ttu-id="4df3c-109">JavaScript を使い慣れていない場合は、[Mozilla の JavaScript チュートリアル](https://developer.mozilla.org/docs/Web/JavaScript/Guide/Introduction)をご覧になることをお勧めします。</span><span class="sxs-lookup"><span data-stu-id="4df3c-109">If you're new to JavaScript, we recommend reviewing the [Mozilla JavaScript tutorial](https://developer.mozilla.org/docs/Web/JavaScript/Guide/Introduction).</span></span> <span data-ttu-id="4df3c-110">スクリプト環境の詳細については、「[Excel on the web の Office スクリプト](../overview/excel.md)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="4df3c-110">Visit [Office Scripts in Excel on the web](../overview/excel.md) to learn more about the script environment.</span></span>

## <a name="read-a-cell"></a><span data-ttu-id="4df3c-111">セルを読み取る。</span><span class="sxs-lookup"><span data-stu-id="4df3c-111">Read a cell</span></span>

<span data-ttu-id="4df3c-112">操作レコーダーで作成したスクリプトは、ブックに情報を書き込む操作のみを実行できます。</span><span class="sxs-lookup"><span data-stu-id="4df3c-112">Scripts made with the Action Recorder can only write information to the workbook.</span></span> <span data-ttu-id="4df3c-113">コード エディターを使用すると、ブックのデータを読み取ることも可能なスクリプトの編集と作成ができます。</span><span class="sxs-lookup"><span data-stu-id="4df3c-113">With the Code Editor, you can edit and make scripts that also read data from a workbook.</span></span>

<span data-ttu-id="4df3c-114">データを読み取り、読み取った内容に基づいて動作するスクリプトを作成しましょう。</span><span class="sxs-lookup"><span data-stu-id="4df3c-114">Let's make a script that reads data and acts based on what was read.</span></span> <span data-ttu-id="4df3c-115">今回は、サンプルの銀行取引明細書を使用します。</span><span class="sxs-lookup"><span data-stu-id="4df3c-115">We're going to work with a sample banking statement.</span></span> <span data-ttu-id="4df3c-116">この明細書は、支払いと貸方がまとまった明細書です。</span><span class="sxs-lookup"><span data-stu-id="4df3c-116">This statement is a combined checking and credit statement.</span></span> <span data-ttu-id="4df3c-117">残念ながら、残高の変化が異なる仕方で報告されています。</span><span class="sxs-lookup"><span data-stu-id="4df3c-117">Unfortunately, they report balance changes differently.</span></span> <span data-ttu-id="4df3c-118">支払い明細では、収入を負の貸方として記録し、支出を負の借方として記録しています。</span><span class="sxs-lookup"><span data-stu-id="4df3c-118">The checking statement gives income as positive credit and costs as negative debit.</span></span> <span data-ttu-id="4df3c-119">貸方明細ではその逆になっています。</span><span class="sxs-lookup"><span data-stu-id="4df3c-119">The credit statement does the opposite.</span></span>

<span data-ttu-id="4df3c-120">チュートリアルの残りの部分で、スクリプトを使用してこのデータを正規化します。</span><span class="sxs-lookup"><span data-stu-id="4df3c-120">Over the rest of the tutorial, we will normalize this data using a script.</span></span> <span data-ttu-id="4df3c-121">まず、ブックからデータを読み取る方法について説明します。</span><span class="sxs-lookup"><span data-stu-id="4df3c-121">First, let's learn how to read data from the workbook.</span></span>

1. <span data-ttu-id="4df3c-122">チュートリアルの残りの部分で使用したブックに新しいワークシートを作成します。</span><span class="sxs-lookup"><span data-stu-id="4df3c-122">Create a new worksheet in the workbook you've used for the rest of the tutorial.</span></span>
2. <span data-ttu-id="4df3c-123">次のデータをコピーし、新しいワークシートのセル **A1** から始まるセル範囲に貼り付けます。</span><span class="sxs-lookup"><span data-stu-id="4df3c-123">Copy the following data and paste it into the new worksheet, starting at cell **A1**.</span></span>

    |<span data-ttu-id="4df3c-124">日付</span><span class="sxs-lookup"><span data-stu-id="4df3c-124">Date</span></span> |<span data-ttu-id="4df3c-125">取引</span><span class="sxs-lookup"><span data-stu-id="4df3c-125">Account</span></span> |<span data-ttu-id="4df3c-126">説明</span><span class="sxs-lookup"><span data-stu-id="4df3c-126">Description</span></span> |<span data-ttu-id="4df3c-127">借方</span><span class="sxs-lookup"><span data-stu-id="4df3c-127">Debit</span></span> |<span data-ttu-id="4df3c-128">貸方</span><span class="sxs-lookup"><span data-stu-id="4df3c-128">Credit</span></span> |
    |:--|:--|:--|:--|:--|
    |<span data-ttu-id="4df3c-129">2019/10/10</span><span class="sxs-lookup"><span data-stu-id="4df3c-129">10/10/2019</span></span> |<span data-ttu-id="4df3c-130">支払い</span><span class="sxs-lookup"><span data-stu-id="4df3c-130">Checking</span></span> |<span data-ttu-id="4df3c-131">Coho Vineyard</span><span class="sxs-lookup"><span data-stu-id="4df3c-131">Coho Vineyard</span></span> |<span data-ttu-id="4df3c-132">-20.05</span><span class="sxs-lookup"><span data-stu-id="4df3c-132">-20.05</span></span> | |
    |<span data-ttu-id="4df3c-133">2019/10/11</span><span class="sxs-lookup"><span data-stu-id="4df3c-133">10/11/2019</span></span> |<span data-ttu-id="4df3c-134">貸方</span><span class="sxs-lookup"><span data-stu-id="4df3c-134">Credit</span></span> |<span data-ttu-id="4df3c-135">The Phone Company</span><span class="sxs-lookup"><span data-stu-id="4df3c-135">The Phone Company</span></span> |<span data-ttu-id="4df3c-136">99.95</span><span class="sxs-lookup"><span data-stu-id="4df3c-136">99.95</span></span> | |
    |<span data-ttu-id="4df3c-137">2019/10/13</span><span class="sxs-lookup"><span data-stu-id="4df3c-137">10/13/2019</span></span> |<span data-ttu-id="4df3c-138">貸方</span><span class="sxs-lookup"><span data-stu-id="4df3c-138">Credit</span></span> |<span data-ttu-id="4df3c-139">Coho Vineyard</span><span class="sxs-lookup"><span data-stu-id="4df3c-139">Coho Vineyard</span></span> |<span data-ttu-id="4df3c-140">154.43</span><span class="sxs-lookup"><span data-stu-id="4df3c-140">154.43</span></span> | |
    |<span data-ttu-id="4df3c-141">2019/10/15</span><span class="sxs-lookup"><span data-stu-id="4df3c-141">10/15/2019</span></span> |<span data-ttu-id="4df3c-142">支払い</span><span class="sxs-lookup"><span data-stu-id="4df3c-142">Checking</span></span> |<span data-ttu-id="4df3c-143">外部預金</span><span class="sxs-lookup"><span data-stu-id="4df3c-143">External Deposit</span></span> | |<span data-ttu-id="4df3c-144">1000</span><span class="sxs-lookup"><span data-stu-id="4df3c-144">1000</span></span> |
    |<span data-ttu-id="4df3c-145">2019/10/20</span><span class="sxs-lookup"><span data-stu-id="4df3c-145">10/20/2019</span></span> |<span data-ttu-id="4df3c-146">貸方</span><span class="sxs-lookup"><span data-stu-id="4df3c-146">Credit</span></span> |<span data-ttu-id="4df3c-147">Coho Vineyard - 返金</span><span class="sxs-lookup"><span data-stu-id="4df3c-147">Coho Vineyard - Refund</span></span> | |<span data-ttu-id="4df3c-148">-35.45</span><span class="sxs-lookup"><span data-stu-id="4df3c-148">-35.45</span></span> |
    |<span data-ttu-id="4df3c-149">2019/10/25</span><span class="sxs-lookup"><span data-stu-id="4df3c-149">10/25/2019</span></span> |<span data-ttu-id="4df3c-150">支払い</span><span class="sxs-lookup"><span data-stu-id="4df3c-150">Checking</span></span> |<span data-ttu-id="4df3c-151">Best For You Organics Company</span><span class="sxs-lookup"><span data-stu-id="4df3c-151">Best For You Organics Company</span></span> | <span data-ttu-id="4df3c-152">-85.64</span><span class="sxs-lookup"><span data-stu-id="4df3c-152">-85.64</span></span> | |
    |<span data-ttu-id="4df3c-153">2019/11/01</span><span class="sxs-lookup"><span data-stu-id="4df3c-153">11/01/2019</span></span> |<span data-ttu-id="4df3c-154">支払い</span><span class="sxs-lookup"><span data-stu-id="4df3c-154">Checking</span></span> |<span data-ttu-id="4df3c-155">外部預金</span><span class="sxs-lookup"><span data-stu-id="4df3c-155">External Deposit</span></span> | |<span data-ttu-id="4df3c-156">1000</span><span class="sxs-lookup"><span data-stu-id="4df3c-156">1000</span></span> |

3. <span data-ttu-id="4df3c-157">**[コード エディター]** を開き、**[新しいスクリプト]** を選択します。</span><span class="sxs-lookup"><span data-stu-id="4df3c-157">Open the **Code Editor** and select **New Script**.</span></span>
4. <span data-ttu-id="4df3c-158">書式設定をクリーンアップします。</span><span class="sxs-lookup"><span data-stu-id="4df3c-158">Let's clean up the formatting.</span></span> <span data-ttu-id="4df3c-159">これは財務ドキュメントなので、**[借方]** 列と **[貸方]** 列の数値の書式設定を変更して、値がドル金額として表示されるようにします。</span><span class="sxs-lookup"><span data-stu-id="4df3c-159">This is a financial document, so let's change the number formatting in the **Debit** and **Credit** columns to show values as dollar amounts.</span></span> <span data-ttu-id="4df3c-160">さらに、列幅をデータに合わせます。</span><span class="sxs-lookup"><span data-stu-id="4df3c-160">Let's also fit the column width to the data.</span></span>

    <span data-ttu-id="4df3c-161">スクリプトの内容を次のコードで置き換えます。</span><span class="sxs-lookup"><span data-stu-id="4df3c-161">Replace the script contents with the following code:</span></span>

    ```TypeScript
    function main(workbook: ExcelScript.Workbook) {
        // Get the current worksheet.
        let selectedSheet = workbook.getActiveWorksheet();

        // Format the range to display numerical dollar amounts.
        selectedSheet.getRange("D2:E8").setNumberFormat("$#,##0.00");

        // Fit the width of all the used columns to the data.
        selectedSheet.getUsedRange().getFormat().autofitColumns();
    }
    ```

5. <span data-ttu-id="4df3c-162">では、いずれかの数値列の値を読み取ってみましょう。</span><span class="sxs-lookup"><span data-stu-id="4df3c-162">Now let's read a value from one of the number columns.</span></span> <span data-ttu-id="4df3c-163">次のコードをスクリプトの最後 (末尾の `}` の前) に追加します。</span><span class="sxs-lookup"><span data-stu-id="4df3c-163">Add the following code to the end of the script (before the closing `}`):</span></span>

    ```TypeScript
    // Get the value of cell D2.
    let range = selectedSheet.getRange("D2");
    console.log(range.getValues());
    ```

6. <span data-ttu-id="4df3c-164">スクリプトを実行します。</span><span class="sxs-lookup"><span data-stu-id="4df3c-164">Run the script.</span></span>
7. <span data-ttu-id="4df3c-165">コンソールを開きます。</span><span class="sxs-lookup"><span data-stu-id="4df3c-165">Open the console.</span></span> <span data-ttu-id="4df3c-166">**省略記号**のメニューを選択し、**[Logs...](ログ...)** を押します。</span><span class="sxs-lookup"><span data-stu-id="4df3c-166">Go to the **Ellipses** menu and press **Logs...**.</span></span>
8. <span data-ttu-id="4df3c-167">コンソールに `[Array[1]]` が表示されます。</span><span class="sxs-lookup"><span data-stu-id="4df3c-167">You should see `[Array[1]]` in the console.</span></span> <span data-ttu-id="4df3c-168">範囲は 2 次元のデータ配列であるため、これは数値ではありません。</span><span class="sxs-lookup"><span data-stu-id="4df3c-168">This is not a number because ranges are two-dimensional arrays of data.</span></span> <span data-ttu-id="4df3c-169">この 2 次元の範囲は、コンソールに直接ログ記録されます。</span><span class="sxs-lookup"><span data-stu-id="4df3c-169">That two-dimensional range is being logged to the console directly.</span></span> <span data-ttu-id="4df3c-170">コード エディターを使用すると、この配列の内容を表示できます。</span><span class="sxs-lookup"><span data-stu-id="4df3c-170">Luckily, the Code Editor does let you see the contents of the array.</span></span>
9. <span data-ttu-id="4df3c-171">2 次元の配列がコンソールにログ記録すると、各行の下に列の値がグループ化されます。</span><span class="sxs-lookup"><span data-stu-id="4df3c-171">When a two-dimensional array is logged to the console, it groups column values under each row.</span></span> <span data-ttu-id="4df3c-172">青い三角形を押して、配列のログを展開します。</span><span class="sxs-lookup"><span data-stu-id="4df3c-172">Expand the array log by pressing the blue triangle.</span></span>
10. <span data-ttu-id="4df3c-173">新たに表示された青い三角形を押して、配列の第 2 レベルを展開します。</span><span class="sxs-lookup"><span data-stu-id="4df3c-173">Expand the second level of the array by pressing the newly revealed blue triangle.</span></span> <span data-ttu-id="4df3c-174">次のように表示されるはずです。</span><span class="sxs-lookup"><span data-stu-id="4df3c-174">You should now see this:</span></span>

    ![出力 "-20.05" が 2 つの配列の下に入れ子になって表示されているコンソール ログ。](../images/tutorial-4.png)

## <a name="modify-the-value-of-a-cell"></a><span data-ttu-id="4df3c-176">セルの値を変更する</span><span class="sxs-lookup"><span data-stu-id="4df3c-176">Modify the value of a cell</span></span>

<span data-ttu-id="4df3c-177">データを読み取れたので、そのデータを使用してブックを変更しましょう。</span><span class="sxs-lookup"><span data-stu-id="4df3c-177">Now that we can read data, let's use that data to modify the workbook.</span></span> <span data-ttu-id="4df3c-178">セル **D2** の値を、`Math.abs` 関数を使用して正の値にします。</span><span class="sxs-lookup"><span data-stu-id="4df3c-178">We'll make the value of the cell **D2** positive with the `Math.abs` function.</span></span> <span data-ttu-id="4df3c-179">[Math](https://developer.mozilla.org/docs/web/javascript/reference/global_objects/math) オブジェクトには、スクリプトでアクセスできる多くの関数が含まれています。</span><span class="sxs-lookup"><span data-stu-id="4df3c-179">The [Math](https://developer.mozilla.org/docs/web/javascript/reference/global_objects/math) object contains many functions to which your scripts have access.</span></span> <span data-ttu-id="4df3c-180">`Math` および他の組み込みオブジェクトの詳細については、「[Office スクリプトでの組み込みの JavaScript オブジェクトの使用](../develop/javascript-objects.md)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="4df3c-180">More information about `Math` and other built-in objects can be found at [Using built-in JavaScript objects in Office Scripts](../develop/javascript-objects.md).</span></span>

1. <span data-ttu-id="4df3c-181">次のコードをスクリプトの最後に追加します。</span><span class="sxs-lookup"><span data-stu-id="4df3c-181">Add the following code to the end of the script:</span></span>

    ```TypeScript
        // Run the `Math.abs` function with the value at D2 and apply that value back to D2.
    let positiveValue = Math.abs(range.getValue());
    range.setValue(positiveValue);
    ```

    <span data-ttu-id="4df3c-182">`getValue` と `setValue` を使用していることに注意してください。</span><span class="sxs-lookup"><span data-stu-id="4df3c-182">Note that we're using `getValue` and `setValue`.</span></span> <span data-ttu-id="4df3c-183">これらの方法は、1 つのセルで使用できます。</span><span class="sxs-lookup"><span data-stu-id="4df3c-183">These methods work on a single cell.</span></span> <span data-ttu-id="4df3c-184">複数のセル範囲を処理する場合は、`getValues` と `setValues` を使用します。</span><span class="sxs-lookup"><span data-stu-id="4df3c-184">When handling multi-cell ranges, you'll want to use `getValues` and `setValues`.</span></span>

2. <span data-ttu-id="4df3c-185">セル **D2** の値が正の値になります。</span><span class="sxs-lookup"><span data-stu-id="4df3c-185">The value of cell **D2** should now be positive.</span></span>

## <a name="modify-the-values-of-a-column"></a><span data-ttu-id="4df3c-186">列の値を変更する</span><span class="sxs-lookup"><span data-stu-id="4df3c-186">Modify the values of a column</span></span>

<span data-ttu-id="4df3c-187">1 つのセルの読み取り方法と書き込み方法がわかったので、スクリプトを一般化して、**[借方]** 列と **[貸方]** 列全体を操作できるようにしましょう。</span><span class="sxs-lookup"><span data-stu-id="4df3c-187">Now that we know how to read and write to a single cell, let's generalize the script to work on the entire **Debit** and **Credit** columns.</span></span>

1. <span data-ttu-id="4df3c-188">1 つのセルにのみ影響するコード (前述の絶対値コード) を削除します。すると、スクリプトは次のようになります。</span><span class="sxs-lookup"><span data-stu-id="4df3c-188">Remove the code that affects only a single cell (the previous absolute value code), such that your script now looks like this:</span></span>

    ```TypeScript
    function main(workbook: ExcelScript.Workbook) {
        // Get the current worksheet.
        let selectedSheet = workbook.getActiveWorksheet();

        // Format the range to display numerical dollar amounts.
        selectedSheet.getRange("D2:E8").setNumberFormat("$#,##0.00");

        // Fit the width of all the used columns to the data.
        selectedSheet.getUsedRange().getFormat().autofitColumns();
    }
    ```

2. <span data-ttu-id="4df3c-189">最後の 2 つの列の行を反復処理するループをスクリプトの最後に追加します。</span><span class="sxs-lookup"><span data-stu-id="4df3c-189">Add a loop to the end of the script that iterates through the rows in the last two columns.</span></span> <span data-ttu-id="4df3c-190">スクリプトにより、各セルの値が現在の値の絶対値に設定されます。</span><span class="sxs-lookup"><span data-stu-id="4df3c-190">For each cell, the script sets the value to the current value's absolute value.</span></span>

    <span data-ttu-id="4df3c-191">セルの位置を定義する配列は 0 から始まることにご注意ください。</span><span class="sxs-lookup"><span data-stu-id="4df3c-191">Note that the array defining cell locations is zero-based.</span></span> <span data-ttu-id="4df3c-192">したがって、セル **A1** は `range[0][0]` になります。</span><span class="sxs-lookup"><span data-stu-id="4df3c-192">That means cell **A1** is `range[0][0]`.</span></span>

    ```TypeScript
    // Get the values of the used range.
    let range = selectedSheet.getUsedRange();
    let rangeValues = range.getValues();

    // Iterate over the fourth and fifth columns and set their values to their absolute value.
    for (let i = 1; i < range.getRowCount(); i++) {
        // The column at index 3 is column "4" in the worksheet.
        if (rangeValues[i][3] != 0) {
            let positiveValue = Math.abs(rangeValues[i][3]);
            selectedSheet.getCell(i, 3).setValue(positiveValue);
        }

        // The column at index 4 is column "5" in the worksheet.
        if (rangeValues[i][4] != 0) {
            let positiveValue = Math.abs(rangeValues[i][4]);
            selectedSheet.getCell(i, 4).setValue(positiveValue);
        }
    }
    ```

    <span data-ttu-id="4df3c-193">スクリプトのこの部分は、いくつかの重要なタスクを実行します。</span><span class="sxs-lookup"><span data-stu-id="4df3c-193">This portion of the script does several important tasks.</span></span> <span data-ttu-id="4df3c-194">まず、指定された範囲の値と行数を取得します。</span><span class="sxs-lookup"><span data-stu-id="4df3c-194">First, it gets the values and row count of the used range.</span></span> <span data-ttu-id="4df3c-195">これにより、値が表示され、いつ停止すればよいかを確認できます。</span><span class="sxs-lookup"><span data-stu-id="4df3c-195">This lets us look at values and know when to stop.</span></span> <span data-ttu-id="4df3c-196">次に、指定された範囲を反復処理し、**[借方]** 列と **[貸方]** 列の各セルをチェックします。</span><span class="sxs-lookup"><span data-stu-id="4df3c-196">Second, it iterates through the used range, checking each cell in the **Debit** or **Credit** columns.</span></span> <span data-ttu-id="4df3c-197">最後に、セルの値が 0 ではない場合、その値が絶対値で置き換えられます。</span><span class="sxs-lookup"><span data-stu-id="4df3c-197">Finally, if the value in the cell is not 0, it is replaced by its absolute value.</span></span> <span data-ttu-id="4df3c-198">0 は使用しないので、空のセルはそのままにしておきます。</span><span class="sxs-lookup"><span data-stu-id="4df3c-198">We're avoiding zeroes so we can leave the blank cells as they were.</span></span>

3. <span data-ttu-id="4df3c-199">スクリプトを実行します。</span><span class="sxs-lookup"><span data-stu-id="4df3c-199">Run the script.</span></span>

    <span data-ttu-id="4df3c-200">銀行取引明細書は次のように表示されるはずです。</span><span class="sxs-lookup"><span data-stu-id="4df3c-200">Your banking statement should now look like this:</span></span>

    ![書式設定された正の値のみを含むテーブル形式の銀行取引明細書。](../images/tutorial-5.png)

## <a name="next-steps"></a><span data-ttu-id="4df3c-202">次の手順</span><span class="sxs-lookup"><span data-stu-id="4df3c-202">Next steps</span></span>

<span data-ttu-id="4df3c-203">コード エディターを開き、「[Excel on the web での Office スクリプトのサンプル スクリプト](../resources/excel-samples.md)」をいくつか試してみます。</span><span class="sxs-lookup"><span data-stu-id="4df3c-203">Open the Code Editor and try out some of our [Sample scripts for Office Scripts in Excel on the web](../resources/excel-samples.md).</span></span> <span data-ttu-id="4df3c-204">Office スクリプトの作成について詳しくは、「[Excel on the web での Office スクリプトのスクリプトの基本事項](../develop/scripting-fundamentals.md)」も参照してください。</span><span class="sxs-lookup"><span data-stu-id="4df3c-204">You can also visit [Scripting Fundamentals for Office Scripts in Excel on the web](../develop/scripting-fundamentals.md) to learn more about creating Office Scripts.</span></span>
