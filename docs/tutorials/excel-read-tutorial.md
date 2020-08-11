---
title: Excel on the web で Office スクリプトを使用してブックのデータを読み取る
description: ブックのデータを読み取り、スクリプトでそのデータを評価する方法について説明した Office スクリプトのチュートリアル。
ms.date: 07/20/2020
localization_priority: Priority
ms.openlocfilehash: cdd09f13bb53cfff8c051360f2306cdb6956d86d
ms.sourcegitcommit: ff7fde04ce5a66d8df06ed505951c8111e2e9833
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 08/11/2020
ms.locfileid: "46616709"
---
# <a name="read-workbook-data-with-office-scripts-in-excel-on-the-web"></a><span data-ttu-id="ca069-103">Excel on the web で Office スクリプトを使用してブックのデータを読み取る</span><span class="sxs-lookup"><span data-stu-id="ca069-103">Read workbook data with Office Scripts in Excel on the web</span></span>

<span data-ttu-id="ca069-104">このチュートリアルでは、Excel on the web 用の Office スクリプトを使用してブックのデータを読み取る方法について説明します。</span><span class="sxs-lookup"><span data-stu-id="ca069-104">This tutorial teaches you how to read data from a workbook with an Office Script for Excel on the web.</span></span> <span data-ttu-id="ca069-105">口座取引明細書の書式設定を行う新しいスクリプトを作成し、明細書のデータを標準化します。</span><span class="sxs-lookup"><span data-stu-id="ca069-105">You'll be writing a new script that formats a bank statement and normalizes the data in that statement.</span></span> <span data-ttu-id="ca069-106">データのクリーンアップの一環として、スクリプトは取引セルの値を読み取り、それぞれの値に簡単な数式を適用し、導き出された回答をブックに書き込みます。</span><span class="sxs-lookup"><span data-stu-id="ca069-106">As part of that data clean-up, your script will read values from the transaction cells, apply a simple formula to each value, and write the resulting answer to the workbook.</span></span> <span data-ttu-id="ca069-107">ブックからデータを読み取ることで、スクリプト内の意思決定プロセスの一部を自動化することができます。</span><span class="sxs-lookup"><span data-stu-id="ca069-107">Reading data from the workbook lets you automate some of your decision making processes in the script.</span></span>

> [!TIP]
> <span data-ttu-id="ca069-108">Office スクリプトを初めて使用する場合は、チュートリアルの「[Excel on the web で Office スクリプトを記録、編集、作成する](excel-tutorial.md)」から始めることをお勧めします。</span><span class="sxs-lookup"><span data-stu-id="ca069-108">If you are new to Office Scripts, we recommend starting with the [Record, edit, and create Office Scripts in Excel on the web](excel-tutorial.md) tutorial.</span></span> <span data-ttu-id="ca069-109">[Office スクリプトは TypeScript を使用](../overview/code-editor-environment.md)します。このチュートリアルは、JavaScript や TypeScript について初級から中級レベルの知識を持つユーザーを対象としています。</span><span class="sxs-lookup"><span data-stu-id="ca069-109">[Office Scripts use TypeScript](../overview/code-editor-environment.md) and this tutorial is intended for people with beginner to intermediate-level knowledge of JavaScript or TypeScript.</span></span> <span data-ttu-id="ca069-110">JavaScript を使い慣れていない場合は、[Mozilla の JavaScript チュートリアル](https://developer.mozilla.org/docs/Web/JavaScript/Guide/Introduction)から始めることをお勧めします。</span><span class="sxs-lookup"><span data-stu-id="ca069-110">If you're new to JavaScript, we recommend starting with the [Mozilla JavaScript tutorial](https://developer.mozilla.org/docs/Web/JavaScript/Guide/Introduction).</span></span>

## <a name="prerequisites"></a><span data-ttu-id="ca069-111">前提条件</span><span class="sxs-lookup"><span data-stu-id="ca069-111">Prerequisites</span></span>

[!INCLUDE [Tutorial prerequisites](../includes/tutorial-prerequisites.md)]

## <a name="read-a-cell"></a><span data-ttu-id="ca069-112">セルを読み取る。</span><span class="sxs-lookup"><span data-stu-id="ca069-112">Read a cell</span></span>

<span data-ttu-id="ca069-113">操作レコーダーで作成したスクリプトは、ブックに情報を書き込む操作のみを実行できます。</span><span class="sxs-lookup"><span data-stu-id="ca069-113">Scripts made with the Action Recorder can only write information to the workbook.</span></span> <span data-ttu-id="ca069-114">コード エディターを使用すると、ブックのデータを読み取ることも可能なスクリプトの編集と作成ができます。</span><span class="sxs-lookup"><span data-stu-id="ca069-114">With the Code Editor, you can edit and make scripts that also read data from a workbook.</span></span>

<span data-ttu-id="ca069-115">データを読み取り、読み取った内容に基づいて動作するスクリプトを作成しましょう。</span><span class="sxs-lookup"><span data-stu-id="ca069-115">Let's make a script that reads data and acts based on what was read.</span></span> <span data-ttu-id="ca069-116">今回は、サンプルの銀行取引明細書を使用します。</span><span class="sxs-lookup"><span data-stu-id="ca069-116">We're going to work with a sample banking statement.</span></span> <span data-ttu-id="ca069-117">この明細書は、支払いと貸方がまとまった明細書です。</span><span class="sxs-lookup"><span data-stu-id="ca069-117">This statement is a combined checking and credit statement.</span></span> <span data-ttu-id="ca069-118">残念ながら、残高の変化が異なる仕方で報告されています。</span><span class="sxs-lookup"><span data-stu-id="ca069-118">Unfortunately, they report balance changes differently.</span></span> <span data-ttu-id="ca069-119">支払い明細では、収入を負の貸方として記録し、支出を負の借方として記録しています。</span><span class="sxs-lookup"><span data-stu-id="ca069-119">The checking statement gives income as positive credit and costs as negative debit.</span></span> <span data-ttu-id="ca069-120">貸方明細ではその逆になっています。</span><span class="sxs-lookup"><span data-stu-id="ca069-120">The credit statement does the opposite.</span></span>

<span data-ttu-id="ca069-121">チュートリアルの残りの部分で、スクリプトを使用してこのデータを正規化します。</span><span class="sxs-lookup"><span data-stu-id="ca069-121">Over the rest of the tutorial, we will normalize this data using a script.</span></span> <span data-ttu-id="ca069-122">まず、ブックからデータを読み取る方法について説明します。</span><span class="sxs-lookup"><span data-stu-id="ca069-122">First, let's learn how to read data from the workbook.</span></span>

1. <span data-ttu-id="ca069-123">チュートリアルの残りの部分で使用したブックに新しいワークシートを作成します。</span><span class="sxs-lookup"><span data-stu-id="ca069-123">Create a new worksheet in the workbook you've used for the rest of the tutorial.</span></span>
2. <span data-ttu-id="ca069-124">次のデータをコピーし、新しいワークシートのセル **A1** から始まるセル範囲に貼り付けます。</span><span class="sxs-lookup"><span data-stu-id="ca069-124">Copy the following data and paste it into the new worksheet, starting at cell **A1**.</span></span>

    |<span data-ttu-id="ca069-125">日付</span><span class="sxs-lookup"><span data-stu-id="ca069-125">Date</span></span> |<span data-ttu-id="ca069-126">取引</span><span class="sxs-lookup"><span data-stu-id="ca069-126">Account</span></span> |<span data-ttu-id="ca069-127">説明</span><span class="sxs-lookup"><span data-stu-id="ca069-127">Description</span></span> |<span data-ttu-id="ca069-128">借方</span><span class="sxs-lookup"><span data-stu-id="ca069-128">Debit</span></span> |<span data-ttu-id="ca069-129">貸方</span><span class="sxs-lookup"><span data-stu-id="ca069-129">Credit</span></span> |
    |:--|:--|:--|:--|:--|
    |<span data-ttu-id="ca069-130">2019/10/10</span><span class="sxs-lookup"><span data-stu-id="ca069-130">10/10/2019</span></span> |<span data-ttu-id="ca069-131">支払い</span><span class="sxs-lookup"><span data-stu-id="ca069-131">Checking</span></span> |<span data-ttu-id="ca069-132">Coho Vineyard</span><span class="sxs-lookup"><span data-stu-id="ca069-132">Coho Vineyard</span></span> |<span data-ttu-id="ca069-133">-20.05</span><span class="sxs-lookup"><span data-stu-id="ca069-133">-20.05</span></span> | |
    |<span data-ttu-id="ca069-134">2019/10/11</span><span class="sxs-lookup"><span data-stu-id="ca069-134">10/11/2019</span></span> |<span data-ttu-id="ca069-135">貸方</span><span class="sxs-lookup"><span data-stu-id="ca069-135">Credit</span></span> |<span data-ttu-id="ca069-136">The Phone Company</span><span class="sxs-lookup"><span data-stu-id="ca069-136">The Phone Company</span></span> |<span data-ttu-id="ca069-137">99.95</span><span class="sxs-lookup"><span data-stu-id="ca069-137">99.95</span></span> | |
    |<span data-ttu-id="ca069-138">2019/10/13</span><span class="sxs-lookup"><span data-stu-id="ca069-138">10/13/2019</span></span> |<span data-ttu-id="ca069-139">貸方</span><span class="sxs-lookup"><span data-stu-id="ca069-139">Credit</span></span> |<span data-ttu-id="ca069-140">Coho Vineyard</span><span class="sxs-lookup"><span data-stu-id="ca069-140">Coho Vineyard</span></span> |<span data-ttu-id="ca069-141">154.43</span><span class="sxs-lookup"><span data-stu-id="ca069-141">154.43</span></span> | |
    |<span data-ttu-id="ca069-142">2019/10/15</span><span class="sxs-lookup"><span data-stu-id="ca069-142">10/15/2019</span></span> |<span data-ttu-id="ca069-143">支払い</span><span class="sxs-lookup"><span data-stu-id="ca069-143">Checking</span></span> |<span data-ttu-id="ca069-144">外部預金</span><span class="sxs-lookup"><span data-stu-id="ca069-144">External Deposit</span></span> | |<span data-ttu-id="ca069-145">1000</span><span class="sxs-lookup"><span data-stu-id="ca069-145">1000</span></span> |
    |<span data-ttu-id="ca069-146">2019/10/20</span><span class="sxs-lookup"><span data-stu-id="ca069-146">10/20/2019</span></span> |<span data-ttu-id="ca069-147">貸方</span><span class="sxs-lookup"><span data-stu-id="ca069-147">Credit</span></span> |<span data-ttu-id="ca069-148">Coho Vineyard - 返金</span><span class="sxs-lookup"><span data-stu-id="ca069-148">Coho Vineyard - Refund</span></span> | |<span data-ttu-id="ca069-149">-35.45</span><span class="sxs-lookup"><span data-stu-id="ca069-149">-35.45</span></span> |
    |<span data-ttu-id="ca069-150">2019/10/25</span><span class="sxs-lookup"><span data-stu-id="ca069-150">10/25/2019</span></span> |<span data-ttu-id="ca069-151">支払い</span><span class="sxs-lookup"><span data-stu-id="ca069-151">Checking</span></span> |<span data-ttu-id="ca069-152">Best For You Organics Company</span><span class="sxs-lookup"><span data-stu-id="ca069-152">Best For You Organics Company</span></span> | <span data-ttu-id="ca069-153">-85.64</span><span class="sxs-lookup"><span data-stu-id="ca069-153">-85.64</span></span> | |
    |<span data-ttu-id="ca069-154">2019/11/01</span><span class="sxs-lookup"><span data-stu-id="ca069-154">11/01/2019</span></span> |<span data-ttu-id="ca069-155">支払い</span><span class="sxs-lookup"><span data-stu-id="ca069-155">Checking</span></span> |<span data-ttu-id="ca069-156">外部預金</span><span class="sxs-lookup"><span data-stu-id="ca069-156">External Deposit</span></span> | |<span data-ttu-id="ca069-157">1000</span><span class="sxs-lookup"><span data-stu-id="ca069-157">1000</span></span> |

3. <span data-ttu-id="ca069-158">**[コード エディター]** を開き、**[新しいスクリプト]** を選択します。</span><span class="sxs-lookup"><span data-stu-id="ca069-158">Open the **Code Editor** and select **New Script**.</span></span>
4. <span data-ttu-id="ca069-159">書式設定をクリーンアップします。</span><span class="sxs-lookup"><span data-stu-id="ca069-159">Let's clean up the formatting.</span></span> <span data-ttu-id="ca069-160">これは財務ドキュメントなので、**[借方]** 列と **[貸方]** 列の数値の書式設定を変更して、値がドル金額として表示されるようにします。</span><span class="sxs-lookup"><span data-stu-id="ca069-160">This is a financial document, so let's change the number formatting in the **Debit** and **Credit** columns to show values as dollar amounts.</span></span> <span data-ttu-id="ca069-161">さらに、列幅をデータに合わせます。</span><span class="sxs-lookup"><span data-stu-id="ca069-161">Let's also fit the column width to the data.</span></span>

    <span data-ttu-id="ca069-162">スクリプトの内容を次のコードで置き換えます。</span><span class="sxs-lookup"><span data-stu-id="ca069-162">Replace the script contents with the following code:</span></span>

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

5. <span data-ttu-id="ca069-163">では、いずれかの数値列の値を読み取ってみましょう。</span><span class="sxs-lookup"><span data-stu-id="ca069-163">Now let's read a value from one of the number columns.</span></span> <span data-ttu-id="ca069-164">次のコードをスクリプトの最後 (末尾の `}` の前) に追加します。</span><span class="sxs-lookup"><span data-stu-id="ca069-164">Add the following code to the end of the script (before the closing `}`):</span></span>

    ```TypeScript
    // Get the value of cell D2.
    let range = selectedSheet.getRange("D2");
    console.log(range.getValues());
    ```

6. <span data-ttu-id="ca069-165">スクリプトを実行します。</span><span class="sxs-lookup"><span data-stu-id="ca069-165">Run the script.</span></span>
7. <span data-ttu-id="ca069-166">コンソールに `[Array[1]]` が表示されます。</span><span class="sxs-lookup"><span data-stu-id="ca069-166">You should see `[Array[1]]` in the console.</span></span> <span data-ttu-id="ca069-167">範囲は 2 次元のデータ配列であるため、これは数値ではありません。</span><span class="sxs-lookup"><span data-stu-id="ca069-167">This is not a number because ranges are two-dimensional arrays of data.</span></span> <span data-ttu-id="ca069-168">この 2 次元の範囲は、コンソールに直接ログ記録されます。</span><span class="sxs-lookup"><span data-stu-id="ca069-168">That two-dimensional range is being logged to the console directly.</span></span> <span data-ttu-id="ca069-169">コード エディターを使用すると、この配列の内容を表示できます。</span><span class="sxs-lookup"><span data-stu-id="ca069-169">Luckily, the Code Editor lets you see the contents of the array.</span></span>
8. <span data-ttu-id="ca069-170">2 次元の配列がコンソールにログ記録すると、各行の下に列の値がグループ化されます。</span><span class="sxs-lookup"><span data-stu-id="ca069-170">When a two-dimensional array is logged to the console, it groups column values under each row.</span></span> <span data-ttu-id="ca069-171">青い三角形を押して、配列のログを展開します。</span><span class="sxs-lookup"><span data-stu-id="ca069-171">Expand the array log by pressing the blue triangle.</span></span>
9. <span data-ttu-id="ca069-172">新たに表示された青い三角形を押して、配列の第 2 レベルを展開します。</span><span class="sxs-lookup"><span data-stu-id="ca069-172">Expand the second level of the array by pressing the newly revealed blue triangle.</span></span> <span data-ttu-id="ca069-173">次のように表示されるはずです。</span><span class="sxs-lookup"><span data-stu-id="ca069-173">You should now see this:</span></span>

    ![出力 "-20.05" が 2 つの配列の下に入れ子になって表示されているコンソール ログ。](../images/tutorial-4.png)

## <a name="modify-the-value-of-a-cell"></a><span data-ttu-id="ca069-175">セルの値を変更する</span><span class="sxs-lookup"><span data-stu-id="ca069-175">Modify the value of a cell</span></span>

<span data-ttu-id="ca069-176">データを読み取れたので、そのデータを使用してブックを変更しましょう。</span><span class="sxs-lookup"><span data-stu-id="ca069-176">Now that we can read data, let's use that data to modify the workbook.</span></span> <span data-ttu-id="ca069-177">セル **D2** の値を、`Math.abs` 関数を使用して正の値にします。</span><span class="sxs-lookup"><span data-stu-id="ca069-177">We'll make the value of the cell **D2** positive with the `Math.abs` function.</span></span> <span data-ttu-id="ca069-178">[Math](https://developer.mozilla.org/docs/web/javascript/reference/global_objects/math) オブジェクトには、スクリプトでアクセスできる多くの関数が含まれています。</span><span class="sxs-lookup"><span data-stu-id="ca069-178">The [Math](https://developer.mozilla.org/docs/web/javascript/reference/global_objects/math) object contains many functions to which your scripts have access.</span></span> <span data-ttu-id="ca069-179">`Math` および他の組み込みオブジェクトの詳細については、「[Office スクリプトでの組み込みの JavaScript オブジェクトの使用](../develop/javascript-objects.md)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="ca069-179">More information about `Math` and other built-in objects can be found at [Using built-in JavaScript objects in Office Scripts](../develop/javascript-objects.md).</span></span>

1. <span data-ttu-id="ca069-180">次のコードをスクリプトの最後に追加します。</span><span class="sxs-lookup"><span data-stu-id="ca069-180">Add the following code to the end of the script:</span></span>

    ```TypeScript
    // Run the `Math.abs` function with the value at D2 and apply that value back to D2.
    let positiveValue = Math.abs(range.getValue());
    range.setValue(positiveValue);
    ```

    <span data-ttu-id="ca069-181">`getValue` と `setValue` を使用していることに注意してください。</span><span class="sxs-lookup"><span data-stu-id="ca069-181">Note that we're using `getValue` and `setValue`.</span></span> <span data-ttu-id="ca069-182">これらの方法は、1 つのセルで使用できます。</span><span class="sxs-lookup"><span data-stu-id="ca069-182">These methods work on a single cell.</span></span> <span data-ttu-id="ca069-183">複数のセル範囲を処理する場合は、`getValues` と `setValues` を使用します。</span><span class="sxs-lookup"><span data-stu-id="ca069-183">When handling multi-cell ranges, you'll want to use `getValues` and `setValues`.</span></span>

2. <span data-ttu-id="ca069-184">セル **D2** の値が正の値になります。</span><span class="sxs-lookup"><span data-stu-id="ca069-184">The value of cell **D2** should now be positive.</span></span>

## <a name="modify-the-values-of-a-column"></a><span data-ttu-id="ca069-185">列の値を変更する</span><span class="sxs-lookup"><span data-stu-id="ca069-185">Modify the values of a column</span></span>

<span data-ttu-id="ca069-186">1 つのセルの読み取り方法と書き込み方法がわかったので、スクリプトを一般化して、**[借方]** 列と **[貸方]** 列全体を操作できるようにしましょう。</span><span class="sxs-lookup"><span data-stu-id="ca069-186">Now that we know how to read and write to a single cell, let's generalize the script to work on the entire **Debit** and **Credit** columns.</span></span>

1. <span data-ttu-id="ca069-187">1 つのセルにのみ影響するコード (前述の絶対値コード) を削除します。すると、スクリプトは次のようになります。</span><span class="sxs-lookup"><span data-stu-id="ca069-187">Remove the code that affects only a single cell (the previous absolute value code), such that your script now looks like this:</span></span>

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

2. <span data-ttu-id="ca069-188">最後の 2 つの列の行を反復処理するループをスクリプトの最後に追加します。</span><span class="sxs-lookup"><span data-stu-id="ca069-188">Add a loop to the end of the script that iterates through the rows in the last two columns.</span></span> <span data-ttu-id="ca069-189">スクリプトにより、各セルの値が現在の値の絶対値に設定されます。</span><span class="sxs-lookup"><span data-stu-id="ca069-189">For each cell, the script sets the value to the current value's absolute value.</span></span>

    <span data-ttu-id="ca069-190">セルの位置を定義する配列は 0 から始まることにご注意ください。</span><span class="sxs-lookup"><span data-stu-id="ca069-190">Note that the array defining cell locations is zero-based.</span></span> <span data-ttu-id="ca069-191">したがって、セル **A1** は `range[0][0]` になります。</span><span class="sxs-lookup"><span data-stu-id="ca069-191">That means cell **A1** is `range[0][0]`.</span></span>

    ```TypeScript
    // Get the values of the used range.
    let range = selectedSheet.getUsedRange();
    let rangeValues = range.getValues();

    // Iterate over the fourth and fifth columns and set their values to their absolute value.
    let rowCount = range.getRowCount();
    for (let i = 1; i < rowCount; i++) {
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

    <span data-ttu-id="ca069-192">スクリプトのこの部分は、いくつかの重要なタスクを実行します。</span><span class="sxs-lookup"><span data-stu-id="ca069-192">This portion of the script does several important tasks.</span></span> <span data-ttu-id="ca069-193">まず、指定された範囲の値と行数を取得します。</span><span class="sxs-lookup"><span data-stu-id="ca069-193">First, it gets the values and row count of the used range.</span></span> <span data-ttu-id="ca069-194">これにより、値が表示され、いつ停止すればよいかを確認できます。</span><span class="sxs-lookup"><span data-stu-id="ca069-194">This lets us look at values and know when to stop.</span></span> <span data-ttu-id="ca069-195">次に、指定された範囲を反復処理し、**[借方]** 列と **[貸方]** 列の各セルをチェックします。</span><span class="sxs-lookup"><span data-stu-id="ca069-195">Second, it iterates through the used range, checking each cell in the **Debit** or **Credit** columns.</span></span> <span data-ttu-id="ca069-196">最後に、セルの値が 0 ではない場合、その値が絶対値で置き換えられます。</span><span class="sxs-lookup"><span data-stu-id="ca069-196">Finally, if the value in the cell is not 0, it is replaced by its absolute value.</span></span> <span data-ttu-id="ca069-197">0 は使用しないので、空のセルはそのままにしておきます。</span><span class="sxs-lookup"><span data-stu-id="ca069-197">We're avoiding zeroes so we can leave the blank cells as they were.</span></span>

3. <span data-ttu-id="ca069-198">スクリプトを実行します。</span><span class="sxs-lookup"><span data-stu-id="ca069-198">Run the script.</span></span>

    <span data-ttu-id="ca069-199">銀行取引明細書は次のように表示されるはずです。</span><span class="sxs-lookup"><span data-stu-id="ca069-199">Your banking statement should now look like this:</span></span>

    ![書式設定された正の値のみを含むテーブル形式の銀行取引明細書。](../images/tutorial-5.png)

## <a name="next-steps"></a><span data-ttu-id="ca069-201">次の手順</span><span class="sxs-lookup"><span data-stu-id="ca069-201">Next steps</span></span>

<span data-ttu-id="ca069-202">コード エディターを開き、「[Excel on the web での Office スクリプトのサンプル スクリプト](../resources/excel-samples.md)」をいくつか試してみます。</span><span class="sxs-lookup"><span data-stu-id="ca069-202">Open the Code Editor and try out some of our [Sample scripts for Office Scripts in Excel on the web](../resources/excel-samples.md).</span></span> <span data-ttu-id="ca069-203">Office スクリプトの作成について詳しくは、「[Excel on the web での Office スクリプトのスクリプトの基本事項](../develop/scripting-fundamentals.md)」も参照してください。</span><span class="sxs-lookup"><span data-stu-id="ca069-203">You can also visit [Scripting Fundamentals for Office Scripts in Excel on the web](../develop/scripting-fundamentals.md) to learn more about creating Office Scripts.</span></span>

<span data-ttu-id="ca069-204">次の一連の Office スクリプトのチュートリアルでは、Power Automate を使用した Office スクリプトの使用法について説明します。</span><span class="sxs-lookup"><span data-stu-id="ca069-204">The next series of Office Scripts tutorials focus on using Office Scripts with Power Automate.</span></span> <span data-ttu-id="ca069-205">2 つのプラットフォームを組み合わせる利点の詳細については、「[Power Automate を使用した Office スクリプトの実行](../develop/power-automate-integration.md) または [手動による Power Automate フローからのスクリプトの呼び出し](excel-power-automate-manual.md) チュートリアルを試して、Office スクリプトを使用した Power Automate フローを作成する」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="ca069-205">Learn more about the advantages combining the two platforms in [Run Office Scripts with Power Automate](../develop/power-automate-integration.md) or try the [Call scripts from a manual Power Automate flow](excel-power-automate-manual.md) tutorial to create a Power Automate flow that uses an Office Script.</span></span>
