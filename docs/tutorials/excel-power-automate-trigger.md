---
title: 自動で実行される Power Automate フロー内で、データをスクリプトに渡す
description: メールを受信し、フロー データをスクリプトに渡すときに、Power Automate を使用して Excel on the web 用の Office スクリプトを実行する方法について説明します。
ms.date: 12/28/2020
localization_priority: Priority
ms.openlocfilehash: c67c1f958099698201db91f368a287805d5aea0d
ms.sourcegitcommit: 45ffe3dbd2c834b78592ad35928cf8096f5e80bc
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 04/14/2021
ms.locfileid: "51754895"
---
# <a name="pass-data-to-scripts-in-an-automatically-run-power-automate-flow-preview"></a><span data-ttu-id="9bba3-103">自動で実行される Power Automate フロー内で、データをスクリプトに渡す(プレビュー)</span><span class="sxs-lookup"><span data-stu-id="9bba3-103">Pass data to scripts in an automatically-run Power Automate flow (preview)</span></span>

<span data-ttu-id="9bba3-104">このチュートリアルでは、自動化された [Power Automate](https://flow.microsoft.com) ワークフローを使用して、Excel on the web 用の Office スクリプトを実行する方法について説明します。</span><span class="sxs-lookup"><span data-stu-id="9bba3-104">This tutorial teaches you how to use an Office Script for Excel on the web with an automated [Power Automate](https://flow.microsoft.com) workflow.</span></span> <span data-ttu-id="9bba3-105">スクリプトは、メールを受信したときに自動的に実行されます。また、Excel ブック内のメールから情報を記録します。</span><span class="sxs-lookup"><span data-stu-id="9bba3-105">Your script will automatically run each time you receive an email, recording information from the email in an Excel workbook.</span></span> <span data-ttu-id="9bba3-106">別のアプリケーションから Office スクリプトにデータを渡すことができるようになると、自動プロセスの柔軟性と自由性が大きく向上します。</span><span class="sxs-lookup"><span data-stu-id="9bba3-106">Being able to pass data from other applications into an Office Script gives you a great deal of flexibility and freedom in your automated processes.</span></span>

> [!TIP]
> <span data-ttu-id="9bba3-107">Office スクリプトを初めて使用する場合は、チュートリアルの「[Excel on the web で Office スクリプトを記録、編集、作成する](excel-tutorial.md)」から始めることをお勧めします。</span><span class="sxs-lookup"><span data-stu-id="9bba3-107">If you are new to Office Scripts, we recommend starting with the [Record, edit, and create Office Scripts in Excel on the web](excel-tutorial.md) tutorial.</span></span> <span data-ttu-id="9bba3-108">Power Automate を初めて使用する場合は、チュートリアルの「[手動 Power Automate フローからスクリプトを呼び出す](excel-power-automate-manual.md)」から始めることを勧めします。</span><span class="sxs-lookup"><span data-stu-id="9bba3-108">If you are new to Power Automate, we recommend starting with the [Call scripts from a manual Power Automate flow](excel-power-automate-manual.md) tutorial.</span></span> <span data-ttu-id="9bba3-109">[Office スクリプトは TypeScript を使用](../overview/code-editor-environment.md)します。このチュートリアルは、JavaScript や TypeScript について初級から中級レベルの知識を持つユーザーを対象としています。</span><span class="sxs-lookup"><span data-stu-id="9bba3-109">[Office Scripts use TypeScript](../overview/code-editor-environment.md) and this tutorial is intended for people with beginner to intermediate-level knowledge of JavaScript or TypeScript.</span></span> <span data-ttu-id="9bba3-110">JavaScript を使い慣れていない場合は、「[Mozilla の JavaScript チュートリアル](https://developer.mozilla.org/docs/Web/JavaScript/Guide/Introduction)」から始めることをお勧めします。</span><span class="sxs-lookup"><span data-stu-id="9bba3-110">If you're new to JavaScript, we recommend starting with the [Mozilla JavaScript tutorial](https://developer.mozilla.org/docs/Web/JavaScript/Guide/Introduction).</span></span>

## <a name="prerequisites"></a><span data-ttu-id="9bba3-111">前提条件</span><span class="sxs-lookup"><span data-stu-id="9bba3-111">Prerequisites</span></span>

[!INCLUDE [Tutorial prerequisites](../includes/power-automate-tutorial-prerequisites.md)]

## <a name="prepare-the-workbook"></a><span data-ttu-id="9bba3-112">ブックを準備する</span><span class="sxs-lookup"><span data-stu-id="9bba3-112">Prepare the workbook</span></span>

<span data-ttu-id="9bba3-113">Power Automate では、ブック コンポーネントにアクセスするために `Workbook.getActiveWorksheet` などの[相対参照](../testing/power-automate-troubleshooting.md#avoid-using-relative-references)を使わないようにする必要があります。</span><span class="sxs-lookup"><span data-stu-id="9bba3-113">Power Automate shouldn't use [relative references](../testing/power-automate-troubleshooting.md#avoid-using-relative-references) like `Workbook.getActiveWorksheet` to access workbook components.</span></span> <span data-ttu-id="9bba3-114">したがって、Power Automate が参照できるように、名前が統一されたブックとワークシートが必要です。</span><span class="sxs-lookup"><span data-stu-id="9bba3-114">So, we need a workbook and worksheet with consistent names for Power Automate to reference.</span></span>

1. <span data-ttu-id="9bba3-115">**MyWorkbook** という名前の新しいブックを作成します。</span><span class="sxs-lookup"><span data-stu-id="9bba3-115">Create a new workbook named **MyWorkbook**.</span></span>

2. <span data-ttu-id="9bba3-116">**[オートメーション]** タブに移動して **[すべてのスクリプト]** を選択します。</span><span class="sxs-lookup"><span data-stu-id="9bba3-116">Go to the **Automate** tab and select **All Scripts**.</span></span>

3. <span data-ttu-id="9bba3-117">**[新しいスクリプト]** を選択します。</span><span class="sxs-lookup"><span data-stu-id="9bba3-117">Select **New Script**.</span></span>

4. <span data-ttu-id="9bba3-118">既存のコードを次のスクリプトで置き換え、**[実行]** を押します。</span><span class="sxs-lookup"><span data-stu-id="9bba3-118">Replace the existing code with the following script and press **Run**.</span></span> <span data-ttu-id="9bba3-119">これにより、統一されたワークシート、テーブル、ピボットテーブルの名前でブックが設定されます。</span><span class="sxs-lookup"><span data-stu-id="9bba3-119">This will setup the workbook with consistent worksheet, table, and PivotTable names.</span></span>

    ```TypeScript
    function main(workbook: ExcelScript.Workbook) {
      // Add a new worksheet to store our email table
      let emailsSheet = workbook.addWorksheet("Emails");

      // Add data and create a table
      emailsSheet.getRange("A1:D1").setValues([
        ["Date", "Day of the week", "Email address", "Subject"]
      ]);
      let newTable = workbook.addTable(emailsSheet.getRange("A1:D2"), true);
      newTable.setName("EmailTable");

      // Add a new PivotTable to a new worksheet
      let pivotWorksheet = workbook.addWorksheet("Subjects");
      let newPivotTable = workbook.addPivotTable("Pivot", "EmailTable", pivotWorksheet.getRange("A3:C20"));

      // Setup the pivot hierarchies
      newPivotTable.addRowHierarchy(newPivotTable.getHierarchy("Day of the week"));
      newPivotTable.addRowHierarchy(newPivotTable.getHierarchy("Email address"));
      newPivotTable.addDataHierarchy(newPivotTable.getHierarchy("Subject"));
    }
    ```

## <a name="create-an-office-script"></a><span data-ttu-id="9bba3-120">Office スクリプトを作成する</span><span class="sxs-lookup"><span data-stu-id="9bba3-120">Create an Office Script</span></span>

<span data-ttu-id="9bba3-121">メールから情報をログに記録するスクリプトを作成してみましょう。</span><span class="sxs-lookup"><span data-stu-id="9bba3-121">Let's create a script that logs information from an email.</span></span> <span data-ttu-id="9bba3-122">最も多くのメールを受信する曜日と、そのメールを送信する固有の送信者の数について知る必要があります。</span><span class="sxs-lookup"><span data-stu-id="9bba3-122">We want to know how which days of the week we receive the most mail and how many unique senders are sending that mail.</span></span> <span data-ttu-id="9bba3-123">ブックには、**[日付]**、**[曜日]**、**[メールアドレス]**、**[件名]** の列を含むテーブルがあります。</span><span class="sxs-lookup"><span data-stu-id="9bba3-123">Our workbook has a table with **Date**, **Day of the week**, **Email address**, and **Subject** columns.</span></span> <span data-ttu-id="9bba3-124">また、ワークシートには、 **[曜日]** と **メールアドレス** (行階層)にピボットしている、ピボットテーブルがあります。</span><span class="sxs-lookup"><span data-stu-id="9bba3-124">Our worksheet also has a PivotTable that is pivoting on the **Day of the week** and **Email address** (those are the row hierarchies).</span></span> <span data-ttu-id="9bba3-125">一意の **[件名]** の数は、表示されている集計情報（データ階層）です。</span><span class="sxs-lookup"><span data-stu-id="9bba3-125">The count of unique **Subjects** is the aggregated information being displayed (the data hierarchy).</span></span> <span data-ttu-id="9bba3-126">メール テーブルを更新した後に、スクリプトがピボットテーブルを更新するようにします。</span><span class="sxs-lookup"><span data-stu-id="9bba3-126">We'll have our script refresh that PivotTable after updating the email table.</span></span>

1. <span data-ttu-id="9bba3-127">**[コード エディター]** 作業ウィンドウ内で、**[新しいスクリプト]** を選択します。</span><span class="sxs-lookup"><span data-stu-id="9bba3-127">From within the **Code Editor** task pane, select **New Script**.</span></span>

2. <span data-ttu-id="9bba3-128">このチュートリアルの後半で作成するフローでは、受信した各メールに関するスクリプト情報を送信します。</span><span class="sxs-lookup"><span data-stu-id="9bba3-128">The flow that we'll create later in the tutorial will send our script information about each email that's received.</span></span> <span data-ttu-id="9bba3-129">スクリプトは、`main`関数のパラメーターを使用して、その入力を受け付ける必要があります。</span><span class="sxs-lookup"><span data-stu-id="9bba3-129">The script needs to accept that input through parameters in the `main` function.</span></span> <span data-ttu-id="9bba3-130">既定のスクリプトを次のスクリプトに置き換えます。</span><span class="sxs-lookup"><span data-stu-id="9bba3-130">Replace the default script with the following script:</span></span>

    ```TypeScript
    function main(
      workbook: ExcelScript.Workbook,
      from: string,
      dateReceived: string,
      subject: string) {

    }
    ```

3. <span data-ttu-id="9bba3-131">スクリプトには、ブックのテーブルとピボットテーブルにアクセスする必要があります。</span><span class="sxs-lookup"><span data-stu-id="9bba3-131">The script needs access to the workbook's table and PivotTable.</span></span> <span data-ttu-id="9bba3-132">`{` を開いた後、次のコードをスクリプトの本文に追加 します。</span><span class="sxs-lookup"><span data-stu-id="9bba3-132">Add the following code to the body of the script, after the opening `{`:</span></span>

    ```TypeScript
    // Get the email table.
    let emailWorksheet = workbook.getWorksheet("Emails");
    let table = emailWorksheet.getTable("EmailTable");
  
    // Get the PivotTable.
    let pivotTableWorksheet = workbook.getWorksheet("Subjects");
    let pivotTable = pivotTableWorksheet.getPivotTable("Pivot");
    ```

4. <span data-ttu-id="9bba3-133">`dateReceived` パラメーターのタイプは `string` です。</span><span class="sxs-lookup"><span data-stu-id="9bba3-133">The `dateReceived` parameter is of type `string`.</span></span> <span data-ttu-id="9bba3-134">それを [`Date` オブジェクト](../develop/javascript-objects.md#date)に変換して、簡単に曜日を取得できるようにしましょう。</span><span class="sxs-lookup"><span data-stu-id="9bba3-134">Let's convert that to a [`Date` object](../develop/javascript-objects.md#date) so we can easily get the day of the week.</span></span> <span data-ttu-id="9bba3-135">その後、日の数値をより読みやすいバージョンにマッピングする必要があります。</span><span class="sxs-lookup"><span data-stu-id="9bba3-135">After doing that, we'll need to map the day's number value to a more readable version.</span></span> <span data-ttu-id="9bba3-136">`}` を閉じる前に、スクリプトの最後に次のコードを追加します。</span><span class="sxs-lookup"><span data-stu-id="9bba3-136">Add the following code to the end of your script, before the closing `}`:</span></span>

    ```TypeScript
      // Parse the received date string to determine the day of the week.
      let emailDate = new Date(dateReceived);
      let dayName = emailDate.toLocaleDateString("en-US", { weekday: 'long' });
    ```

5. <span data-ttu-id="9bba3-137">`subject` 文字列には、"RE:" という返信タグを含めることができます。</span><span class="sxs-lookup"><span data-stu-id="9bba3-137">The `subject` string may include the "RE:" reply tag.</span></span> <span data-ttu-id="9bba3-138">同じスレッドのメールがテーブルに対して同じ件名になるよう、文字列からそれを削除します。</span><span class="sxs-lookup"><span data-stu-id="9bba3-138">Let's remove that from the string so that emails in the same thread have the same subject for the table.</span></span> <span data-ttu-id="9bba3-139">`}` を閉じる前に、スクリプトの最後に次のコードを追加します。</span><span class="sxs-lookup"><span data-stu-id="9bba3-139">Add the following code to the end of your script, before the closing `}`:</span></span>

    ```TypeScript
    // Remove the reply tag from the email subject to group emails on the same thread.
    let subjectText = subject.replace("Re: ", "");
    subjectText = subjectText.replace("RE: ", "");
    ```

6. <span data-ttu-id="9bba3-140">これでメールのデータがお好みの書式に設定されたので、メール テーブルに行を追加しましょう。</span><span class="sxs-lookup"><span data-stu-id="9bba3-140">Now that the email data has been formatted to our liking, let's add a row to the email table.</span></span> <span data-ttu-id="9bba3-141">`}` を閉じる前に、スクリプトの最後に次のコードを追加します。</span><span class="sxs-lookup"><span data-stu-id="9bba3-141">Add the following code to the end of your script, before the closing `}`:</span></span>

    ```TypeScript
    // Add the parsed text to the table.
    table.addRow(-1, [dateReceived, dayName, from, subjectText]);
    ```

7. <span data-ttu-id="9bba3-142">最後に、ピボットテーブルを更新されていることを確認しましょう。</span><span class="sxs-lookup"><span data-stu-id="9bba3-142">Finally, let's make sure the PivotTable is refreshed.</span></span> <span data-ttu-id="9bba3-143">`}` を閉じる前に、スクリプトの最後に次のコードを追加します。</span><span class="sxs-lookup"><span data-stu-id="9bba3-143">Add the following code to the end of your script, before the closing `}`:</span></span>

    ```TypeScript
    // Refresh the PivotTable to include the new row.
    pivotTable.refresh();
    ```

8. <span data-ttu-id="9bba3-144">スクリプトの名前を **[メールを記録]** に変更し、**[スクリプトの保存]** を押します。</span><span class="sxs-lookup"><span data-stu-id="9bba3-144">Rename your script **Record Email** and press **Save script**.</span></span>

<span data-ttu-id="9bba3-145">これで、スクリプトは Power Automate ワークフローで使用できるようになりました。</span><span class="sxs-lookup"><span data-stu-id="9bba3-145">Your script is now ready for a Power Automate workflow.</span></span> <span data-ttu-id="9bba3-146">次のようにスクリプトが表示されます。</span><span class="sxs-lookup"><span data-stu-id="9bba3-146">It should look like the following script:</span></span>

```TypeScript
function main(
  workbook: ExcelScript.Workbook,
  from: string,
  dateReceived: string,
  subject: string) {
  // Get the email table.
  let emailWorksheet = workbook.getWorksheet("Emails");
  let table = emailWorksheet.getTable("EmailTable");

  // Get the PivotTable.
  let pivotTableWorksheet = workbook.getWorksheet("Subjects");
  let pivotTable = pivotTableWorksheet.getPivotTable("Pivot");

  // Parse the received date string to determine the day of the week.
  let emailDate = new Date(dateReceived);
  let dayName = emailDate.toLocaleDateString("en-US", { weekday: 'long' });

  // Remove the reply tag from the email subject to group emails on the same thread.
  let subjectText = subject.replace("Re: ", "");
  subjectText = subjectText.replace("RE: ", "");

  // Add the parsed text to the table.
  table.addRow(-1, [dateReceived, dayName, from, subjectText]);

  // Refresh the PivotTable to include the new row.
  pivotTable.refresh();
}
```

## <a name="create-an-automated-workflow-with-power-automate"></a><span data-ttu-id="9bba3-147">Power Automate を使用して自動化されたワークフローを作成する</span><span class="sxs-lookup"><span data-stu-id="9bba3-147">Create an automated workflow with Power Automate</span></span>

1. <span data-ttu-id="9bba3-148">[「Power Automate のサイト」](https://flow.microsoft.com)にサインインします。</span><span class="sxs-lookup"><span data-stu-id="9bba3-148">Sign in to the [Power Automate site](https://flow.microsoft.com).</span></span>

2. <span data-ttu-id="9bba3-149">画面の左側に表示されるメニューで、**[作成]** を押します。</span><span class="sxs-lookup"><span data-stu-id="9bba3-149">In the menu that's displayed on the left side of the screen, press **Create**.</span></span> <span data-ttu-id="9bba3-150">これにより、新しいワークフローを作成する方法の一覧を表示できます。</span><span class="sxs-lookup"><span data-stu-id="9bba3-150">This brings you to list of ways to create new workflows.</span></span>

    :::image type="content" source="../images/power-automate-tutorial-1.png" alt-text="Power Automate の [作成] ボタン。":::

3. <span data-ttu-id="9bba3-152">**[白紙から初める]** セクションで、**[自動フロー]** を選択します。</span><span class="sxs-lookup"><span data-stu-id="9bba3-152">In the **Start from blank** section, select **Automated flow**.</span></span> <span data-ttu-id="9bba3-153">これにより、メールの受信などのイベントによってトリガーされるワークフローが作成されます。</span><span class="sxs-lookup"><span data-stu-id="9bba3-153">This creates a workflow triggered by an event, such as receiving an email.</span></span>

    :::image type="content" source="../images/power-automate-params-tutorial-1.png" alt-text="Power Automate の自動化したフロー オプション":::

4. <span data-ttu-id="9bba3-155">ダイアログ ウインドウが表示されたら、**[フロー名]** のテキスト ボックスに、フローの名前を入力します。</span><span class="sxs-lookup"><span data-stu-id="9bba3-155">In the dialog window that appears, enter a name for your flow in the **Flow name** text box.</span></span> <span data-ttu-id="9bba3-156">次に、**[フローのトリガーを選択]** の下のオプションの一覧から、**[新しいメールが届いたとき]** を選択します。</span><span class="sxs-lookup"><span data-stu-id="9bba3-156">Then select **When a new email arrives** from the list of options under **Choose your flow's trigger**.</span></span> <span data-ttu-id="9bba3-157">検索ボックスを使用して、オプションを検索することが必要になる場合があります。</span><span class="sxs-lookup"><span data-stu-id="9bba3-157">You may need to search for the option using the search box.</span></span> <span data-ttu-id="9bba3-158">最後に、**[作成]** を押します。</span><span class="sxs-lookup"><span data-stu-id="9bba3-158">Finally, press **Create**.</span></span>

    :::image type="content" source="../images/power-automate-params-tutorial-2.png" alt-text="[フロー名] と [choose your folow’s trigger (フローのトリガーを選択する)] オプションを示す Power Automate Flow の一部。フロー名は &quot;メール フローの記録&quot; で、トリガーは [When a new email arrives in Outlook (Outlook で新しいメールを受信する場合)] オプションです。":::

    > [!NOTE]
    > <span data-ttu-id="9bba3-160">このチュートリアルでは、Outlook を使用します。</span><span class="sxs-lookup"><span data-stu-id="9bba3-160">This tutorial uses Outlook.</span></span> <span data-ttu-id="9bba3-161">代わりに、お好きなメール サービスを自由に使用することもできますが、一部のオプションは異なる場合があります。</span><span class="sxs-lookup"><span data-stu-id="9bba3-161">Feel free to use your preferred email service instead, though some options may be different.</span></span>

5. <span data-ttu-id="9bba3-162">**[新しいステップ]** を押します。</span><span class="sxs-lookup"><span data-stu-id="9bba3-162">Press **New step**.</span></span>

6. <span data-ttu-id="9bba3-163">**[標準]** タブを選択し、**Excel Online (ビジネス)** を選択します。</span><span class="sxs-lookup"><span data-stu-id="9bba3-163">Select the **Standard** tab, then select **Excel Online (Business)**.</span></span>

    :::image type="content" source="../images/power-automate-tutorial-4.png" alt-text="Power Automate の [Excel Online (Business)] オプション":::

7. <span data-ttu-id="9bba3-165">**[アクション]** の下の **[スクリプトの実行 (プレビュー)]** を選択します。</span><span class="sxs-lookup"><span data-stu-id="9bba3-165">Under **Actions**, select **Run script (preview)**.</span></span>

    :::image type="content" source="../images/power-automate-tutorial-5.png" alt-text="Power Automate の [スクリプトの実行 (プレビュー)] アクションのオプション":::

8. <span data-ttu-id="9bba3-167">次に、フロー ステップで使用するブック、スクリプト、およびスクリプトの入力引数を選択します。</span><span class="sxs-lookup"><span data-stu-id="9bba3-167">Next, you'll select the workbook, script, and script input arguments to use in the flow step.</span></span> <span data-ttu-id="9bba3-168">このチュートリアルでは、OneDrive に作成したブックを使用しますが、OneDrive サイトまたは SharePoint サイトでは任意のブックを使用できます。</span><span class="sxs-lookup"><span data-stu-id="9bba3-168">For the tutorial, you'll use the workbook you created in your OneDrive, but you could use any workbook in a OneDrive or SharePoint site.</span></span> <span data-ttu-id="9bba3-169">**スクリプトの実行** コネクタには、次の設定を指定します。</span><span class="sxs-lookup"><span data-stu-id="9bba3-169">Specify the following settings for the **Run script** connector:</span></span>

    - <span data-ttu-id="9bba3-170">**場所**: OneDrive for Business</span><span class="sxs-lookup"><span data-stu-id="9bba3-170">**Location**: OneDrive for Business</span></span>
    - <span data-ttu-id="9bba3-171">**ドキュメント ライブラリ**: OneDrive</span><span class="sxs-lookup"><span data-stu-id="9bba3-171">**Document Library**: OneDrive</span></span>
    - <span data-ttu-id="9bba3-172">**ファイル**: MyWorkbook.xlsx *(ファイル ブラウザーを使用して選択されています)*</span><span class="sxs-lookup"><span data-stu-id="9bba3-172">**File**: MyWorkbook.xlsx *(Chosen through the file browser)*</span></span>
    - <span data-ttu-id="9bba3-173">**スクリプト**: メールの記録</span><span class="sxs-lookup"><span data-stu-id="9bba3-173">**Script**: Record Email</span></span>
    - <span data-ttu-id="9bba3-174">**から**: *(Outlook の動的コンテンツ)*</span><span class="sxs-lookup"><span data-stu-id="9bba3-174">**from**: From *(dynamic content from Outlook)*</span></span>
    - <span data-ttu-id="9bba3-175">**dateReceived**: 受信時刻 *(Outlook の動的コンテンツ)*</span><span class="sxs-lookup"><span data-stu-id="9bba3-175">**dateReceived**: Received Time *(dynamic content from Outlook)*</span></span>
    - <span data-ttu-id="9bba3-176">**件名**: 件名 *(Outlook の動的コンテンツ)*</span><span class="sxs-lookup"><span data-stu-id="9bba3-176">**subject**: Subject *(dynamic content from Outlook)*</span></span>

    <span data-ttu-id="9bba3-177">*スクリプトのパラメーターは、スクリプトが選択された後にのみ表示されるので、注意してください。*</span><span class="sxs-lookup"><span data-stu-id="9bba3-177">*Note that the parameters for the script will only appear once the script is selected.*</span></span>

    :::image type="content" source="../images/power-automate-params-tutorial-3.png" alt-text="スクリプトを選択すると表示されるオプションを示す Power Automate のスクリプトを実行するアクション。":::

9. <span data-ttu-id="9bba3-179">**[保存]** を押します。</span><span class="sxs-lookup"><span data-stu-id="9bba3-179">Press **Save**.</span></span>

<span data-ttu-id="9bba3-180">フローが有効になります。</span><span class="sxs-lookup"><span data-stu-id="9bba3-180">Your flow is now enabled.</span></span> <span data-ttu-id="9bba3-181">Outlook でメールを受信するたびに、スクリプトが自動的に実行されます。</span><span class="sxs-lookup"><span data-stu-id="9bba3-181">It will automatically run your script each time you receive an email through Outlook.</span></span>

## <a name="manage-the-script-in-power-automate"></a><span data-ttu-id="9bba3-182">Power Automate でスクリプトを管理する</span><span class="sxs-lookup"><span data-stu-id="9bba3-182">Manage the script in Power Automate</span></span>

1. <span data-ttu-id="9bba3-183">Power Automate のメイン ページで、**[自分のフロー]** を選択します。</span><span class="sxs-lookup"><span data-stu-id="9bba3-183">From the main Power Automate page, select **My flows**.</span></span>

    :::image type="content" source="../images/power-automate-tutorial-7.png" alt-text="Power Automate の [自分のフロー] ボタン":::

2. <span data-ttu-id="9bba3-185">フローを選択します。</span><span class="sxs-lookup"><span data-stu-id="9bba3-185">Select your flow.</span></span> <span data-ttu-id="9bba3-186">ここでは、実行履歴を表示することができます。</span><span class="sxs-lookup"><span data-stu-id="9bba3-186">Here you can see the run history.</span></span> <span data-ttu-id="9bba3-187">ページを更新するか、**[すべての実行]** を更新するボタンを押して、履歴を更新することができます。</span><span class="sxs-lookup"><span data-stu-id="9bba3-187">You can refresh the page or press the refresh **All runs** button to update the history.</span></span> <span data-ttu-id="9bba3-188">フローは、メールを受信するとすぐにトリガーされます。</span><span class="sxs-lookup"><span data-stu-id="9bba3-188">The flow will trigger shortly after an email is received.</span></span> <span data-ttu-id="9bba3-189">メッセージを送信してフローをテストします。</span><span class="sxs-lookup"><span data-stu-id="9bba3-189">Test the flow by sending yourself mail.</span></span>

<span data-ttu-id="9bba3-190">フローがトリガーされて、スクリプトが正常に実行されると、ブックのテーブルとピボットテーブルの更新が表示されます。</span><span class="sxs-lookup"><span data-stu-id="9bba3-190">When the flow is triggered and successfully runs your script, you should see the workbook's table and PivotTable update.</span></span>

:::image type="content" source="../images/power-automate-params-tutorial-4.png" alt-text="フローの実行を 3 回行った後のメール テーブルを示すワークシート。":::

:::image type="content" source="../images/power-automate-params-tutorial-5.png" alt-text="フローの実行を 3 回行った後のピボット テーブルを示すワークシート。":::

## <a name="next-steps"></a><span data-ttu-id="9bba3-193">次の手順</span><span class="sxs-lookup"><span data-stu-id="9bba3-193">Next steps</span></span>

<span data-ttu-id="9bba3-194">「[自動で実行される Power Automate フローにスクリプトからデータを返す](excel-power-automate-returns.md)」のチュートリアルを完了します。</span><span class="sxs-lookup"><span data-stu-id="9bba3-194">Complete the [Return data from a script to an automatically-run Power Automate flow](excel-power-automate-returns.md) tutorial.</span></span> <span data-ttu-id="9bba3-195">このチュートリアルでは、スクリプトからフローにデータを返す方法を説明します。</span><span class="sxs-lookup"><span data-stu-id="9bba3-195">It teaches you how to return data from a script to the flow.</span></span>

<span data-ttu-id="9bba3-196">[「自動タスク リマインダーのサンプル シナリオ」](../resources/scenarios/task-reminders.md)では、Office スクリプトと Power Automate を Teams アダプティブ カードと組み合わせる方法についても説明します。</span><span class="sxs-lookup"><span data-stu-id="9bba3-196">You can also check out the [Automated task reminders sample scenario](../resources/scenarios/task-reminders.md) to learn how to combine Office Scripts and Power Automate with Teams Adaptive Cards.</span></span>
