---
title: 自動で実行される Power Automate フロー内で、データをスクリプトに渡す
description: メールを受信し、フロー データをスクリプトに渡すときに、Power Automate を使用して Excel on the web 用の Office スクリプトを実行する方法について説明します。
ms.date: 07/14/2020
localization_priority: Priority
ms.openlocfilehash: c024891e187f22b7d10f6e9d52d262dc2ec4057f
ms.sourcegitcommit: ebd1079c7e2695ac0e7e4c616f2439975e196875
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 07/17/2020
ms.locfileid: "45160482"
---
# <a name="pass-data-to-scripts-in-an-automatically-run-power-automate-flow-preview"></a><span data-ttu-id="cbf57-103">自動で実行される Power Automate フロー内で、データをスクリプトに渡す(プレビュー)</span><span class="sxs-lookup"><span data-stu-id="cbf57-103">Pass data to scripts in an automatically-run Power Automate flow (preview)</span></span>

<span data-ttu-id="cbf57-104">このチュートリアルでは、自動化された [Power Automate](https://flow.microsoft.com) ワークフローを使用して、Excel on the web 用の Office スクリプトを実行する方法について説明します。</span><span class="sxs-lookup"><span data-stu-id="cbf57-104">This tutorial teaches you how to use an Office Script for Excel on the web with an automated [Power Automate](https://flow.microsoft.com) workflow.</span></span> <span data-ttu-id="cbf57-105">スクリプトは、メールを受信したときに自動的に実行されます。また、Excel ブック内のメールから情報を記録します。</span><span class="sxs-lookup"><span data-stu-id="cbf57-105">Your script will automatically run each time you receive an email, recording information from the email in an Excel workbook.</span></span>

## <a name="prerequisites"></a><span data-ttu-id="cbf57-106">前提条件</span><span class="sxs-lookup"><span data-stu-id="cbf57-106">Prerequisites</span></span>

[!INCLUDE [Tutorial prerequisites](../includes/power-automate-tutorial-prerequisites.md)]

> [!IMPORTANT]
> <span data-ttu-id="cbf57-107">このチュートリアルは、お客様が[「Power Automate を使用して、Excel on the web で Office スクリプトを実行する」](excel-power-automate-manual.md)のチュートリアルを既に完了していることを前提にしています。</span><span class="sxs-lookup"><span data-stu-id="cbf57-107">This tutorial assumes you have completed the [Run Office Scripts in Excel on the web with Power Automate](excel-power-automate-manual.md) tutorial.</span></span>

## <a name="prepare-the-workbook"></a><span data-ttu-id="cbf57-108">ブックを準備する</span><span class="sxs-lookup"><span data-stu-id="cbf57-108">Prepare the workbook</span></span>

<span data-ttu-id="cbf57-109">Power Automate は、`Workbook.getActiveWorksheet`のような[相対参照](../develop/power-automate-integration.md#avoid-using-relative-references)を使用して、ブック コンポーネントにアクセスすることはできません。</span><span class="sxs-lookup"><span data-stu-id="cbf57-109">Power Automate can't use [relative references](../develop/power-automate-integration.md#avoid-using-relative-references) like `Workbook.getActiveWorksheet` to access workbook components.</span></span> <span data-ttu-id="cbf57-110">したがって、Power Automate が参照できるように、名前が統一されたブックとワークシートが必要です。</span><span class="sxs-lookup"><span data-stu-id="cbf57-110">So, we need a workbook and worksheet with consistent names for Power Automate to reference.</span></span>

1. <span data-ttu-id="cbf57-111">**MyWorkbook** という名前の新しいブックを作成します。</span><span class="sxs-lookup"><span data-stu-id="cbf57-111">Create a new workbook named **MyWorkbook**.</span></span>

2. <span data-ttu-id="cbf57-112">**[オートメーション]** タブに移動して **[コード エディター]** を選択します。</span><span class="sxs-lookup"><span data-stu-id="cbf57-112">Go to the **Automate** tab and select **Code Editor**.</span></span>

3. <span data-ttu-id="cbf57-113">**[新しいスクリプト]** を選択します。</span><span class="sxs-lookup"><span data-stu-id="cbf57-113">Select **New Script**.</span></span>

4. <span data-ttu-id="cbf57-114">既存のコードを次のスクリプトで置き換え、**[実行]** を押します。</span><span class="sxs-lookup"><span data-stu-id="cbf57-114">Replace the existing code with the following script and press **Run**.</span></span> <span data-ttu-id="cbf57-115">これにより、統一されたワークシート、テーブル、ピボットテーブルの名前でブックが設定されます。</span><span class="sxs-lookup"><span data-stu-id="cbf57-115">This will setup the workbook with consistent worksheet, table, and PivotTable names.</span></span>

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
      let pivotWorksheet = workbook.addWorksheet("SubjectPivot");
      let newPivotTable = workbook.addPivotTable("Pivot", "EmailTable", pivotWorksheet.getRange("A3:C20"));

      // Setup the pivot hierarchies
      newPivotTable.addRowHierarchy(newPivotTable.getHierarchy("Day of the week"));
      newPivotTable.addRowHierarchy(newPivotTable.getHierarchy("Email address"));
      newPivotTable.addDataHierarchy(newPivotTable.getHierarchy("Subject"));
    }
    ```

## <a name="create-an-office-script-for-your-automated-workflow"></a><span data-ttu-id="cbf57-116">自動化されたワークフロー用のオフィス スクリプトの作成</span><span class="sxs-lookup"><span data-stu-id="cbf57-116">Create an Office Script for your automated workflow</span></span>

<span data-ttu-id="cbf57-117">メールから情報をログに記録するスクリプトを作成してみましょう。</span><span class="sxs-lookup"><span data-stu-id="cbf57-117">Let's create a script that logs information from an email.</span></span> <span data-ttu-id="cbf57-118">最も多くのメールを受信する曜日と、そのメールを送信する固有の送信者の数について知る必要があります。</span><span class="sxs-lookup"><span data-stu-id="cbf57-118">We want to know how which days of the week we receive the most mail and how many unique senders are sending that mail.</span></span> <span data-ttu-id="cbf57-119">ブックには、**[日付]**、**[曜日]**、**[メールアドレス]**、**[件名]** の列を含むテーブルがあります。</span><span class="sxs-lookup"><span data-stu-id="cbf57-119">Our workbook has a table with **Date**, **Day of the week**, **Email address**, and **Subject** columns.</span></span> <span data-ttu-id="cbf57-120">また、ワークシートには、 **[曜日]** と **メールアドレス** (行階層)にピボットしている、ピボットテーブルがあります。</span><span class="sxs-lookup"><span data-stu-id="cbf57-120">Our worksheet also has a PivotTable that is pivoting on the **Day of the week** and **Email address** (those are the row hierarchies).</span></span> <span data-ttu-id="cbf57-121">一意の **[件名]** の数は、表示されている集計情報（データ階層）です。</span><span class="sxs-lookup"><span data-stu-id="cbf57-121">The count of unique **Subjects** is the aggregated information being displayed (the data hierarchy).</span></span> <span data-ttu-id="cbf57-122">メール テーブルを更新した後に、スクリプトがピボットテーブルを更新するようにします。</span><span class="sxs-lookup"><span data-stu-id="cbf57-122">We'll have our script refresh that PivotTable after updating the email table.</span></span>

1. <span data-ttu-id="cbf57-123">**[コード エディター]** 内で、**[新しいスクリプト]** を選択します。</span><span class="sxs-lookup"><span data-stu-id="cbf57-123">From within the **Code Editor**, select **New Script**.</span></span>

2. <span data-ttu-id="cbf57-124">このチュートリアルの後半で作成するフローでは、受信した各メールに関するスクリプト情報を送信します。</span><span class="sxs-lookup"><span data-stu-id="cbf57-124">The flow that we'll create later in the tutorial will send our script information about each email that's received.</span></span> <span data-ttu-id="cbf57-125">スクリプトは、`main`関数のパラメーターを使用して、その入力を受け付ける必要があります。</span><span class="sxs-lookup"><span data-stu-id="cbf57-125">The script needs to accept that input through parameters in the `main` function.</span></span> <span data-ttu-id="cbf57-126">既定のスクリプトを次のスクリプトに置き換えます。</span><span class="sxs-lookup"><span data-stu-id="cbf57-126">Replace the default script with the following script:</span></span>

    ```TypeScript
    function main(
      workbook: ExcelScript.Workbook,
      from: string,
      dateReceived: string,
      subject: string) {

    }
    ```

3. <span data-ttu-id="cbf57-127">スクリプトには、ブックのテーブルとピボットテーブルにアクセスする必要があります。</span><span class="sxs-lookup"><span data-stu-id="cbf57-127">The script needs access to the workbook's table and PivotTable.</span></span> <span data-ttu-id="cbf57-128">`{` を開いた後、次のコードをスクリプトの本文に追加 します。</span><span class="sxs-lookup"><span data-stu-id="cbf57-128">Add the following code to the body of the script, after the opening `{`:</span></span>

    ```TypeScript
    // Get the email table.
    let emailWorksheet = workbook.getWorksheet("Emails");
    let table = emailWorksheet.getTable("EmailTable");
  
    // Get the PivotTable.
    let pivotTableWorksheet = workbook.getWorksheet("SubjectPivot");
    let pivotTable = pivotTableWorksheet.getPivotTable("Pivot");
    ```

4. <span data-ttu-id="cbf57-129">`dateReceived` パラメーターのタイプは `string` です。</span><span class="sxs-lookup"><span data-stu-id="cbf57-129">The `dateReceived` parameter is of type `string`.</span></span> <span data-ttu-id="cbf57-130">それを [`Date` オブジェクト](../develop/javascript-objects.md#date)に変換して、簡単に曜日を取得できるようにしましょう。</span><span class="sxs-lookup"><span data-stu-id="cbf57-130">Let's convert that to a [`Date` object](../develop/javascript-objects.md#date) so we can easily get the day of the week.</span></span> <span data-ttu-id="cbf57-131">その後、日の数値をより読みやすいバージョンにマッピングする必要があります。</span><span class="sxs-lookup"><span data-stu-id="cbf57-131">After doing that, we'll need to map the day's number value to a more readable version.</span></span> <span data-ttu-id="cbf57-132">`}` を閉じる前に、スクリプトの最後に次のコードを追加します。</span><span class="sxs-lookup"><span data-stu-id="cbf57-132">Add the following code to the end of your script, before the closing `}`:</span></span>

    ```TypeScript
    // Parse the received date string.
    let date = new Date(dateReceived);

    // Convert number representing the day of the week into the name of the day.
    let dayText : string;
    switch (date.getDay()) {
      case 0:
        dayText = "Sunday";
        break;
      case 1:
        dayText = "Monday";
        break;
      case 2:
        dayText = "Tuesday";
        break;
      case 3:
        dayText = "Wednesday";
        break;
      case 4:
        dayText = "Thursday";
        break;
      case 5:
        dayText = "Friday";
        break;
      default:
        dayText = "Saturday";
        break;
    }
    ```

5. <span data-ttu-id="cbf57-133">`subject` 文字列には、"RE:" という返信タグを含めることができます。</span><span class="sxs-lookup"><span data-stu-id="cbf57-133">The `subject` string may include the "RE:" reply tag.</span></span> <span data-ttu-id="cbf57-134">同じスレッドのメールがテーブルに対して同じ件名になるよう、文字列からそれを削除します。</span><span class="sxs-lookup"><span data-stu-id="cbf57-134">Let's remove that from the string so that emails in the same thread have the same subject for the table.</span></span> <span data-ttu-id="cbf57-135">`}` を閉じる前に、スクリプトの最後に次のコードを追加します。</span><span class="sxs-lookup"><span data-stu-id="cbf57-135">Add the following code to the end of your script, before the closing `}`:</span></span>

    ```TypeScript
    // Remove the reply tag from the email subject to group emails on the same thread.
    let subjectText = subject.replace("Re: ", "");
    subjectText = subjectText.replace("RE: ", "");
    ```

6. <span data-ttu-id="cbf57-136">これでメールのデータがお好みの書式に設定されたので、メール テーブルに行を追加しましょう。</span><span class="sxs-lookup"><span data-stu-id="cbf57-136">Now that the email data has been formatted to our liking, let's add a row to the email table.</span></span> <span data-ttu-id="cbf57-137">`}` を閉じる前に、スクリプトの最後に次のコードを追加します。</span><span class="sxs-lookup"><span data-stu-id="cbf57-137">Add the following code to the end of your script, before the closing `}`:</span></span>

    ```TypeScript
    // Add the parsed text to the table.
    table.addRow(-1, [dateReceived, dayText, from, subjectText]);
    ```

7. <span data-ttu-id="cbf57-138">最後に、ピボットテーブルを更新されていることを確認しましょう。</span><span class="sxs-lookup"><span data-stu-id="cbf57-138">Finally, let's make sure the PivotTable is refreshed.</span></span> <span data-ttu-id="cbf57-139">`}` を閉じる前に、スクリプトの最後に次のコードを追加します。</span><span class="sxs-lookup"><span data-stu-id="cbf57-139">Add the following code to the end of your script, before the closing `}`:</span></span>

    ```TypeScript
    // Refresh the PivotTable to include the new row.
    pivotTable.refresh();
    ```

8. <span data-ttu-id="cbf57-140">スクリプトの名前を **[メールを記録]** に変更し、**[スクリプトの保存]** を押します。</span><span class="sxs-lookup"><span data-stu-id="cbf57-140">Rename your script **Record Email** and press **Save script**.</span></span>

<span data-ttu-id="cbf57-141">これで、スクリプトは Power Automate ワークフローで使用できるようになりました。</span><span class="sxs-lookup"><span data-stu-id="cbf57-141">Your script is now ready for a Power Automate workflow.</span></span> <span data-ttu-id="cbf57-142">次のようにスクリプトが表示されます。</span><span class="sxs-lookup"><span data-stu-id="cbf57-142">It should look like the following script:</span></span>

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
  let pivotTableWorksheet = workbook.getWorksheet("Pivot");
  let pivotTable = pivotTableWorksheet.getPivotTable("SubjectPivot");

  // Parse the received date string.
  let date = new Date(dateReceived);

  // Convert number representing the day of the week into the name of the day.
  let dayText: string;
  switch (date.getDay()) {
    case 0:
      dayText = "Sunday";
      break;
    case 1:
      dayText = "Monday";
      break;
    case 2:
      dayText = "Tuesday";
      break;
    case 3:
      dayText = "Wednesday";
      break;
    case 4:
      dayText = "Thursday";
      break;
    case 5:
      dayText = "Friday";
      break;
    default:
      dayText = "Saturday";
      break;
  }

  // Remove the reply tag from the email subject to group emails on the same thread.
  let subjectText = subject.replace("Re: ", "");
  subjectText = subjectText.replace("RE: ", "");

  // Add the parsed text to the table.
  table.addRow(-1, [dateReceived, dayText, from, subjectText]);

  // Refresh the PivotTable to include the new row.
  pivotTable.refresh();
}
```

## <a name="create-an-automated-workflow-with-power-automate"></a><span data-ttu-id="cbf57-143">Power Automate を使用して自動化されたワークフローを作成する</span><span class="sxs-lookup"><span data-stu-id="cbf57-143">Create an automated workflow with Power Automate</span></span>

1. <span data-ttu-id="cbf57-144">[「Power Automate のサイト」](https://flow.microsoft.com)にサインインします。</span><span class="sxs-lookup"><span data-stu-id="cbf57-144">Sign in to the [Power Automate site](https://flow.microsoft.com).</span></span>

2. <span data-ttu-id="cbf57-145">画面の左側に表示されるメニューで、**[作成]** を押します。</span><span class="sxs-lookup"><span data-stu-id="cbf57-145">In the menu that's displayed on the left side of the screen, press **Create**.</span></span> <span data-ttu-id="cbf57-146">これにより、新しいワークフローを作成する方法の一覧を表示できます。</span><span class="sxs-lookup"><span data-stu-id="cbf57-146">This brings you to list of ways to create new workflows.</span></span>

    ![Power Automate の [作成] ボタン。](../images/power-automate-tutorial-1.png)

3. <span data-ttu-id="cbf57-148">**[白紙から初める]** セクションで、**[自動フロー]** を選択します。</span><span class="sxs-lookup"><span data-stu-id="cbf57-148">In the **Start from blank** section, select **Automated flow**.</span></span> <span data-ttu-id="cbf57-149">これにより、メールの受信などのイベントによってトリガーされるワークフローが作成されます。</span><span class="sxs-lookup"><span data-stu-id="cbf57-149">This creates a workflow triggered by an event, such as receiving an email.</span></span>

    ![Power Automate の自動フロー オプション。](../images/power-automate-params-tutorial-1.png)

4. <span data-ttu-id="cbf57-151">ダイアログ ウインドウが表示されたら、**[フロー名]** のテキスト ボックスに、フローの名前を入力します。</span><span class="sxs-lookup"><span data-stu-id="cbf57-151">In the dialog window that appears, enter a name for your flow in the **Flow name** text box.</span></span> <span data-ttu-id="cbf57-152">次に、**[フローのトリガーを選択]** の下のオプションの一覧から、**[新しいメールが届いたとき]** を選択します。</span><span class="sxs-lookup"><span data-stu-id="cbf57-152">Then select **When a new email arrives** from the list of options under **Choose your flow's trigger**.</span></span> <span data-ttu-id="cbf57-153">検索ボックスを使用して、オプションを検索することが必要になる場合があります。</span><span class="sxs-lookup"><span data-stu-id="cbf57-153">You may need to search for the option using the search box.</span></span> <span data-ttu-id="cbf57-154">最後に、**[作成]** を押します。</span><span class="sxs-lookup"><span data-stu-id="cbf57-154">Finally, press **Create**.</span></span>

    ![Power Automate の [自動フローの作成]ウィンドウの一部で、”新しいメールが届きました” オプションが表示されます。](../images/power-automate-params-tutorial-2.png)

    > [!NOTE]
    > <span data-ttu-id="cbf57-156">このチュートリアルでは、Outlook を使用します。</span><span class="sxs-lookup"><span data-stu-id="cbf57-156">This tutorial uses Outlook.</span></span> <span data-ttu-id="cbf57-157">代わりに、お好きなメール サービスを自由に使用することもできますが、一部のオプションは異なる場合があります。</span><span class="sxs-lookup"><span data-stu-id="cbf57-157">Feel free to use your preferred email service instead, though some options may be different.</span></span>

5. <span data-ttu-id="cbf57-158">**[新しいステップ]** を押します。</span><span class="sxs-lookup"><span data-stu-id="cbf57-158">Press **New step**.</span></span>

6. <span data-ttu-id="cbf57-159">**[標準]** タブを選択し、**Excel Online (ビジネス)** を選択します。</span><span class="sxs-lookup"><span data-stu-id="cbf57-159">Select the **Standard** tab, then select **Excel Online (Business)**.</span></span>

    ![Excel Online (ビジネス) 用の Power Automate オプション。](../images/power-automate-tutorial-4.png)

7. <span data-ttu-id="cbf57-161">**[アクション]** の下から、**[スクリプトの実行 (プレビュー)]** を選択します。</span><span class="sxs-lookup"><span data-stu-id="cbf57-161">Under **Actions**, select **Run script (preview)**.</span></span>

    ![スクリプトの実行 (プレビュー)用の Power Automate アクションのオプション。](../images/power-automate-tutorial-5.png)

8. <span data-ttu-id="cbf57-163">**スクリプトの実行**コネクタには、次の設定を指定します。</span><span class="sxs-lookup"><span data-stu-id="cbf57-163">Specify the following settings for the **Run script** connector:</span></span>

    - <span data-ttu-id="cbf57-164">**場所**: OneDrive for Business</span><span class="sxs-lookup"><span data-stu-id="cbf57-164">**Location**: OneDrive for Business</span></span>
    - <span data-ttu-id="cbf57-165">**ドキュメント ライブラリ**: OneDrive</span><span class="sxs-lookup"><span data-stu-id="cbf57-165">**Document Library**: OneDrive</span></span>
    - <span data-ttu-id="cbf57-166">**ファイル**: MyWorkbook.xlsx</span><span class="sxs-lookup"><span data-stu-id="cbf57-166">**File**: MyWorkbook.xlsx</span></span>
    - <span data-ttu-id="cbf57-167">**スクリプト**: メールの記録</span><span class="sxs-lookup"><span data-stu-id="cbf57-167">**Script**: Record Email</span></span>
    - <span data-ttu-id="cbf57-168">**から**: *(Outlook の動的コンテンツ)*</span><span class="sxs-lookup"><span data-stu-id="cbf57-168">**from**: From *(dynamic content from Outlook)*</span></span>
    - <span data-ttu-id="cbf57-169">**dateReceived**: 受信時刻 *(Outlook の動的コンテンツ)*</span><span class="sxs-lookup"><span data-stu-id="cbf57-169">**dateReceived**: Received Time *(dynamic content from Outlook)*</span></span>
    - <span data-ttu-id="cbf57-170">**件名**: 件名 *(Outlook の動的コンテンツ)*</span><span class="sxs-lookup"><span data-stu-id="cbf57-170">**subject**: Subject *(dynamic content from Outlook)*</span></span>

    <span data-ttu-id="cbf57-171">*スクリプトのパラメーターは、スクリプトが選択された後にのみ表示されるので、注意してください。*</span><span class="sxs-lookup"><span data-stu-id="cbf57-171">*Note that the parameters for the script will only appear once the script is selected.*</span></span>

    ![スクリプトの実行 (プレビュー)用の Power Automate アクションのオプション。](../images/power-automate-params-tutorial-3.png)

9. <span data-ttu-id="cbf57-173">**[保存]** を押します。</span><span class="sxs-lookup"><span data-stu-id="cbf57-173">Press **Save**.</span></span>

<span data-ttu-id="cbf57-174">フローが有効になります。</span><span class="sxs-lookup"><span data-stu-id="cbf57-174">Your flow is now enabled.</span></span> <span data-ttu-id="cbf57-175">Outlook でメールを受信するたびに、スクリプトが自動的に実行されます。</span><span class="sxs-lookup"><span data-stu-id="cbf57-175">It will automatically run your script each time you receive an email through Outlook.</span></span>

## <a name="manage-the-script-in-power-automate"></a><span data-ttu-id="cbf57-176">Power Automate でスクリプトを管理する</span><span class="sxs-lookup"><span data-stu-id="cbf57-176">Manage the script in Power Automate</span></span>

1. <span data-ttu-id="cbf57-177">Power Automate のメイン ページで、**[自分のフロー]** を選択します。</span><span class="sxs-lookup"><span data-stu-id="cbf57-177">From the main Power Automate page, select **My flows**.</span></span>

    ![Power Automate の [自分のフロー] ボタン。](../images/power-automate-tutorial-7.png)

2. <span data-ttu-id="cbf57-179">フローを選択します。</span><span class="sxs-lookup"><span data-stu-id="cbf57-179">Select your flow.</span></span> <span data-ttu-id="cbf57-180">ここでは、実行履歴を表示することができます。</span><span class="sxs-lookup"><span data-stu-id="cbf57-180">Here you can see the run history.</span></span> <span data-ttu-id="cbf57-181">ページを更新するか、**[すべての実行]** を更新するボタンを押して、履歴を更新することができます。</span><span class="sxs-lookup"><span data-stu-id="cbf57-181">You can refresh the page or press the refresh **All runs** button to update the history.</span></span> <span data-ttu-id="cbf57-182">フローは、メールを受信するとすぐにトリガーされます。</span><span class="sxs-lookup"><span data-stu-id="cbf57-182">The flow will trigger shortly after an email is received.</span></span> <span data-ttu-id="cbf57-183">メッセージを送信してフローをテストします。</span><span class="sxs-lookup"><span data-stu-id="cbf57-183">Test the flow by sending yourself mail.</span></span>

<span data-ttu-id="cbf57-184">フローがトリガーされて、スクリプトが正常に実行されると、ブックのテーブルとピボットテーブルの更新が表示されます。</span><span class="sxs-lookup"><span data-stu-id="cbf57-184">When the flow is triggered and successfully runs your script, you should see the workbook's table and PivotTable update.</span></span>

![フローが数回実行された後の、メール テーブル。](../images/power-automate-params-tutorial-4.png)

![フローが数回実行された後の、ピボットテーブル。](../images/power-automate-params-tutorial-5.png)

## <a name="next-steps"></a><span data-ttu-id="cbf57-187">次の手順</span><span class="sxs-lookup"><span data-stu-id="cbf57-187">Next steps</span></span>

<span data-ttu-id="cbf57-188">Office スクリプトを Power Automate に接続する方法に関する詳細については、 [「Power Automate で Office スクリプトを実行する」](../develop/power-automate-integration.md)を参照してください。</span><span class="sxs-lookup"><span data-stu-id="cbf57-188">Visit [Run Office Scripts with Power Automate](../develop/power-automate-integration.md) to learn more about connecting Office Scripts with Power Automate.</span></span>

<span data-ttu-id="cbf57-189">[「自動タスク リマインダーのサンプル シナリオ」](../resources/scenarios/task-reminders.md)では、Office スクリプトと Power Automate を Teams アダプティブ カードと組み合わせる方法についても説明します。</span><span class="sxs-lookup"><span data-stu-id="cbf57-189">You can also check out the [Automated task reminders sample scenario](../resources/scenarios/task-reminders.md) to learn how to combine Office Scripts and Power Automate with Teams Adaptive Cards.</span></span>
