---
title: 自動電源自動化フローを使用してスクリプトを自動的に実行する
description: 自動的な外部トリガー (Outlook 経由でメールを受信する) を使用して、Power automatic を使用して、web 上で Excel の Office スクリプトを実行する方法についてのチュートリアルです。
ms.date: 07/01/2020
localization_priority: Priority
ms.openlocfilehash: fc98fb36fd5a8c5ef10bc3b767d6f5add0306246
ms.sourcegitcommit: edf58aed3cd38f57e5e7227465a1ef5515e15703
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 07/08/2020
ms.locfileid: "45081629"
---
# <a name="automatically-run-scripts-with-automated-power-automate-flows-preview"></a><span data-ttu-id="644b6-103">自動電源自動化フロー (プレビュー) を使用してスクリプトを自動的に実行する</span><span class="sxs-lookup"><span data-stu-id="644b6-103">Automatically run scripts with automated Power Automate flows (preview)</span></span>

<span data-ttu-id="644b6-104">このチュートリアルでは、自動[電源自動化](https://flow.microsoft.com)ワークフローを使用して web 上の Excel 用 Office スクリプトを使用する方法について説明します。</span><span class="sxs-lookup"><span data-stu-id="644b6-104">This tutorial teaches you how to use an Office Script for Excel on the web with an automated [Power Automate](https://flow.microsoft.com) workflow.</span></span> <span data-ttu-id="644b6-105">スクリプトは、電子メールを受信するたびに自動的に実行され、Excel ブックに電子メールの情報を記録します。</span><span class="sxs-lookup"><span data-stu-id="644b6-105">Your script will automatically run each time you receive an email, recording information from the email in an Excel workbook.</span></span>

## <a name="prerequisites"></a><span data-ttu-id="644b6-106">前提条件</span><span class="sxs-lookup"><span data-stu-id="644b6-106">Prerequisites</span></span>

[!INCLUDE [Tutorial prerequisites](../includes/tutorial-prerequisites.md)]

> [!IMPORTANT]
> <span data-ttu-id="644b6-107">このチュートリアルでは、 [「Power オートメーションチュートリアルを使用して web 上の Excel で Office スクリプトを実行する」](excel-power-automate-manual.md)を完了していることを前提としています。</span><span class="sxs-lookup"><span data-stu-id="644b6-107">This tutorial assumes you have completed the [Run Office Scripts in Excel on the web with Power Automate](excel-power-automate-manual.md) tutorial.</span></span>

## <a name="prepare-the-workbook"></a><span data-ttu-id="644b6-108">ブックの準備</span><span class="sxs-lookup"><span data-stu-id="644b6-108">Prepare the workbook</span></span>

<span data-ttu-id="644b6-109">Power オートメーションは、ブックコンポーネントへのアクセスなどの[相対参照](../develop/power-automate-integration.md#avoid-using-relative-references)を使用できません `Workbook.getActiveWorksheet` 。</span><span class="sxs-lookup"><span data-stu-id="644b6-109">Power Automate can't use [relative references](../develop/power-automate-integration.md#avoid-using-relative-references) like `Workbook.getActiveWorksheet` to access workbook components.</span></span> <span data-ttu-id="644b6-110">そのため、Power オートメーションが参照できるように、名前が一貫したブックとワークシートが必要です。</span><span class="sxs-lookup"><span data-stu-id="644b6-110">So, we need a workbook and worksheet with consistent names for Power Automate to reference.</span></span>

1. <span data-ttu-id="644b6-111">**Myworkbook**という名前の新しいブックを作成します。</span><span class="sxs-lookup"><span data-stu-id="644b6-111">Create a new workbook named **MyWorkbook**.</span></span>

2. <span data-ttu-id="644b6-112">[**自動化**] タブに移動して、[**コードエディター**] を選択します。</span><span class="sxs-lookup"><span data-stu-id="644b6-112">Go to the **Automate** tab and select **Code Editor**.</span></span>

3. <span data-ttu-id="644b6-113">[**新しいスクリプト**] を選択します。</span><span class="sxs-lookup"><span data-stu-id="644b6-113">Select **New Script**.</span></span>

4. <span data-ttu-id="644b6-114">既存のコードを次のスクリプトに置き換え、[**実行**] を押します。</span><span class="sxs-lookup"><span data-stu-id="644b6-114">Replace the existing code with the following script and press **Run**.</span></span> <span data-ttu-id="644b6-115">これにより、ワークシート、テーブル、およびピボットテーブル名が一致するブックが設定されます。</span><span class="sxs-lookup"><span data-stu-id="644b6-115">This will setup the workbook with consistent worksheet, table, and PivotTable names.</span></span>

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

## <a name="create-an-office-script-for-your-automated-workflow"></a><span data-ttu-id="644b6-116">自動化されたワークフロー用の Office スクリプトを作成する</span><span class="sxs-lookup"><span data-stu-id="644b6-116">Create an Office Script for your automated workflow</span></span>

<span data-ttu-id="644b6-117">電子メールから情報をログに記録するスクリプトを作成してみましょう。</span><span class="sxs-lookup"><span data-stu-id="644b6-117">Let's create a script that logs information from an email.</span></span> <span data-ttu-id="644b6-118">最もメールを受信する曜日と、そのメールを送信している一意の送信者の数を知りたいと考えています。</span><span class="sxs-lookup"><span data-stu-id="644b6-118">We want to know how which days of the week we receive the most mail and how many unique senders are sending that mail.</span></span> <span data-ttu-id="644b6-119">ブックに**は、\*\*\*\*日付**、曜日、**電子メールアドレス**、および**件名**の列を持つテーブルがあります。</span><span class="sxs-lookup"><span data-stu-id="644b6-119">Our workbook has a table with **Date**, **Day of the week**, **Email address**, and **Subject** columns.</span></span> <span data-ttu-id="644b6-120">また、このワークシートに**は、曜日と\*\*\*\*電子メールアドレス**(行階層) に対してピボットされたピボットテーブルもあります。</span><span class="sxs-lookup"><span data-stu-id="644b6-120">Our worksheet also has a PivotTable that is pivoting on the **Day of the week** and **Email address** (those are the row hierarchies).</span></span> <span data-ttu-id="644b6-121">一意の**件名**の数は、表示される集計情報 (データ階層) です。</span><span class="sxs-lookup"><span data-stu-id="644b6-121">The count of unique **Subjects** is the aggregated information being displayed (the data hierarchy).</span></span> <span data-ttu-id="644b6-122">メールテーブルを更新した後、スクリプトによってピボットテーブルが更新されるようになります。</span><span class="sxs-lookup"><span data-stu-id="644b6-122">We'll have our script refresh that PivotTable after updating the email table.</span></span>

1. <span data-ttu-id="644b6-123">**コードエディター**で、[**新しいスクリプト**] を選択します。</span><span class="sxs-lookup"><span data-stu-id="644b6-123">From within the **Code Editor**, select **New Script**.</span></span>

2. <span data-ttu-id="644b6-124">このチュートリアルで後で作成するフローによって、受信した各電子メールについてのスクリプト情報が送信されます。</span><span class="sxs-lookup"><span data-stu-id="644b6-124">The flow that we'll create later in the tutorial will send our script information about each email that's received.</span></span> <span data-ttu-id="644b6-125">スクリプトは、関数内のパラメーターを使用して、その入力を受け入れる必要があり `main` ます。</span><span class="sxs-lookup"><span data-stu-id="644b6-125">The script needs to accept that input through parameters in the `main` function.</span></span> <span data-ttu-id="644b6-126">既定のスクリプトを次のスクリプトに置き換えます。</span><span class="sxs-lookup"><span data-stu-id="644b6-126">Replace the default script with the following script:</span></span>

    ```TypeScript
    function main(
      workbook: ExcelScript.Workbook,
      from: string,
      dateReceived: string,
      subject: string) {

    }
    ```

3. <span data-ttu-id="644b6-127">このスクリプトには、ブックのテーブルとピボットテーブルへのアクセス権が必要です。</span><span class="sxs-lookup"><span data-stu-id="644b6-127">The script needs access to the workbook's table and PivotTable.</span></span> <span data-ttu-id="644b6-128">次のコードをスクリプトの本文に追加します。その後、次のコードを開き `{` ます。</span><span class="sxs-lookup"><span data-stu-id="644b6-128">Add the following code to the body of the script, after the opening `{`:</span></span>

    ```TypeScript
    // Get the email table.
    let emailWorksheet = workbook.getWorksheet("Emails");
    let table = emailWorksheet.getTable("EmailTable");
  
    // Get the PivotTable.
    let pivotTableWorksheet = workbook.getWorksheet("SubjectPivot");
    let pivotTable = pivotTableWorksheet.getPivotTable("Pivot");
    ```

4. <span data-ttu-id="644b6-129">`dateReceived`パラメーターの型がである `string` 。</span><span class="sxs-lookup"><span data-stu-id="644b6-129">The `dateReceived` parameter is of type `string`.</span></span> <span data-ttu-id="644b6-130">曜日を簡単に取得できるように、を[ `Date` オブジェクト](../develop/javascript-objects.md#date)に変換しましょう。</span><span class="sxs-lookup"><span data-stu-id="644b6-130">Let's convert that to a [`Date` object](../develop/javascript-objects.md#date) so we can easily get the day of the week.</span></span> <span data-ttu-id="644b6-131">その後、その日の番号の値をより読みやすいバージョンにマップする必要があります。</span><span class="sxs-lookup"><span data-stu-id="644b6-131">After doing that, we'll need to map the day's number value to a more readable version.</span></span> <span data-ttu-id="644b6-132">次のコードをスクリプトの最後に追加してから、閉じる前にし `}` ます。</span><span class="sxs-lookup"><span data-stu-id="644b6-132">Add the following code to the end of your script, before the closing `}`:</span></span>

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

5. <span data-ttu-id="644b6-133">文字列には `subject` 、"RE:" という返信タグを含めることができます。</span><span class="sxs-lookup"><span data-stu-id="644b6-133">The `subject` string may include the "RE:" reply tag.</span></span> <span data-ttu-id="644b6-134">これを文字列から削除して、同じスレッドの電子メールがテーブルの同じ件名を持つようにしましょう。</span><span class="sxs-lookup"><span data-stu-id="644b6-134">Let's remove that from the string so that emails in the same thread have the same subject for the table.</span></span> <span data-ttu-id="644b6-135">次のコードをスクリプトの最後に追加してから、閉じる前にし `}` ます。</span><span class="sxs-lookup"><span data-stu-id="644b6-135">Add the following code to the end of your script, before the closing `}`:</span></span>

    ```TypeScript
    // Remove the reply tag from the email subject to group emails on the same thread.
    let subjectText = subject.replace("Re: ", "");
    subjectText = subjectText.replace("RE: ", "");
    ```

6. <span data-ttu-id="644b6-136">これで、電子メールデータの形式が希望どおりになったので、電子メールの表に行を追加しましょう。</span><span class="sxs-lookup"><span data-stu-id="644b6-136">Now that the email data has been formatted to our liking, let's add a row to the email table.</span></span> <span data-ttu-id="644b6-137">次のコードをスクリプトの最後に追加してから、閉じる前にし `}` ます。</span><span class="sxs-lookup"><span data-stu-id="644b6-137">Add the following code to the end of your script, before the closing `}`:</span></span>

    ```TypeScript
    // Add the parsed text to the table.
    table.addRow(-1, [dateReceived, dayText, from, subjectText]);
    ```

7. <span data-ttu-id="644b6-138">最後に、ピボットテーブルが更新されていることを確認してみましょう。</span><span class="sxs-lookup"><span data-stu-id="644b6-138">Finally, let's make sure the PivotTable is refreshed.</span></span> <span data-ttu-id="644b6-139">次のコードをスクリプトの最後に追加してから、閉じる前にし `}` ます。</span><span class="sxs-lookup"><span data-stu-id="644b6-139">Add the following code to the end of your script, before the closing `}`:</span></span>

    ```TypeScript
    // Refresh the PivotTable to include the new row.
    pivotTable.refresh();
    ```

8. <span data-ttu-id="644b6-140">スクリプト**レコード**の名前を変更し、[**スクリプトの保存**] をクリックします。</span><span class="sxs-lookup"><span data-stu-id="644b6-140">Rename your script **Record Email** and press **Save script**.</span></span>

<span data-ttu-id="644b6-141">これで、パワー自動化ワークフローのためのスクリプトの準備が整いました。</span><span class="sxs-lookup"><span data-stu-id="644b6-141">Your script is now ready for a Power Automate workflow.</span></span> <span data-ttu-id="644b6-142">次のスクリプトのようになります。</span><span class="sxs-lookup"><span data-stu-id="644b6-142">It should look like the following script:</span></span>

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

## <a name="create-an-automated-workflow-with-power-automate"></a><span data-ttu-id="644b6-143">Power 自動化を使用して自動化されたワークフローを作成する</span><span class="sxs-lookup"><span data-stu-id="644b6-143">Create an automated workflow with Power Automate</span></span>

1. <span data-ttu-id="644b6-144">[パワー自動化プレビューサイト](https://flow.microsoft.com)にサインインします。</span><span class="sxs-lookup"><span data-stu-id="644b6-144">Sign in to the [Power Automate preview site](https://flow.microsoft.com).</span></span>

2. <span data-ttu-id="644b6-145">画面の左側に表示されるメニューで、[**作成**] を押します。</span><span class="sxs-lookup"><span data-stu-id="644b6-145">In the menu that's displayed on the left side of the screen, press **Create**.</span></span> <span data-ttu-id="644b6-146">これにより、新しいワークフローを作成する方法の一覧が表示されます。</span><span class="sxs-lookup"><span data-stu-id="644b6-146">This brings you to list of ways to create new workflows.</span></span>

    ![パワー自動化の [作成] ボタン。](../images/power-automate-tutorial-1.png)

3. <span data-ttu-id="644b6-148">[**空白から開始**] セクションで、[**自動フロー**] を選択します。</span><span class="sxs-lookup"><span data-stu-id="644b6-148">In the **Start from blank** section, select **Automated flow**.</span></span> <span data-ttu-id="644b6-149">これにより、電子メールの受信など、イベントによってトリガーされるワークフローが作成されます。</span><span class="sxs-lookup"><span data-stu-id="644b6-149">This creates a workflow triggered by an event, such as receiving an email.</span></span>

    ![電源自動化の [フローの自動化] オプション。](../images/power-automate-params-tutorial-1.png)

4. <span data-ttu-id="644b6-151">表示されるダイアログウィンドウで、[**フロー名**] テキストボックスにフローの名前を入力します。</span><span class="sxs-lookup"><span data-stu-id="644b6-151">In the dialog window that appears, enter a name for your flow in the **Flow name** text box.</span></span> <span data-ttu-id="644b6-152">次に、[**フローのトリガーを選択して**ください] の一覧から **、新しい電子メールを受信するタイミング**を選択します。</span><span class="sxs-lookup"><span data-stu-id="644b6-152">Then select **When a new email arrives** from the list of options under **Choose your flow's trigger**.</span></span> <span data-ttu-id="644b6-153">検索ボックスを使用してオプションを検索する必要がある場合があります。</span><span class="sxs-lookup"><span data-stu-id="644b6-153">You may need to search for the option using the search box.</span></span> <span data-ttu-id="644b6-154">最後に、[**作成**] を押します。</span><span class="sxs-lookup"><span data-stu-id="644b6-154">Finally, press **Create**.</span></span>

    ![「新しい電子メールの到着」オプションを示す [パワー・自動化] の [自動フロー] ウィンドウの構築の一部。](../images/power-automate-params-tutorial-2.png)

    > [!NOTE]
    > <span data-ttu-id="644b6-156">このチュートリアルでは、Outlook を使用します。</span><span class="sxs-lookup"><span data-stu-id="644b6-156">This tutorial uses Outlook.</span></span> <span data-ttu-id="644b6-157">代わりに、優先する電子メールサービスを自由に使用できますが、一部のオプションは異なる場合があります。</span><span class="sxs-lookup"><span data-stu-id="644b6-157">Feel free to use your preferred email service instead, though some options may be different.</span></span>

5. <span data-ttu-id="644b6-158">**新しい手順**を押します。</span><span class="sxs-lookup"><span data-stu-id="644b6-158">Press **New step**.</span></span>

6. <span data-ttu-id="644b6-159">[**標準**] タブを選択し、[ **Excel Online (Business)**] を選択します。</span><span class="sxs-lookup"><span data-stu-id="644b6-159">Select the **Standard** tab, then select **Excel Online (Business)**.</span></span>

    ![Excel Online (Business) の電源自動化オプション。](../images/power-automate-tutorial-4.png)

7. <span data-ttu-id="644b6-161">[**アクション**] で、[**スクリプトを実行する (プレビュー)**] を選択します。</span><span class="sxs-lookup"><span data-stu-id="644b6-161">Under **Actions**, select **Run script (preview)**.</span></span>

    ![実行スクリプトのパワー自動処理オプション (プレビュー)。](../images/power-automate-tutorial-5.png)

8. <span data-ttu-id="644b6-163">**実行スクリプト**コネクタについて、次の設定を指定します。</span><span class="sxs-lookup"><span data-stu-id="644b6-163">Specify the following settings for the **Run script** connector:</span></span>

    - <span data-ttu-id="644b6-164">**場所**: OneDrive for business</span><span class="sxs-lookup"><span data-stu-id="644b6-164">**Location**: OneDrive for Business</span></span>
    - <span data-ttu-id="644b6-165">**ドキュメントライブラリ**: OneDrive</span><span class="sxs-lookup"><span data-stu-id="644b6-165">**Document Library**: OneDrive</span></span>
    - <span data-ttu-id="644b6-166">**ファイル**: MyWorkbook.xlsx</span><span class="sxs-lookup"><span data-stu-id="644b6-166">**File**: MyWorkbook.xlsx</span></span>
    - <span data-ttu-id="644b6-167">**スクリプト**: メールの録音</span><span class="sxs-lookup"><span data-stu-id="644b6-167">**Script**: Record Email</span></span>
    - <span data-ttu-id="644b6-168">**from**: From *(Outlook の動的コンテンツ)*</span><span class="sxs-lookup"><span data-stu-id="644b6-168">**from**: From *(dynamic content from Outlook)*</span></span>
    - <span data-ttu-id="644b6-169">**dateReceived**: 受信時刻 *(Outlook からの動的なコンテンツ)*</span><span class="sxs-lookup"><span data-stu-id="644b6-169">**dateReceived**: Received Time *(dynamic content from Outlook)*</span></span>
    - <span data-ttu-id="644b6-170">**件名**: 件名 *(Outlook の動的コンテンツ)*</span><span class="sxs-lookup"><span data-stu-id="644b6-170">**subject**: Subject *(dynamic content from Outlook)*</span></span>

    <span data-ttu-id="644b6-171">*スクリプトのパラメーターは、スクリプトが選択された後にのみ表示されることに注意してください。*</span><span class="sxs-lookup"><span data-stu-id="644b6-171">*Note that the parameters for the script will only appear once the script is selected.*</span></span>

    ![実行スクリプトのパワー自動処理オプション (プレビュー)。](../images/power-automate-params-tutorial-3.png)

9. <span data-ttu-id="644b6-173">[**保存**します。</span><span class="sxs-lookup"><span data-stu-id="644b6-173">Press **Save**.</span></span>

<span data-ttu-id="644b6-174">これでフローが有効になります。</span><span class="sxs-lookup"><span data-stu-id="644b6-174">Your flow is now enabled.</span></span> <span data-ttu-id="644b6-175">Outlook を使用して電子メールを受信するたびに、スクリプトが自動的に実行されます。</span><span class="sxs-lookup"><span data-stu-id="644b6-175">It will automatically run your script each time you receive an email through Outlook.</span></span>

## <a name="manage-the-script-in-power-automate"></a><span data-ttu-id="644b6-176">パワー自動化でスクリプトを管理する</span><span class="sxs-lookup"><span data-stu-id="644b6-176">Manage the script in Power Automate</span></span>

1. <span data-ttu-id="644b6-177">[メインパワーの自動化] ページで、[**マイフロー**] を選択します。</span><span class="sxs-lookup"><span data-stu-id="644b6-177">From the main Power Automate page, select **My flows**.</span></span>

    ![パワー自動化の [マイフロー] ボタン。](../images/power-automate-tutorial-7.png)

2. <span data-ttu-id="644b6-179">フローを選択します。</span><span class="sxs-lookup"><span data-stu-id="644b6-179">Select your flow.</span></span> <span data-ttu-id="644b6-180">ここに、実行履歴が表示されます。</span><span class="sxs-lookup"><span data-stu-id="644b6-180">Here you can see the run history.</span></span> <span data-ttu-id="644b6-181">ページを更新するか、[すべての**実行**の更新] ボタンをクリックすると、履歴を更新できます。</span><span class="sxs-lookup"><span data-stu-id="644b6-181">You can refresh the page or press the refresh **All runs** button to update the history.</span></span> <span data-ttu-id="644b6-182">このフローは、電子メールの受信後すぐにトリガーされます。</span><span class="sxs-lookup"><span data-stu-id="644b6-182">The flow will trigger shortly after an email is received.</span></span> <span data-ttu-id="644b6-183">自分のメールを送信してフローをテストします。</span><span class="sxs-lookup"><span data-stu-id="644b6-183">Test the flow by sending yourself mail.</span></span>

<span data-ttu-id="644b6-184">フローがトリガーされ、スクリプトが正常に実行されると、ブックのテーブルとピボットテーブルの更新が表示されます。</span><span class="sxs-lookup"><span data-stu-id="644b6-184">When the flow is triggered and successfully runs your script, you should see the workbook's table and PivotTable update.</span></span>

![フローが2回実行された後の電子メールテーブル。](../images/power-automate-params-tutorial-4.png)

![フローが2回実行された後のピボットテーブル。](../images/power-automate-params-tutorial-5.png)

## <a name="next-steps"></a><span data-ttu-id="644b6-187">次の手順</span><span class="sxs-lookup"><span data-stu-id="644b6-187">Next steps</span></span>

<span data-ttu-id="644b6-188">Office スクリプトを Power オートメーションで接続する方法の詳細については、「 [Power オートメーションで Office スクリプトを実行](../develop/power-automate-integration.md)する」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="644b6-188">Visit [Run Office Scripts with Power Automate](../develop/power-automate-integration.md) to learn more about connecting Office Scripts with Power Automate.</span></span>

<span data-ttu-id="644b6-189">また、[自動タスクリマインダーのサンプルシナリオ](../resources/scenarios/task-reminders.md)を参照して、Office スクリプトと Teams のアダプティブカードを組み合わせたパワーオートメーションを組み合わせる方法を確認することもできます。</span><span class="sxs-lookup"><span data-stu-id="644b6-189">You can also check out the [Automated task reminders sample scenario](../resources/scenarios/task-reminders.md) to learn how to combine Office Scripts and Power Automate with Teams Adaptive Cards.</span></span>
