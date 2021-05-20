---
title: 'Officeスクリプトのサンプル シナリオ: 自動タスク アラーム'
description: Power Automateとアダプティブ カードを使用するサンプルは、プロジェクト管理スプレッドシートでタスクのリマインダーを自動化します。
ms.date: 11/30/2020
localization_priority: Normal
ms.openlocfilehash: c254a627da8442c0974263908a41275182740b6e
ms.sourcegitcommit: 4687693f02fc90a57ba30c461f35046e02e6f5fb
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 05/19/2021
ms.locfileid: "52545611"
---
# <a name="office-scripts-sample-scenario-automated-task-reminders"></a><span data-ttu-id="f72b7-103">Officeスクリプトのサンプル シナリオ: 自動タスク アラーム</span><span class="sxs-lookup"><span data-stu-id="f72b7-103">Office Scripts sample scenario: Automated task reminders</span></span>

<span data-ttu-id="f72b7-104">このシナリオでは、プロジェクトを管理しています。</span><span class="sxs-lookup"><span data-stu-id="f72b7-104">In this scenario you're managing a project.</span></span> <span data-ttu-id="f72b7-105">Excelワークシートを使用して、毎月従業員のステータスを追跡します。</span><span class="sxs-lookup"><span data-stu-id="f72b7-105">You use an Excel worksheet to track your employees' status every month.</span></span> <span data-ttu-id="f72b7-106">多くの場合、ユーザーにステータスを入力するよう促す必要があるため、リマインダープロセスを自動化することにしました。</span><span class="sxs-lookup"><span data-stu-id="f72b7-106">You often need to remind people to fill out their status, so you've decided to automate that reminder process.</span></span>

<span data-ttu-id="f72b7-107">ステータスフィールドが欠落している人にメッセージを送信するPower Automateフローを作成し、その回答をスプレッドシートに適用します。</span><span class="sxs-lookup"><span data-stu-id="f72b7-107">You'll create a Power Automate flow to message people with missing status fields and apply their responses to the spreadsheet.</span></span> <span data-ttu-id="f72b7-108">これを行うには、ブックの操作を処理するスクリプトのペアを開発します。</span><span class="sxs-lookup"><span data-stu-id="f72b7-108">To do this, you'll develop a pair of scripts to handle the working with the workbook.</span></span> <span data-ttu-id="f72b7-109">最初のスクリプトは、空白の状態のユーザーのリストを取得し、2 番目のスクリプトは、右側の行にステータス文字列を追加します。</span><span class="sxs-lookup"><span data-stu-id="f72b7-109">The first script gets a list of people with blank statuses and the second script adds a status string to the right row.</span></span> <span data-ttu-id="f72b7-110">また[、Teamsアダプティブカード](/microsoftteams/platform/task-modules-and-cards/what-are-cards)を使用して、従業員に通知から直接ステータスを入力させます。</span><span class="sxs-lookup"><span data-stu-id="f72b7-110">You'll also make use of [Teams Adaptive Cards](/microsoftteams/platform/task-modules-and-cards/what-are-cards) to have employees enter their status directly from the notification.</span></span>

## <a name="scripting-skills-covered"></a><span data-ttu-id="f72b7-111">スクリプティングのスキルをカバー</span><span class="sxs-lookup"><span data-stu-id="f72b7-111">Scripting skills covered</span></span>

- <span data-ttu-id="f72b7-112">Power Automateでのフローの作成</span><span class="sxs-lookup"><span data-stu-id="f72b7-112">Create flows in Power Automate</span></span>
- <span data-ttu-id="f72b7-113">スクリプトへのデータの受け渡し</span><span class="sxs-lookup"><span data-stu-id="f72b7-113">Pass data to scripts</span></span>
- <span data-ttu-id="f72b7-114">スクリプトからデータを返す</span><span class="sxs-lookup"><span data-stu-id="f72b7-114">Return data from scripts</span></span>
- <span data-ttu-id="f72b7-115">Teamsアダプティブカード</span><span class="sxs-lookup"><span data-stu-id="f72b7-115">Teams Adaptive Cards</span></span>
- <span data-ttu-id="f72b7-116">テーブル</span><span class="sxs-lookup"><span data-stu-id="f72b7-116">Tables</span></span>

## <a name="prerequisites"></a><span data-ttu-id="f72b7-117">前提条件</span><span class="sxs-lookup"><span data-stu-id="f72b7-117">Prerequisites</span></span>

<span data-ttu-id="f72b7-118">このシナリオでは[、Power Automate](https://flow.microsoft.com)と[Microsoft Teams](https://www.microsoft.com/microsoft-365/microsoft-teams/group-chat-software)を使用します。</span><span class="sxs-lookup"><span data-stu-id="f72b7-118">This scenario uses [Power Automate](https://flow.microsoft.com) and [Microsoft Teams](https://www.microsoft.com/microsoft-365/microsoft-teams/group-chat-software).</span></span> <span data-ttu-id="f72b7-119">Office スクリプトの開発に使用するアカウントに関連付ける必要があります。</span><span class="sxs-lookup"><span data-stu-id="f72b7-119">You will need both associated with the account that you use for developing Office Scripts.</span></span> <span data-ttu-id="f72b7-120">これらのアプリケーションについて学び、これらのアプリケーションを操作するための Microsoft Developer サブスクリプションへの無料アクセスについては[、Microsoft 365開発者プログラム](https://developer.microsoft.com/microsoft-365/dev-program)への参加を検討してください。</span><span class="sxs-lookup"><span data-stu-id="f72b7-120">For free access to a Microsoft Developer subscription to learn about and work with these applications, consider joining the [Microsoft 365 Developer Program](https://developer.microsoft.com/microsoft-365/dev-program).</span></span>

## <a name="setup-instructions"></a><span data-ttu-id="f72b7-121">セットアップ手順</span><span class="sxs-lookup"><span data-stu-id="f72b7-121">Setup instructions</span></span>

1. <span data-ttu-id="f72b7-122"><a href="task-reminders.xlsx">OneDriveにtask-reminders.xlsx</a>をダウンロードします。</span><span class="sxs-lookup"><span data-stu-id="f72b7-122">Download <a href="task-reminders.xlsx">task-reminders.xlsx</a> to your OneDrive.</span></span>

2. <span data-ttu-id="f72b7-123">ブックをExcel on the webで開きます。</span><span class="sxs-lookup"><span data-stu-id="f72b7-123">Open the workbook in Excel on the web.</span></span>

3. <span data-ttu-id="f72b7-124">[ **自動化** ] タブで、[ **すべてのスクリプト]** を開きます。</span><span class="sxs-lookup"><span data-stu-id="f72b7-124">Under the **Automate** tab, open **All Scripts**.</span></span>

4. <span data-ttu-id="f72b7-125">まず、スプレッドシートに不足しているステータス レポートを持つすべての従業員を取得するスクリプトが必要です。</span><span class="sxs-lookup"><span data-stu-id="f72b7-125">First, we need a script to get all the employees with status reports that are missing from the spreadsheet.</span></span> <span data-ttu-id="f72b7-126">[ **コード エディター** ] 作業ウィンドウで[ **新規スクリプト]** を押し、次のスクリプトをエディタに貼り付けます。</span><span class="sxs-lookup"><span data-stu-id="f72b7-126">In the **Code Editor** task pane, press **New Script** and paste the following script into the editor.</span></span>

    ```TypeScript
    /**
     * This script looks for missing status reports in a project management table.
     *
     * @returns An array of Employee objects (containing their names and emails).
     */
    function main(workbook: ExcelScript.Workbook): Employee[] {
      // Get the first worksheet and the first table on that worksheet.
      let sheet = workbook.getFirstWorksheet()
      let table = sheet.getTables()[0];

      // Give the column indices names matching their expected content.
      const NAME_INDEX = 0;
      const EMAIL_INDEX = 1;
      const STATUS_REPORT_INDEX = 2;

      // Get the data for the whole table.
      let bodyRangeValues = table.getRangeBetweenHeaderAndTotal().getValues();

      // Create the array of Employee objects to return.
      let people: Employee[] = [];

      // Loop through the table and check each row for completion.
      for (let i = 0; i < bodyRangeValues.length; i++) {
        let row = bodyRangeValues[i];
        if (row[STATUS_REPORT_INDEX] === "") {
          // Save the email to return.
          people.push({ name: row[NAME_INDEX].toString(), email: row[EMAIL_INDEX].toString() });
        }
      }

      // Log the array to verify we're getting the right rows.
      console.log(people);

      // Return the array of Employees.
      return people;
    }

    /**
     * An interface representing an employee.
     * An array of Employees will be returned from the script
     * for the Power Automate flow.
     */
    interface Employee {
      name: string;
      email: string;
    }
    ```

5. <span data-ttu-id="f72b7-127">スクリプトを **Get People** という名前で保存します。</span><span class="sxs-lookup"><span data-stu-id="f72b7-127">Save the script with the name **Get People**.</span></span>

6. <span data-ttu-id="f72b7-128">次に、ステータス レポート カードを処理し、新しい情報をスプレッドシートに配置する 2 番目のスクリプトが必要です。</span><span class="sxs-lookup"><span data-stu-id="f72b7-128">Next, we need a second script to process the status report cards and put the new information in the spreadsheet.</span></span> <span data-ttu-id="f72b7-129">[ **コード エディター** ] 作業ウィンドウで[ **新規スクリプト]** を押し、次のスクリプトをエディタに貼り付けます。</span><span class="sxs-lookup"><span data-stu-id="f72b7-129">In the **Code Editor** task pane, press **New Script** and paste the following script into the editor.</span></span>

    ```TypeScript
    /**
     * This script applies the results of a Teams Adaptive Card about
     * a status update to a project management table.
     *
     * @param senderEmail - The email address of the employee updating their status.
     * @param statusReportResponse - The employee's status report.
     */
    function main(workbook: ExcelScript.Workbook,
      senderEmail: string,
      statusReportResponse: string) {

      // Get the first worksheet and the first table in that worksheet.
      let sheet = workbook.getFirstWorksheet();
      let table = sheet.getTables()[0];

      // Give the column indices names matching their expected content.
      const NAME_INDEX = 0;
      const EMAIL_INDEX = 1;
      const STATUS_REPORT_INDEX = 2;

      // Get the range and data for the whole table.
      let bodyRange = table.getRangeBetweenHeaderAndTotal();
      let tableRowCount = bodyRange.getRowCount();
      let bodyRangeValues = bodyRange.getValues();

      // Create a flag to denote success.
      let statusAdded = false;

      // Loop through the table and check each row for a matching email address.
      for (let i = 0; i < tableRowCount && !statusAdded; i++) {
        let row = bodyRangeValues[i];

        // Check if the row's email address matches.
        if (row[EMAIL_INDEX] === senderEmail) {
          // Add the Teams Adaptive Card response to the table.
          bodyRange.getCell(i, STATUS_REPORT_INDEX).setValues([
            [statusReportResponse]
          ]);
          statusAdded = true;
        }
      }

      // If successful, log the status update.
      if (statusAdded) {
        console.log(
          `Successfully added status report for ${senderEmail} containing: ${statusReportResponse}`
        );
      }
    }
    ```

7. <span data-ttu-id="f72b7-130">スクリプトを保存ステータスという名前で **保存します**。</span><span class="sxs-lookup"><span data-stu-id="f72b7-130">Save the script with the name **Save Status**.</span></span>

8. <span data-ttu-id="f72b7-131">ここで、フローを作成する必要があります。</span><span class="sxs-lookup"><span data-stu-id="f72b7-131">Now, we need to create the flow.</span></span> <span data-ttu-id="f72b7-132">[Power Automate](https://flow.microsoft.com/)を開きます。</span><span class="sxs-lookup"><span data-stu-id="f72b7-132">Open [Power Automate](https://flow.microsoft.com/).</span></span>

    > [!TIP]
    > <span data-ttu-id="f72b7-133">フローを作成したことがない場合は、チュートリアル[「Power Automateを含むスクリプトの使用を開始](../../tutorials/excel-power-automate-manual.md)する」を参照して基本を学習してください。</span><span class="sxs-lookup"><span data-stu-id="f72b7-133">If you haven't created a flow before, please check out our tutorial [Start using scripts with Power Automate](../../tutorials/excel-power-automate-manual.md) to learn the basics.</span></span>

9. <span data-ttu-id="f72b7-134">新しい **インスタントフロー** を作成する:</span><span class="sxs-lookup"><span data-stu-id="f72b7-134">Create a new **Instant flow**.</span></span>

10. <span data-ttu-id="f72b7-135">オプションから **[フローを手動でトリガー** ] を選択し、[ **作成]** を押します。</span><span class="sxs-lookup"><span data-stu-id="f72b7-135">Choose **Manually trigger a flow** from the options and press **Create**.</span></span>

11. <span data-ttu-id="f72b7-136">フローは、空のステータスフィールドを持つすべての従業員を取得するために **Get People** スクリプトを呼び出す必要があります。</span><span class="sxs-lookup"><span data-stu-id="f72b7-136">The flow needs to call the **Get People** script to get all the employees with empty status fields.</span></span> <span data-ttu-id="f72b7-137">**[新規ステップ]** を押して、[**オンライン (ビジネス)] Excel** 選択します。</span><span class="sxs-lookup"><span data-stu-id="f72b7-137">Press **New step** and select **Excel Online (Business)**.</span></span> <span data-ttu-id="f72b7-138">[ **アクション]** で、[ **スクリプトの実行**] を選択します。</span><span class="sxs-lookup"><span data-stu-id="f72b7-138">Under **Actions**, select **Run script**.</span></span> <span data-ttu-id="f72b7-139">フローステップに対して以下のエントリを入力します。</span><span class="sxs-lookup"><span data-stu-id="f72b7-139">Provide the following entries for the flow step:</span></span>

    - <span data-ttu-id="f72b7-140">**場所**: OneDrive for Business</span><span class="sxs-lookup"><span data-stu-id="f72b7-140">**Location**: OneDrive for Business</span></span>
    - <span data-ttu-id="f72b7-141">**ドキュメント ライブラリ**: OneDrive</span><span class="sxs-lookup"><span data-stu-id="f72b7-141">**Document Library**: OneDrive</span></span>
    - <span data-ttu-id="f72b7-142">**ファイル**: task-reminders.xlsx *(ファイルブラウザを使用して選択)*</span><span class="sxs-lookup"><span data-stu-id="f72b7-142">**File**: task-reminders.xlsx *(Chosen through the file browser)*</span></span>
    - <span data-ttu-id="f72b7-143">**スクリプト**: 人をゲット</span><span class="sxs-lookup"><span data-stu-id="f72b7-143">**Script**: Get People</span></span>

    :::image type="content" source="../../images/scenario-task-reminders-first-flow-step.png" alt-text="スクリプトフローの最初の実行ステップを示すPower Automateフロー":::

12. <span data-ttu-id="f72b7-145">次に、フローは、スクリプトによって返される配列内の各従業員を処理する必要があります。</span><span class="sxs-lookup"><span data-stu-id="f72b7-145">Next, the flow needs to process each Employee in the array returned by the script.</span></span> <span data-ttu-id="f72b7-146">**[新規ステップ]** を押して、[**アダプティブ カードをTeamsユーザーにポストし、応答を待つ**] を選択します。</span><span class="sxs-lookup"><span data-stu-id="f72b7-146">Press **New step** and select **Post an Adaptive Card to a Teams user and wait for a response**.</span></span>

13. <span data-ttu-id="f72b7-147">[**受信者**] フィールドに、動的コンテンツから **電子メール** を追加します (選択内容には、そのコンテンツによってExcelロゴが表示されます)。</span><span class="sxs-lookup"><span data-stu-id="f72b7-147">For the **Recipient** field, add **email** from the dynamic content (the selection will have the Excel logo by it).</span></span> <span data-ttu-id="f72b7-148">**電子メール** を追加すると、フロー ステップは **各ブロックに適用** によって囲まれます。</span><span class="sxs-lookup"><span data-stu-id="f72b7-148">Adding **email** causes the flow step to be surrounded by an **Apply to each** block.</span></span> <span data-ttu-id="f72b7-149">つまり、配列はPower Automateによって反復処理されます。</span><span class="sxs-lookup"><span data-stu-id="f72b7-149">That means the array will be iterated over by Power Automate.</span></span>

14. <span data-ttu-id="f72b7-150">アダプティブ カードを送信するには、カードの JSON を **メッセージ** として指定する必要があります。</span><span class="sxs-lookup"><span data-stu-id="f72b7-150">Sending an Adaptive Card requires the card's JSON to be provided as the **Message**.</span></span> <span data-ttu-id="f72b7-151">[アダプティブ カード デザイナー](https://adaptivecards.io/designer/)を使用して、カスタム カードを作成できます。</span><span class="sxs-lookup"><span data-stu-id="f72b7-151">You can use the [Adaptive Card Designer](https://adaptivecards.io/designer/) to create custom cards.</span></span> <span data-ttu-id="f72b7-152">このサンプルでは、次の JSON を使用します。</span><span class="sxs-lookup"><span data-stu-id="f72b7-152">For this sample, use the following JSON.</span></span>  

    ```json
    {
      "$schema": "http://adaptivecards.io/schemas/adaptive-card.json",
      "type": "AdaptiveCard",
      "version": "1.0",
      "body": [
        {
          "type": "TextBlock",
          "size": "Medium",
          "weight": "Bolder",
          "text": "Update your Status Report"
        },
        {
          "type": "Image",
          "altText": "",
          "url": "https://i.imgur.com/f5RcuF3.png"
        },
        {
          "type": "TextBlock",
          "text": "This is a reminder to update your status report for this month's review. You can do so right here in this card, or by adding it directly to the spreadsheet.",
          "wrap": true
        },
        {
          "type": "Input.Text",
          "placeholder": "My status report for this month is...",
          "id": "response",
          "isMultiline": true
        }
      ],
      "actions": [
        {
          "type": "Action.Submit",
          "title": "Submit",
          "id": "submit"
        }
      ]
    }
    ```

15. <span data-ttu-id="f72b7-153">残りのフィールドに次のように入力します。</span><span class="sxs-lookup"><span data-stu-id="f72b7-153">Fill out the remaining fields as follows:</span></span>

    - <span data-ttu-id="f72b7-154">**更新メッセージ**: 進捗レポートを提出していただきありがとうございます。</span><span class="sxs-lookup"><span data-stu-id="f72b7-154">**Update message**: Thank you for submitting your status report.</span></span> <span data-ttu-id="f72b7-155">スプレッドシートに返信が追加されました。</span><span class="sxs-lookup"><span data-stu-id="f72b7-155">Your response has been successfully added to the spreadsheet.</span></span>
    - <span data-ttu-id="f72b7-156">**カードを更新する必要があります**: はい</span><span class="sxs-lookup"><span data-stu-id="f72b7-156">**Should update card**: Yes</span></span>

16. <span data-ttu-id="f72b7-157">**[各ブロックに適用]** で、[**アダプティブ カードをTeamsユーザーにポストして応答を待つ**] に続いて、[**アクションの追加]** をクリックします。</span><span class="sxs-lookup"><span data-stu-id="f72b7-157">In the **Apply to each** block, following the **Post an Adaptive Card to a Teams user and wait for a response**, press **Add an action**.</span></span> <span data-ttu-id="f72b7-158">[**オンライン (ビジネス)] Excel** 選択します。</span><span class="sxs-lookup"><span data-stu-id="f72b7-158">Select **Excel Online (Business)**.</span></span> <span data-ttu-id="f72b7-159">[ **アクション]** で、[ **スクリプトの実行**] を選択します。</span><span class="sxs-lookup"><span data-stu-id="f72b7-159">Under **Actions**, select **Run script**.</span></span> <span data-ttu-id="f72b7-160">フローステップに対して以下のエントリを入力します。</span><span class="sxs-lookup"><span data-stu-id="f72b7-160">Provide the following entries for the flow step:</span></span>

    - <span data-ttu-id="f72b7-161">**場所**: OneDrive for Business</span><span class="sxs-lookup"><span data-stu-id="f72b7-161">**Location**: OneDrive for Business</span></span>
    - <span data-ttu-id="f72b7-162">**ドキュメント ライブラリ**: OneDrive</span><span class="sxs-lookup"><span data-stu-id="f72b7-162">**Document Library**: OneDrive</span></span>
    - <span data-ttu-id="f72b7-163">**ファイル**: task-reminders.xlsx *(ファイルブラウザを使用して選択)*</span><span class="sxs-lookup"><span data-stu-id="f72b7-163">**File**: task-reminders.xlsx *(Chosen through the file browser)*</span></span>
    - <span data-ttu-id="f72b7-164">**スクリプト**: 保存ステータス</span><span class="sxs-lookup"><span data-stu-id="f72b7-164">**Script**: Save Status</span></span>
    - <span data-ttu-id="f72b7-165">**送信者電子メール**: 電子メール *(Excelからの動的コンテンツ)*</span><span class="sxs-lookup"><span data-stu-id="f72b7-165">**senderEmail**: email *(dynamic content from Excel)*</span></span>
    - <span data-ttu-id="f72b7-166">**ステータスレポート応答**: 応答 *(Teamsからの動的コンテンツ)*</span><span class="sxs-lookup"><span data-stu-id="f72b7-166">**statusReportResponse**: response *(dynamic content from Teams)*</span></span>

    :::image type="content" source="../../images/scenario-task-reminders-last-flow-step.png" alt-text="適用先のステップを示すPower Automateフロー":::

17. <span data-ttu-id="f72b7-168">フローを保存します。</span><span class="sxs-lookup"><span data-stu-id="f72b7-168">Save the flow.</span></span>

## <a name="running-the-flow"></a><span data-ttu-id="f72b7-169">フローの実行</span><span class="sxs-lookup"><span data-stu-id="f72b7-169">Running the flow</span></span>

<span data-ttu-id="f72b7-170">フローをテストするには、空白のステータスを持つテーブル行で、Teamsアカウントに関連付けられた電子メール アドレスが使用されていることを確認します (テスト中は、独自の電子メール アドレスを使用する必要があります)。</span><span class="sxs-lookup"><span data-stu-id="f72b7-170">To test the flow, make sure any table rows with blank status use an email address tied to a Teams account (you should probably use your own email address while testing).</span></span>

<span data-ttu-id="f72b7-171">フロー デザイナーから **[テスト** ] を選択するか、[ **自分のフロー** ] ページからフローを実行できます。</span><span class="sxs-lookup"><span data-stu-id="f72b7-171">You can either select **Test** from the flow designer, or run the flow from the **My flows** page.</span></span> <span data-ttu-id="f72b7-172">フローを開始し、必要な接続の使用を受け入れた後、Power AutomateからTeamsまでのアダプティブ カードを受け取る必要があります。</span><span class="sxs-lookup"><span data-stu-id="f72b7-172">After starting the flow and accepting the use of the required connections, you should receive an Adaptive Card from Power Automate through Teams.</span></span> <span data-ttu-id="f72b7-173">カードのステータスフィールドに入力すると、フローが続行され、提供したステータスでスプレッドシートが更新されます。</span><span class="sxs-lookup"><span data-stu-id="f72b7-173">Once you fill out the status field in the card, the flow will continue and update the spreadsheet with the status you provide.</span></span>

### <a name="before-running-the-flow"></a><span data-ttu-id="f72b7-174">フローを実行する前に</span><span class="sxs-lookup"><span data-stu-id="f72b7-174">Before running the flow</span></span>

:::image type="content" source="../../images/scenario-task-reminders-spreadsheet-before.png" alt-text="1 つの不足している状態エントリを含む進捗レポートを含むワークシート":::

### <a name="receiving-the-adaptive-card"></a><span data-ttu-id="f72b7-176">アダプティブカードの受け取り</span><span class="sxs-lookup"><span data-stu-id="f72b7-176">Receiving the Adaptive Card</span></span>

:::image type="content" source="../../images/scenario-task-reminders-adaptive-card.png" alt-text="従業員にステータスの更新を求めるTeamsのアダプティブ カード":::

### <a name="after-running-the-flow"></a><span data-ttu-id="f72b7-178">フローの実行後</span><span class="sxs-lookup"><span data-stu-id="f72b7-178">After running the flow</span></span>

:::image type="content" source="../../images/scenario-task-reminders-spreadsheet-after.png" alt-text="現在入力された状態エントリを含む進捗レポートを含むワークシート":::
