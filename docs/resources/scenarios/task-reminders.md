---
title: 'Office スクリプトのサンプルシナリオ: 自動化されたタスクの事前通知'
description: Power オートメーションとアダプティブカードを使用するサンプルは、プロジェクト管理スプレッドシートでタスクリマインダーを自動化します。
ms.date: 06/09/2020
localization_priority: Normal
ms.openlocfilehash: f764c37dafdd964e9435d504770d10b1608428b8
ms.sourcegitcommit: aec3c971c6640429f89b6bb99d2c95ea06725599
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 06/25/2020
ms.locfileid: "44878907"
---
# <a name="office-scripts-sample-scenario-automated-task-reminders"></a><span data-ttu-id="c98d0-103">Office スクリプトのサンプルシナリオ: 自動化されたタスクの事前通知</span><span class="sxs-lookup"><span data-stu-id="c98d0-103">Office Scripts sample scenario: Automated task reminders</span></span>

<span data-ttu-id="c98d0-104">このシナリオでは、プロジェクトを管理しています。</span><span class="sxs-lookup"><span data-stu-id="c98d0-104">In this scenario you're managing a project.</span></span> <span data-ttu-id="c98d0-105">毎月、従業員の状態を追跡するには、Excel ワークシートを使用します。</span><span class="sxs-lookup"><span data-stu-id="c98d0-105">You use an Excel worksheet to track your employees' status every month.</span></span> <span data-ttu-id="c98d0-106">そのような場合、アラーム処理を自動化することを決定したので、ユーザーに状態を記入するように通知する必要があります。</span><span class="sxs-lookup"><span data-stu-id="c98d0-106">You often need to remind people to fill out their status, so you've decided to automate that reminder process.</span></span>

<span data-ttu-id="c98d0-107">状態フィールドが不足しているメッセージユーザーへの電力自動化フローを作成し、その応答をスプレッドシートに適用します。</span><span class="sxs-lookup"><span data-stu-id="c98d0-107">You'll create a Power Automate flow to message people with missing status fields and apply their responses to the spreadsheet.</span></span> <span data-ttu-id="c98d0-108">これを行うには、ブックの操作を処理するためのスクリプトを作成します。</span><span class="sxs-lookup"><span data-stu-id="c98d0-108">To do this, you'll develop a pair of scripts to handle the working with the workbook.</span></span> <span data-ttu-id="c98d0-109">最初のスクリプトは、空の状態のユーザーのリストを取得し、2番目のスクリプトは、右側の行にステータス文字列を追加します。</span><span class="sxs-lookup"><span data-stu-id="c98d0-109">The first script gets a list of people with blank statuses and the second script adds a status string to the right row.</span></span> <span data-ttu-id="c98d0-110">また、 [Teams のアダプティブカード](/microsoftteams/platform/task-modules-and-cards/what-are-cards) を使用して、従業員が通知から直接状態を入力できるようにします。</span><span class="sxs-lookup"><span data-stu-id="c98d0-110">You'll also make use of [Teams Adaptive Cards](/microsoftteams/platform/task-modules-and-cards/what-are-cards) to have employees enter their status directly from the notification.</span></span>

## <a name="scripting-skills-covered"></a><span data-ttu-id="c98d0-111">スクリプト作成スキルの説明</span><span class="sxs-lookup"><span data-stu-id="c98d0-111">Scripting skills covered</span></span>

- <span data-ttu-id="c98d0-112">パワー自動化でフローを作成する</span><span class="sxs-lookup"><span data-stu-id="c98d0-112">Create flows in Power Automate</span></span>
- <span data-ttu-id="c98d0-113">スクリプトにデータを渡す</span><span class="sxs-lookup"><span data-stu-id="c98d0-113">Pass data to scripts</span></span>
- <span data-ttu-id="c98d0-114">スクリプトからデータを返す</span><span class="sxs-lookup"><span data-stu-id="c98d0-114">Return data from scripts</span></span>
- <span data-ttu-id="c98d0-115">Teams のアダプティブカード</span><span class="sxs-lookup"><span data-stu-id="c98d0-115">Teams Adaptive Cards</span></span>
- <span data-ttu-id="c98d0-116">テーブル</span><span class="sxs-lookup"><span data-stu-id="c98d0-116">Tables</span></span>

## <a name="prerequisites"></a><span data-ttu-id="c98d0-117">前提条件</span><span class="sxs-lookup"><span data-stu-id="c98d0-117">Prerequisites</span></span>

<span data-ttu-id="c98d0-118">このシナリオでは、 [Power オートメーション](https://flow.microsoft.com) と [Microsoft Teams](https://www.microsoft.com/microsoft-365/microsoft-teams/group-chat-software)を使用します。</span><span class="sxs-lookup"><span data-stu-id="c98d0-118">This scenario uses [Power Automate](https://flow.microsoft.com) and [Microsoft Teams](https://www.microsoft.com/microsoft-365/microsoft-teams/group-chat-software).</span></span> <span data-ttu-id="c98d0-119">Office スクリプトの開発に使用するアカウントに両方が関連付けられている必要があります。</span><span class="sxs-lookup"><span data-stu-id="c98d0-119">You will need both associated with the account that you use for developing Office Scripts.</span></span> <span data-ttu-id="c98d0-120">Microsoft 開発者向けサブスクリプションに無料でアクセスし、これらのアプリケーションについて学習して作業するには、 [microsoft 365 開発者プログラム](https://developer.microsoft.com/microsoft-365/dev-program)に参加することを検討してください。</span><span class="sxs-lookup"><span data-stu-id="c98d0-120">For free access to a Microsoft Developer subscription to learn about and work with these applications, consider joining the [Microsoft 365 Developer Program](https://developer.microsoft.com/microsoft-365/dev-program).</span></span>

## <a name="setup-instructions"></a><span data-ttu-id="c98d0-121">セットアップの手順</span><span class="sxs-lookup"><span data-stu-id="c98d0-121">Setup instructions</span></span>

1. <span data-ttu-id="c98d0-122">OneDrive に <a href="task-reminders.xlsx">task-reminders.xlsx</a> をダウンロードします。</span><span class="sxs-lookup"><span data-stu-id="c98d0-122">Download <a href="task-reminders.xlsx">task-reminders.xlsx</a> to your OneDrive.</span></span>

2. <span data-ttu-id="c98d0-123">Web 上の Excel でブックを開きます。</span><span class="sxs-lookup"><span data-stu-id="c98d0-123">Open the workbook in Excel on the web.</span></span>

3. <span data-ttu-id="c98d0-124">[ **自動化** ] タブで、 **コードエディター**を開きます。</span><span class="sxs-lookup"><span data-stu-id="c98d0-124">Under the **Automate** tab, open the **Code Editor**.</span></span>

4. <span data-ttu-id="c98d0-125">最初に、すべての従業員に対して、スプレッドシートから不足している状態レポートを取得するためのスクリプトが必要です。</span><span class="sxs-lookup"><span data-stu-id="c98d0-125">First, we need a script to get all the employees with status reports that are missing from the spreadsheet.</span></span> <span data-ttu-id="c98d0-126">[ **コードエディター** ] 作業ウィンドウで、[ **新しいスクリプト** ] をクリックし、次のスクリプトをエディターに貼り付けます。</span><span class="sxs-lookup"><span data-stu-id="c98d0-126">In the **Code Editor** task pane, press **New Script** and paste the following script into the editor.</span></span>

    ```typescript
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
          people.push({ name: row[NAME_INDEX], email: row[EMAIL_INDEX] });
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

5. <span data-ttu-id="c98d0-127">「 **Get People**」という名前のスクリプトを保存します。</span><span class="sxs-lookup"><span data-stu-id="c98d0-127">Save the script with the name **Get People**.</span></span>

6. <span data-ttu-id="c98d0-128">次に、進捗レポートカードを処理し、新しい情報をスプレッドシートに格納するための2番目のスクリプトが必要です。</span><span class="sxs-lookup"><span data-stu-id="c98d0-128">Next, we need a second script to process the status report cards and put the new information in the spreadsheet.</span></span> <span data-ttu-id="c98d0-129">[ **コードエディター** ] 作業ウィンドウで、[ **新しいスクリプト** ] をクリックし、次のスクリプトをエディターに貼り付けます。</span><span class="sxs-lookup"><span data-stu-id="c98d0-129">In the **Code Editor** task pane, press **New Script** and paste the following script into the editor.</span></span>

    ```typescript
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

7. <span data-ttu-id="c98d0-130">**Save Status**という名前でスクリプトを保存します。</span><span class="sxs-lookup"><span data-stu-id="c98d0-130">Save the script with the name **Save Status**.</span></span>

8. <span data-ttu-id="c98d0-131">次に、フローを作成する必要があります。</span><span class="sxs-lookup"><span data-stu-id="c98d0-131">Now, we need to create the flow.</span></span> <span data-ttu-id="c98d0-132">[電源自動化](https://flow.microsoft.com/)を開きます。</span><span class="sxs-lookup"><span data-stu-id="c98d0-132">Open [Power Automate](https://flow.microsoft.com/).</span></span>

    > [!TIP]
    > <span data-ttu-id="c98d0-133">以前にフローを作成していない場合は、チュートリアル「 [Power オートメーションを使用したスクリプトの使用を開始](../../tutorials/excel-power-automate-manual.md) する」を参照して、基本事項を確認してください。</span><span class="sxs-lookup"><span data-stu-id="c98d0-133">If you haven't created a flow before, please check out our tutorial [Start using scripts with Power Automate](../../tutorials/excel-power-automate-manual.md) to learn the basics.</span></span>

9. <span data-ttu-id="c98d0-134">新しい **インスタントフロー**を作成します。</span><span class="sxs-lookup"><span data-stu-id="c98d0-134">Create a new **Instant flow**.</span></span>

10. <span data-ttu-id="c98d0-135">[オプションから **フローを手動でトリガーする** ] を選択し、[ **作成**] をクリックします。</span><span class="sxs-lookup"><span data-stu-id="c98d0-135">Choose **Manually trigger a flow** from the options and press **Create**.</span></span>

11. <span data-ttu-id="c98d0-136">このフローでは、すべての従業員に空の状態フィールドを取得するために、 **Get People** スクリプトを呼び出す必要があります。</span><span class="sxs-lookup"><span data-stu-id="c98d0-136">The flow needs to call the **Get People** script to get all the employees with empty status fields.</span></span> <span data-ttu-id="c98d0-137">[ **新しい手順** ] をクリックし、[ **Excel Online (Business)**] を選択します。</span><span class="sxs-lookup"><span data-stu-id="c98d0-137">Press **New step** and select **Excel Online (Business)**.</span></span> <span data-ttu-id="c98d0-138">**[アクション]** の下から、**[スクリプトの実行 (プレビュー)]** を選択します。</span><span class="sxs-lookup"><span data-stu-id="c98d0-138">Under **Actions**, select **Run script (preview)**.</span></span> <span data-ttu-id="c98d0-139">フローステップに対して次のエントリを指定します。</span><span class="sxs-lookup"><span data-stu-id="c98d0-139">Provide the following entries for the flow step:</span></span>

    - <span data-ttu-id="c98d0-140">**場所**: OneDrive for Business</span><span class="sxs-lookup"><span data-stu-id="c98d0-140">**Location**: OneDrive for Business</span></span>
    - <span data-ttu-id="c98d0-141">**ドキュメント ライブラリ**: OneDrive</span><span class="sxs-lookup"><span data-stu-id="c98d0-141">**Document Library**: OneDrive</span></span>
    - <span data-ttu-id="c98d0-142">**ファイル**: task-reminders.xlsx</span><span class="sxs-lookup"><span data-stu-id="c98d0-142">**File**: task-reminders.xlsx</span></span>
    - <span data-ttu-id="c98d0-143">**スクリプト**: ユーザーを取得する</span><span class="sxs-lookup"><span data-stu-id="c98d0-143">**Script**: Get People</span></span>

    ![最初に実行するスクリプトフローステップ。](../../images/scenario-task-reminders-first-flow-step.png)

12. <span data-ttu-id="c98d0-145">次に、このフローは、スクリプトから返される配列内の各従業員を処理する必要があります。</span><span class="sxs-lookup"><span data-stu-id="c98d0-145">Next, the flow needs to process each Employee in the array returned by the script.</span></span> <span data-ttu-id="c98d0-146">[ **新規作成** ] をクリックして、[ **Teams ユーザーにアダプティブカードを送信**し、応答を待つ] を選択します。</span><span class="sxs-lookup"><span data-stu-id="c98d0-146">Press **New step** and select **Post an Adaptive Card to a Teams user and wait for a response**.</span></span>

13. <span data-ttu-id="c98d0-147">[ **受信者** ] フィールドでは、動的コンテンツから **電子メール** を追加します (選択範囲には Excel ロゴが表示されます)。</span><span class="sxs-lookup"><span data-stu-id="c98d0-147">For the **Recipient** field, add **email** from the dynamic content (the selection will have the Excel logo by it).</span></span> <span data-ttu-id="c98d0-148">**電子メール**を追加すると、**各ブロックに適用**されるフローステップが囲まれます。</span><span class="sxs-lookup"><span data-stu-id="c98d0-148">Adding **email** causes the flow step to be surrounded by an **Apply to each** block.</span></span> <span data-ttu-id="c98d0-149">これは、アレイが電力自動化によって反復処理されることを意味します。</span><span class="sxs-lookup"><span data-stu-id="c98d0-149">That means the array will be iterated over by Power Automate.</span></span>

14. <span data-ttu-id="c98d0-150">アダプティブカードを送信するには、 **メッセージ**としてカードの JSON を提供する必要があります。</span><span class="sxs-lookup"><span data-stu-id="c98d0-150">Sending an Adaptive Card requires the card's JSON to be provided as the **Message**.</span></span> <span data-ttu-id="c98d0-151">[アダプティブカードデザイナー](https://adaptivecards.io/designer/)を使用して、カスタムカードを作成できます。</span><span class="sxs-lookup"><span data-stu-id="c98d0-151">You can use the [Adaptive Card Designer](https://adaptivecards.io/designer/) to create custom cards.</span></span> <span data-ttu-id="c98d0-152">この例では、次の JSON を使用します。</span><span class="sxs-lookup"><span data-stu-id="c98d0-152">For this sample, use the following JSON.</span></span>  

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

15. <span data-ttu-id="c98d0-153">残りのフィールドに、次のように入力します。</span><span class="sxs-lookup"><span data-stu-id="c98d0-153">Fill out the remaining fields as follows:</span></span>

    - <span data-ttu-id="c98d0-154">**メッセージの更新**: 進捗レポートを提出していただきありがとうございます。</span><span class="sxs-lookup"><span data-stu-id="c98d0-154">**Update message**: Thank you for submitting your status report.</span></span> <span data-ttu-id="c98d0-155">応答が正常にスプレッドシートに追加されました。</span><span class="sxs-lookup"><span data-stu-id="c98d0-155">Your response has been successfully added to the spreadsheet.</span></span>
    - <span data-ttu-id="c98d0-156">**カードを更新する必要があり**ます。はい</span><span class="sxs-lookup"><span data-stu-id="c98d0-156">**Should update card**: Yes</span></span>

16. <span data-ttu-id="c98d0-157">[ **各ブロックに適用** ] で、[ **Teams ユーザーにアダプティブカードを投稿**し、応答を待機する] の下にある [アクションの **追加**] をクリックします。</span><span class="sxs-lookup"><span data-stu-id="c98d0-157">In the **Apply to each** block, following the **Post an Adaptive Card to a Teams user and wait for a response**, press **Add an action**.</span></span> <span data-ttu-id="c98d0-158">[ **Excel Online (Business)**] を選択します。</span><span class="sxs-lookup"><span data-stu-id="c98d0-158">Select **Excel Online (Business)**.</span></span> <span data-ttu-id="c98d0-159">**[アクション]** の下から、**[スクリプトの実行 (プレビュー)]** を選択します。</span><span class="sxs-lookup"><span data-stu-id="c98d0-159">Under **Actions**, select **Run script (preview)**.</span></span> <span data-ttu-id="c98d0-160">フローステップに対して次のエントリを指定します。</span><span class="sxs-lookup"><span data-stu-id="c98d0-160">Provide the following entries for the flow step:</span></span>

    - <span data-ttu-id="c98d0-161">**場所**: OneDrive for Business</span><span class="sxs-lookup"><span data-stu-id="c98d0-161">**Location**: OneDrive for Business</span></span>
    - <span data-ttu-id="c98d0-162">**ドキュメント ライブラリ**: OneDrive</span><span class="sxs-lookup"><span data-stu-id="c98d0-162">**Document Library**: OneDrive</span></span>
    - <span data-ttu-id="c98d0-163">**ファイル**: task-reminders.xlsx</span><span class="sxs-lookup"><span data-stu-id="c98d0-163">**File**: task-reminders.xlsx</span></span>
    - <span data-ttu-id="c98d0-164">**スクリプト**: 状態の保存</span><span class="sxs-lookup"><span data-stu-id="c98d0-164">**Script**: Save Status</span></span>
    - <span data-ttu-id="c98d0-165">**senderEmail**: 電子メール *(Excel の動的コンテンツ)*</span><span class="sxs-lookup"><span data-stu-id="c98d0-165">**senderEmail**: email *(dynamic content from Excel)*</span></span>
    - <span data-ttu-id="c98d0-166">**Statusreportresponse**: Response *(Teams からの動的コンテンツ)*</span><span class="sxs-lookup"><span data-stu-id="c98d0-166">**statusReportResponse**: response *(dynamic content from Teams)*</span></span>

    ![各フローステップに適用されます。](../../images/scenario-task-reminders-last-flow-step.png)

17. <span data-ttu-id="c98d0-168">フローを保存します。</span><span class="sxs-lookup"><span data-stu-id="c98d0-168">Save the flow.</span></span>

## <a name="running-the-flow"></a><span data-ttu-id="c98d0-169">フローの実行</span><span class="sxs-lookup"><span data-stu-id="c98d0-169">Running the flow</span></span>

<span data-ttu-id="c98d0-170">フローをテストするには、空の状態の表の行が Teams アカウントに関連付けられた電子メールアドレスを使用していることを確認してください (テスト中は、自分の電子メールアドレスを使用する必要があります)。</span><span class="sxs-lookup"><span data-stu-id="c98d0-170">To test the flow, make sure any table rows with blank status use an email address tied to a Teams account (you should probably use your own email address while testing).</span></span>

<span data-ttu-id="c98d0-171">フローデザイナーで [ **テスト** ] を選択するか、[ **自分** のフロー] ページからフローを実行することができます。</span><span class="sxs-lookup"><span data-stu-id="c98d0-171">You can either select **Test** from the flow designer, or run the flow from the **My flows** page.</span></span> <span data-ttu-id="c98d0-172">フローを開始し、必要な接続の使用を承諾した後、Teams を通じて省電力処理カードを受信する必要があります。</span><span class="sxs-lookup"><span data-stu-id="c98d0-172">After starting the flow and accepting the use of the required connections, you should receive an Adaptive Card from Power Automate through Teams.</span></span> <span data-ttu-id="c98d0-173">カードで [状態] フィールドに入力すると、フローは続行され、指定した状態でスプレッドシートが更新されます。</span><span class="sxs-lookup"><span data-stu-id="c98d0-173">Once you fill out the status field in the card, the flow will continue and update the spreadsheet with the status you provide.</span></span>

### <a name="before-running-the-flow"></a><span data-ttu-id="c98d0-174">フローを実行する前に</span><span class="sxs-lookup"><span data-stu-id="c98d0-174">Before running the flow</span></span>

![進捗レポートを含むワークシートに、不足しているステータスエントリが1つ含まれています。](../../images/scenario-task-reminders-spreadsheet-before.png)

### <a name="receiving-the-adaptive-card"></a><span data-ttu-id="c98d0-176">アダプティブカードの受信</span><span class="sxs-lookup"><span data-stu-id="c98d0-176">Receiving the Adaptive Card</span></span>

![従業員に進捗の更新を求める、Teams のアダプティブカード。](../../images/scenario-task-reminders-adaptive-card.png)

### <a name="after-running-the-flow"></a><span data-ttu-id="c98d0-178">フローの実行後</span><span class="sxs-lookup"><span data-stu-id="c98d0-178">After running the flow</span></span>

![現在入力されている進捗レポートを含むワークシート。](../../images/scenario-task-reminders-spreadsheet-after.png)
