---
title: 'Officeスクリプトのサンプル シナリオ: タスクの自動アラーム'
description: プロジェクト管理スプレッドシートでPower Automateアダプティブ カードを使用するサンプルは、タスクリマインダーを自動化します。
ms.date: 06/29/2021
localization_priority: Normal
ms.openlocfilehash: cf25b81ad44bbe963083f6a8346c0fd59a514305
ms.sourcegitcommit: 211c157ca746e266eeb079f5fa1925a1e35ab702
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 07/07/2021
ms.locfileid: "53313982"
---
# <a name="office-scripts-sample-scenario-automated-task-reminders"></a><span data-ttu-id="3fbd3-103">Officeスクリプトのサンプル シナリオ: タスクの自動アラーム</span><span class="sxs-lookup"><span data-stu-id="3fbd3-103">Office Scripts sample scenario: Automated task reminders</span></span>

<span data-ttu-id="3fbd3-104">このシナリオでは、プロジェクトを管理しています。</span><span class="sxs-lookup"><span data-stu-id="3fbd3-104">In this scenario you're managing a project.</span></span> <span data-ttu-id="3fbd3-105">毎月従業員のExcelを追跡するには、ユーザーのワークシートを使用します。</span><span class="sxs-lookup"><span data-stu-id="3fbd3-105">You use an Excel worksheet to track your employees' status every month.</span></span> <span data-ttu-id="3fbd3-106">多くの場合、ユーザーに自分の状態を入力することを通知する必要があります。そのため、そのリマインダー プロセスを自動化することを決めました。</span><span class="sxs-lookup"><span data-stu-id="3fbd3-106">You often need to remind people to fill out their status, so you've decided to automate that reminder process.</span></span>

<span data-ttu-id="3fbd3-107">ステータス フィールドが見つからないPower Automateに対するメッセージ フローを作成し、その応答をスプレッドシートに適用します。</span><span class="sxs-lookup"><span data-stu-id="3fbd3-107">You'll create a Power Automate flow to message people with missing status fields and apply their responses to the spreadsheet.</span></span> <span data-ttu-id="3fbd3-108">これを行うには、ブックの操作を処理するためのスクリプトのペアを開発します。</span><span class="sxs-lookup"><span data-stu-id="3fbd3-108">To do this, you'll develop a pair of scripts to handle the working with the workbook.</span></span> <span data-ttu-id="3fbd3-109">最初のスクリプトは、空の状態を持つユーザーの一覧を取得し、2 番目のスクリプトは、右側の行に状態文字列を追加します。</span><span class="sxs-lookup"><span data-stu-id="3fbd3-109">The first script gets a list of people with blank statuses and the second script adds a status string to the right row.</span></span> <span data-ttu-id="3fbd3-110">また、アダプティブ カードをTeams[して、](/microsoftteams/platform/task-modules-and-cards/what-are-cards)従業員に通知から直接ステータスを入力します。</span><span class="sxs-lookup"><span data-stu-id="3fbd3-110">You'll also make use of [Teams Adaptive Cards](/microsoftteams/platform/task-modules-and-cards/what-are-cards) to have employees enter their status directly from the notification.</span></span>

## <a name="scripting-skills-covered"></a><span data-ttu-id="3fbd3-111">スクリプティングのスキルをカバー</span><span class="sxs-lookup"><span data-stu-id="3fbd3-111">Scripting skills covered</span></span>

- <span data-ttu-id="3fbd3-112">[フローの作成] Power Automate</span><span class="sxs-lookup"><span data-stu-id="3fbd3-112">Create flows in Power Automate</span></span>
- <span data-ttu-id="3fbd3-113">スクリプトにデータを渡す</span><span class="sxs-lookup"><span data-stu-id="3fbd3-113">Pass data to scripts</span></span>
- <span data-ttu-id="3fbd3-114">スクリプトからデータを返す</span><span class="sxs-lookup"><span data-stu-id="3fbd3-114">Return data from scripts</span></span>
- <span data-ttu-id="3fbd3-115">Teamsアダプティブ カード</span><span class="sxs-lookup"><span data-stu-id="3fbd3-115">Teams Adaptive Cards</span></span>
- <span data-ttu-id="3fbd3-116">テーブル</span><span class="sxs-lookup"><span data-stu-id="3fbd3-116">Tables</span></span>

## <a name="prerequisites"></a><span data-ttu-id="3fbd3-117">前提条件</span><span class="sxs-lookup"><span data-stu-id="3fbd3-117">Prerequisites</span></span>

<span data-ttu-id="3fbd3-118">このシナリオでは[、Power Automate](https://flow.microsoft.com)とMicrosoft Teamsを[使用します](https://www.microsoft.com/microsoft-365/microsoft-teams/group-chat-software)。</span><span class="sxs-lookup"><span data-stu-id="3fbd3-118">This scenario uses [Power Automate](https://flow.microsoft.com) and [Microsoft Teams](https://www.microsoft.com/microsoft-365/microsoft-teams/group-chat-software).</span></span> <span data-ttu-id="3fbd3-119">両方とも、スクリプトの開発に使用するアカウントに関連付Officeがあります。</span><span class="sxs-lookup"><span data-stu-id="3fbd3-119">You will need both associated with the account that you use for developing Office Scripts.</span></span> <span data-ttu-id="3fbd3-120">Microsoft Developer サブスクリプションに無料でアクセスして、これらのアプリケーションについて学び、これらのアプリケーションを使用するには、開発者プログラムに参加Microsoft 365[検討してください](https://developer.microsoft.com/microsoft-365/dev-program)。</span><span class="sxs-lookup"><span data-stu-id="3fbd3-120">For free access to a Microsoft Developer subscription to learn about and work with these applications, consider joining the [Microsoft 365 Developer Program](https://developer.microsoft.com/microsoft-365/dev-program).</span></span>

## <a name="setup-instructions"></a><span data-ttu-id="3fbd3-121">セットアップ手順</span><span class="sxs-lookup"><span data-stu-id="3fbd3-121">Setup instructions</span></span>

1. <span data-ttu-id="3fbd3-122">ユーザー <a href="task-reminders.xlsx">task-reminders.xlsx</a>にダウンロードOneDrive。</span><span class="sxs-lookup"><span data-stu-id="3fbd3-122">Download <a href="task-reminders.xlsx">task-reminders.xlsx</a> to your OneDrive.</span></span>

1. <span data-ttu-id="3fbd3-123">ブックを開Excel on the web。</span><span class="sxs-lookup"><span data-stu-id="3fbd3-123">Open the workbook in Excel on the web.</span></span>

1. <span data-ttu-id="3fbd3-124">まず、スプレッドシートに不足している状態レポートを持つすべての従業員を取得するスクリプトが必要です。</span><span class="sxs-lookup"><span data-stu-id="3fbd3-124">First, we need a script to get all the employees with status reports that are missing from the spreadsheet.</span></span> <span data-ttu-id="3fbd3-125">[自動化] **タブで** 、[新しい **スクリプト] を選択** し、次のスクリプトをエディターに貼り付けます。</span><span class="sxs-lookup"><span data-stu-id="3fbd3-125">Under the **Automate** tab, select **New Script** and paste the following script into the editor.</span></span>

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

1. <span data-ttu-id="3fbd3-126">[ユーザーの取得] という名前のスクリプト **を保存します**。</span><span class="sxs-lookup"><span data-stu-id="3fbd3-126">Save the script with the name **Get People**.</span></span>

1. <span data-ttu-id="3fbd3-127">次に、ステータス レポート カードを処理し、新しい情報をスプレッドシートに入れる 2 番目のスクリプトが必要です。</span><span class="sxs-lookup"><span data-stu-id="3fbd3-127">Next, we need a second script to process the status report cards and put the new information in the spreadsheet.</span></span> <span data-ttu-id="3fbd3-128">[コード エディター] 作業ウィンドウで、[ **新しいスクリプト** ] を選択し、次のスクリプトをエディターに貼り付けます。</span><span class="sxs-lookup"><span data-stu-id="3fbd3-128">In the Code Editor task pane, select **New Script** and paste the following script into the editor.</span></span>

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

1. <span data-ttu-id="3fbd3-129">[状態の保存] という名前のスクリプト **を保存します**。</span><span class="sxs-lookup"><span data-stu-id="3fbd3-129">Save the script with the name **Save Status**.</span></span>

1. <span data-ttu-id="3fbd3-130">次に、フローを作成する必要があります。</span><span class="sxs-lookup"><span data-stu-id="3fbd3-130">Now, we need to create the flow.</span></span> <span data-ttu-id="3fbd3-131">[ファイル[Power Automate] を開きます](https://flow.microsoft.com/)。</span><span class="sxs-lookup"><span data-stu-id="3fbd3-131">Open [Power Automate](https://flow.microsoft.com/).</span></span>

    > [!TIP]
    > <span data-ttu-id="3fbd3-132">前にフローを作成したことがない場合は、チュートリアル「スクリプトの使用[](../../tutorials/excel-power-automate-manual.md)を開始する」を参照し、Power Automateを確認してください。</span><span class="sxs-lookup"><span data-stu-id="3fbd3-132">If you haven't created a flow before, please check out our tutorial [Start using scripts with Power Automate](../../tutorials/excel-power-automate-manual.md) to learn the basics.</span></span>

1. <span data-ttu-id="3fbd3-133">新しいインスタント フロー **を作成します**。</span><span class="sxs-lookup"><span data-stu-id="3fbd3-133">Create a new **Instant flow**.</span></span>

1. <span data-ttu-id="3fbd3-134">[オプション **からフローを手動でトリガーする** ] を選択し、[作成] を **選択します**。</span><span class="sxs-lookup"><span data-stu-id="3fbd3-134">Choose **Manually trigger a flow** from the options and select **Create**.</span></span>

1. <span data-ttu-id="3fbd3-135">フローは、空の状態フィールドを持つすべての従業員を取得するために Get **People** スクリプトを呼び出す必要があります。</span><span class="sxs-lookup"><span data-stu-id="3fbd3-135">The flow needs to call the **Get People** script to get all the employees with empty status fields.</span></span> <span data-ttu-id="3fbd3-136">[**新しい手順]** を選択し、[オンライン **Excel (Business) を選択します**。</span><span class="sxs-lookup"><span data-stu-id="3fbd3-136">Select **New step**, then select **Excel Online (Business)**.</span></span> <span data-ttu-id="3fbd3-137">**[アクション]** で、**[スクリプトの実行]** を選択します。</span><span class="sxs-lookup"><span data-stu-id="3fbd3-137">Under **Actions**, select **Run script**.</span></span> <span data-ttu-id="3fbd3-138">フロー ステップに次のエントリを指定します。</span><span class="sxs-lookup"><span data-stu-id="3fbd3-138">Provide the following entries for the flow step:</span></span>

    - <span data-ttu-id="3fbd3-139">**場所**: OneDrive for Business</span><span class="sxs-lookup"><span data-stu-id="3fbd3-139">**Location**: OneDrive for Business</span></span>
    - <span data-ttu-id="3fbd3-140">**ドキュメント ライブラリ**: OneDrive</span><span class="sxs-lookup"><span data-stu-id="3fbd3-140">**Document Library**: OneDrive</span></span>
    - <span data-ttu-id="3fbd3-141">**ファイル**: task-reminders.xlsx *(ファイル ブラウザーから選択)*</span><span class="sxs-lookup"><span data-stu-id="3fbd3-141">**File**: task-reminders.xlsx *(Chosen through the file browser)*</span></span>
    - <span data-ttu-id="3fbd3-142">**スクリプト**: ユーザーを取得する</span><span class="sxs-lookup"><span data-stu-id="3fbd3-142">**Script**: Get People</span></span>

    :::image type="content" source="../../images/scenario-task-reminders-first-flow-step.png" alt-text="最初Power Automateスクリプト フロー の手順を示す手順を示す手順を示します。":::

1. <span data-ttu-id="3fbd3-144">次に、フローは、スクリプトによって返される配列内の各 Employee を処理する必要があります。</span><span class="sxs-lookup"><span data-stu-id="3fbd3-144">Next, the flow needs to process each Employee in the array returned by the script.</span></span> <span data-ttu-id="3fbd3-145">[**新しい手順]** を選択し、[アダプティブ カードをユーザーに投稿Teams **応答を待ちます**。</span><span class="sxs-lookup"><span data-stu-id="3fbd3-145">Select **New step**, then choose **Post an Adaptive Card to a Teams user and wait for a response**.</span></span>

1. <span data-ttu-id="3fbd3-146">[受信者 **] フィールド** に動的 **コンテンツから電子** メールを追加します (選択すると、Excelロゴが表示されます)。</span><span class="sxs-lookup"><span data-stu-id="3fbd3-146">For the **Recipient** field, add **email** from the dynamic content (the selection will have the Excel logo by it).</span></span> <span data-ttu-id="3fbd3-147">メール **を** 追加すると、フロー ステップは各ブロックに **適用されます** 。</span><span class="sxs-lookup"><span data-stu-id="3fbd3-147">Adding **email** causes the flow step to be surrounded by an **Apply to each** block.</span></span> <span data-ttu-id="3fbd3-148">つまり、配列は配列によって反復処理Power Automate。</span><span class="sxs-lookup"><span data-stu-id="3fbd3-148">That means the array will be iterated over by Power Automate.</span></span>

1. <span data-ttu-id="3fbd3-149">アダプティブ カードを送信するには、カードの JSON をメッセージとして提供する必要 **があります**。</span><span class="sxs-lookup"><span data-stu-id="3fbd3-149">Sending an Adaptive Card requires the card's JSON to be provided as the **Message**.</span></span> <span data-ttu-id="3fbd3-150">アダプティブ カード デザイナーを [使用してカスタム](https://adaptivecards.io/designer/) カードを作成できます。</span><span class="sxs-lookup"><span data-stu-id="3fbd3-150">You can use the [Adaptive Card Designer](https://adaptivecards.io/designer/) to create custom cards.</span></span> <span data-ttu-id="3fbd3-151">このサンプルでは、次の JSON を使用します。</span><span class="sxs-lookup"><span data-stu-id="3fbd3-151">For this sample, use the following JSON.</span></span>  

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

1. <span data-ttu-id="3fbd3-152">残りのフィールドに次のように入力します。</span><span class="sxs-lookup"><span data-stu-id="3fbd3-152">Fill out the remaining fields as follows:</span></span>

    - <span data-ttu-id="3fbd3-153">**更新メッセージ**: ステータス レポートを提出してありがとうございます。</span><span class="sxs-lookup"><span data-stu-id="3fbd3-153">**Update message**: Thank you for submitting your status report.</span></span> <span data-ttu-id="3fbd3-154">応答がスプレッドシートに正常に追加されました。</span><span class="sxs-lookup"><span data-stu-id="3fbd3-154">Your response has been successfully added to the spreadsheet.</span></span>
    - <span data-ttu-id="3fbd3-155">**カードを更新する必要があります**: はい</span><span class="sxs-lookup"><span data-stu-id="3fbd3-155">**Should update card**: Yes</span></span>

1. <span data-ttu-id="3fbd3-156">[各 **ブロックに適用]** で、[アダプティブ カードをユーザーに投稿Teams応答を待つ] の後、[アクションの追加 **] を選択します**。</span><span class="sxs-lookup"><span data-stu-id="3fbd3-156">In the **Apply to each** block, following the **Post an Adaptive Card to a Teams user and wait for a response**, select **Add an action**.</span></span> <span data-ttu-id="3fbd3-157">[オンライン **Excel (Business) を選択します**。</span><span class="sxs-lookup"><span data-stu-id="3fbd3-157">Select **Excel Online (Business)**.</span></span> <span data-ttu-id="3fbd3-158">**[アクション]** で、**[スクリプトの実行]** を選択します。</span><span class="sxs-lookup"><span data-stu-id="3fbd3-158">Under **Actions**, select **Run script**.</span></span> <span data-ttu-id="3fbd3-159">フロー ステップに次のエントリを指定します。</span><span class="sxs-lookup"><span data-stu-id="3fbd3-159">Provide the following entries for the flow step:</span></span>

    - <span data-ttu-id="3fbd3-160">**場所**: OneDrive for Business</span><span class="sxs-lookup"><span data-stu-id="3fbd3-160">**Location**: OneDrive for Business</span></span>
    - <span data-ttu-id="3fbd3-161">**ドキュメント ライブラリ**: OneDrive</span><span class="sxs-lookup"><span data-stu-id="3fbd3-161">**Document Library**: OneDrive</span></span>
    - <span data-ttu-id="3fbd3-162">**ファイル**: task-reminders.xlsx *(ファイル ブラウザーから選択)*</span><span class="sxs-lookup"><span data-stu-id="3fbd3-162">**File**: task-reminders.xlsx *(Chosen through the file browser)*</span></span>
    - <span data-ttu-id="3fbd3-163">**スクリプト**: 状態の保存</span><span class="sxs-lookup"><span data-stu-id="3fbd3-163">**Script**: Save Status</span></span>
    - <span data-ttu-id="3fbd3-164">**senderEmail**: メール *(メールからの動的Excel)*</span><span class="sxs-lookup"><span data-stu-id="3fbd3-164">**senderEmail**: email *(dynamic content from Excel)*</span></span>
    - <span data-ttu-id="3fbd3-165">**statusReportResponse**: 応答 *(Teams)*</span><span class="sxs-lookup"><span data-stu-id="3fbd3-165">**statusReportResponse**: response *(dynamic content from Teams)*</span></span>

    :::image type="content" source="../../images/scenario-task-reminders-last-flow-step.png" alt-text="各Power Automate適用を示すフローを示す手順を示します。":::

1. <span data-ttu-id="3fbd3-167">フローを保存します。</span><span class="sxs-lookup"><span data-stu-id="3fbd3-167">Save the flow.</span></span>

## <a name="running-the-flow"></a><span data-ttu-id="3fbd3-168">フローの実行</span><span class="sxs-lookup"><span data-stu-id="3fbd3-168">Running the flow</span></span>

<span data-ttu-id="3fbd3-169">フローをテストするには、状態が空白のテーブル行で Teams アカウントに関連付けられている電子メール アドレスを使用します (テスト中は、独自の電子メール アドレスを使用する必要があります)。</span><span class="sxs-lookup"><span data-stu-id="3fbd3-169">To test the flow, make sure any table rows with blank status use an email address tied to a Teams account (you should probably use your own email address while testing).</span></span> <span data-ttu-id="3fbd3-170">[フロー エディター **] ページ** の [テスト] ボタンを使用するか、[マイ フロー] タブでフロー **を実行** します。メッセージが表示されたら、必ずアクセスを許可してください。</span><span class="sxs-lookup"><span data-stu-id="3fbd3-170">Use the **Test** button on the flow editor page or run the flow through your **My flows** tab. Be sure to allow access when prompted.</span></span>

<span data-ttu-id="3fbd3-171">アダプティブ カードを受け取る必要Power AutomateからTeams。</span><span class="sxs-lookup"><span data-stu-id="3fbd3-171">You should receive an Adaptive Card from Power Automate through Teams.</span></span> <span data-ttu-id="3fbd3-172">カードの状態フィールドに入力すると、フローは続行され、指定した状態でスプレッドシートが更新されます。</span><span class="sxs-lookup"><span data-stu-id="3fbd3-172">Once you fill out the status field in the card, the flow will continue and update the spreadsheet with the status you provide.</span></span>

### <a name="before-running-the-flow"></a><span data-ttu-id="3fbd3-173">フローを実行する前に</span><span class="sxs-lookup"><span data-stu-id="3fbd3-173">Before running the flow</span></span>

:::image type="content" source="../../images/scenario-task-reminders-spreadsheet-before.png" alt-text="不足している状態エントリが 1 つ含まれる状態レポートを含むワークシート。":::

### <a name="receiving-the-adaptive-card"></a><span data-ttu-id="3fbd3-175">アダプティブ カードの受信</span><span class="sxs-lookup"><span data-stu-id="3fbd3-175">Receiving the Adaptive Card</span></span>

:::image type="content" source="../../images/scenario-task-reminders-adaptive-card.png" alt-text="ステータスの更新をTeamsするアダプティブ カード。":::

### <a name="after-running-the-flow"></a><span data-ttu-id="3fbd3-177">フローの実行後</span><span class="sxs-lookup"><span data-stu-id="3fbd3-177">After running the flow</span></span>

:::image type="content" source="../../images/scenario-task-reminders-spreadsheet-after.png" alt-text="現在入力されている状態エントリを持つ状態レポートを含むワークシート。":::
