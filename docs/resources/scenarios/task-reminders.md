---
title: 'Officeスクリプトのサンプル シナリオ: タスクの自動アラーム'
description: プロジェクト管理スプレッドシートでPower Automateアダプティブ カードを使用するサンプルは、タスクリマインダーを自動化します。
ms.date: 11/30/2020
localization_priority: Normal
ms.openlocfilehash: 1297f10e45c515079994d659378331fc4a2be744
ms.sourcegitcommit: 4693c8f79428ec74695328275703af0ba1bfea8f
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 06/23/2021
ms.locfileid: "53074663"
---
# <a name="office-scripts-sample-scenario-automated-task-reminders"></a>Officeスクリプトのサンプル シナリオ: タスクの自動アラーム

このシナリオでは、プロジェクトを管理しています。 毎月従業員のExcelを追跡するには、ユーザーのワークシートを使用します。 多くの場合、ユーザーに自分の状態を入力することを通知する必要があります。そのため、そのリマインダー プロセスを自動化することを決めました。

ステータス フィールドが見つからないPower Automateに対するメッセージ フローを作成し、その応答をスプレッドシートに適用します。 これを行うには、ブックの操作を処理するためのスクリプトのペアを開発します。 最初のスクリプトは、空の状態を持つユーザーの一覧を取得し、2 番目のスクリプトは、右側の行に状態文字列を追加します。 また、アダプティブ カードをTeams[して、](/microsoftteams/platform/task-modules-and-cards/what-are-cards)従業員に通知から直接ステータスを入力します。

## <a name="scripting-skills-covered"></a>スクリプティングのスキルをカバー

- [フローの作成] Power Automate
- スクリプトにデータを渡す
- スクリプトからデータを返す
- Teamsアダプティブ カード
- テーブル

## <a name="prerequisites"></a>前提条件

このシナリオでは[、Power Automate](https://flow.microsoft.com)とMicrosoft Teamsを[使用します](https://www.microsoft.com/microsoft-365/microsoft-teams/group-chat-software)。 両方とも、スクリプトの開発に使用するアカウントに関連付Officeがあります。 Microsoft Developer サブスクリプションに無料でアクセスして、これらのアプリケーションについて学び、これらのアプリケーションを使用するには、開発者プログラムに参加Microsoft 365[検討してください](https://developer.microsoft.com/microsoft-365/dev-program)。

## <a name="setup-instructions"></a>セットアップ手順

1. ユーザー <a href="task-reminders.xlsx">task-reminders.xlsx</a>にダウンロードOneDrive。

2. ブックを開Excel on the web。

3. [自動化] **タブで** 、[すべてのスクリプト] **を開きます**。

4. まず、スプレッドシートに不足している状態レポートを持つすべての従業員を取得するスクリプトが必要です。 [コード **エディター] 作業ウィンドウ** で、[新しいスクリプト] **を押** して、次のスクリプトをエディターに貼り付けます。

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

5. [ユーザーの取得] という名前のスクリプト **を保存します**。

6. 次に、ステータス レポート カードを処理し、新しい情報をスプレッドシートに入れる 2 番目のスクリプトが必要です。 [コード **エディター] 作業ウィンドウ** で、[新しいスクリプト] **を押** して、次のスクリプトをエディターに貼り付けます。

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

7. [状態の保存] という名前のスクリプト **を保存します**。

8. 次に、フローを作成する必要があります。 [ファイル[Power Automate] を開きます](https://flow.microsoft.com/)。

    > [!TIP]
    > 前にフローを作成したことがない場合は、チュートリアル「スクリプトの使用[](../../tutorials/excel-power-automate-manual.md)を開始する」を参照し、Power Automateを確認してください。

9. 新しいインスタント フロー **を作成します**。

10. [オプション **からフローを手動でトリガーする** ] を選択し、[作成] を **押します**。

11. フローは、空の状態フィールドを持つすべての従業員を取得するために Get **People** スクリプトを呼び出す必要があります。 [**新しい手順] を** 押し、[**オンラインExcel (Business) を選択します**。 **[アクション]** で、**[スクリプトの実行]** を選択します。 フロー ステップに次のエントリを指定します。

    - **場所**: OneDrive for Business
    - **ドキュメント ライブラリ**: OneDrive
    - **ファイル**: task-reminders.xlsx *(ファイル ブラウザーから選択)*
    - **スクリプト**: ユーザーを取得する

    :::image type="content" source="../../images/scenario-task-reminders-first-flow-step.png" alt-text="最初Power Automateスクリプト フロー の手順を示す手順を示す手順を示します。":::

12. 次に、フローは、スクリプトによって返される配列内の各 Employee を処理する必要があります。 [**新しい手順]** を押し、[アダプティブ カードをユーザーにTeamsして応答を **待つ] を選択します**。

13. [受信者 **] フィールド** に動的 **コンテンツから電子** メールを追加します (選択すると、Excelロゴが表示されます)。 メール **を** 追加すると、フロー ステップは各ブロックに **適用されます** 。 つまり、配列は配列によって反復処理Power Automate。

14. アダプティブ カードを送信するには、カードの JSON をメッセージとして提供する必要 **があります**。 アダプティブ カード デザイナーを [使用してカスタム](https://adaptivecards.io/designer/) カードを作成できます。 このサンプルでは、次の JSON を使用します。  

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

15. 残りのフィールドに次のように入力します。

    - **更新メッセージ**: ステータス レポートを提出してありがとうございます。 応答がスプレッドシートに正常に追加されました。
    - **カードを更新する必要があります**: はい

16. [各 **ブロックに適用]** で、[アダプティブ カードをユーザーに投稿Teams応答を待つ] の後、[アクションの追加 **] を押します**。 [オンライン **Excel (Business) を選択します**。 **[アクション]** で、**[スクリプトの実行]** を選択します。 フロー ステップに次のエントリを指定します。

    - **場所**: OneDrive for Business
    - **ドキュメント ライブラリ**: OneDrive
    - **ファイル**: task-reminders.xlsx *(ファイル ブラウザーから選択)*
    - **スクリプト**: 状態の保存
    - **senderEmail**: メール *(メールからの動的Excel)*
    - **statusReportResponse**: 応答 *(Teams)*

    :::image type="content" source="../../images/scenario-task-reminders-last-flow-step.png" alt-text="各Power Automate適用を示すフローを示す手順を示します。":::

17. フローを保存します。

## <a name="running-the-flow"></a>フローの実行

フローをテストするには、状態が空白のテーブル行で Teams アカウントに関連付けられている電子メール アドレスを使用します (テスト中は、独自の電子メール アドレスを使用する必要があります)。

フロー デザイナーから **[テスト]** を選択するか、[マイ フロー] ページから **フローを実行** できます。 フローを開始し、必要な接続の使用を受け入れた後、アダプティブ カードを受信する必要Power AutomateからTeams。 カードの状態フィールドに入力すると、フローは続行され、指定した状態でスプレッドシートが更新されます。

### <a name="before-running-the-flow"></a>フローを実行する前に

:::image type="content" source="../../images/scenario-task-reminders-spreadsheet-before.png" alt-text="不足している状態エントリが 1 つ含まれる状態レポートを含むワークシート。":::

### <a name="receiving-the-adaptive-card"></a>アダプティブ カードの受信

:::image type="content" source="../../images/scenario-task-reminders-adaptive-card.png" alt-text="ステータスの更新をTeamsするアダプティブ カード。":::

### <a name="after-running-the-flow"></a>フローの実行後

:::image type="content" source="../../images/scenario-task-reminders-spreadsheet-after.png" alt-text="現在入力されている状態エントリを持つ状態レポートを含むワークシート。":::
