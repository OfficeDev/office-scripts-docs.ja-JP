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
# <a name="office-scripts-sample-scenario-automated-task-reminders"></a>Officeスクリプトのサンプル シナリオ: 自動タスク アラーム

このシナリオでは、プロジェクトを管理しています。 Excelワークシートを使用して、毎月従業員のステータスを追跡します。 多くの場合、ユーザーにステータスを入力するよう促す必要があるため、リマインダープロセスを自動化することにしました。

ステータスフィールドが欠落している人にメッセージを送信するPower Automateフローを作成し、その回答をスプレッドシートに適用します。 これを行うには、ブックの操作を処理するスクリプトのペアを開発します。 最初のスクリプトは、空白の状態のユーザーのリストを取得し、2 番目のスクリプトは、右側の行にステータス文字列を追加します。 また[、Teamsアダプティブカード](/microsoftteams/platform/task-modules-and-cards/what-are-cards)を使用して、従業員に通知から直接ステータスを入力させます。

## <a name="scripting-skills-covered"></a>スクリプティングのスキルをカバー

- Power Automateでのフローの作成
- スクリプトへのデータの受け渡し
- スクリプトからデータを返す
- Teamsアダプティブカード
- テーブル

## <a name="prerequisites"></a>前提条件

このシナリオでは[、Power Automate](https://flow.microsoft.com)と[Microsoft Teams](https://www.microsoft.com/microsoft-365/microsoft-teams/group-chat-software)を使用します。 Office スクリプトの開発に使用するアカウントに関連付ける必要があります。 これらのアプリケーションについて学び、これらのアプリケーションを操作するための Microsoft Developer サブスクリプションへの無料アクセスについては[、Microsoft 365開発者プログラム](https://developer.microsoft.com/microsoft-365/dev-program)への参加を検討してください。

## <a name="setup-instructions"></a>セットアップ手順

1. <a href="task-reminders.xlsx">OneDriveにtask-reminders.xlsx</a>をダウンロードします。

2. ブックをExcel on the webで開きます。

3. [ **自動化** ] タブで、[ **すべてのスクリプト]** を開きます。

4. まず、スプレッドシートに不足しているステータス レポートを持つすべての従業員を取得するスクリプトが必要です。 [ **コード エディター** ] 作業ウィンドウで[ **新規スクリプト]** を押し、次のスクリプトをエディタに貼り付けます。

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

5. スクリプトを **Get People** という名前で保存します。

6. 次に、ステータス レポート カードを処理し、新しい情報をスプレッドシートに配置する 2 番目のスクリプトが必要です。 [ **コード エディター** ] 作業ウィンドウで[ **新規スクリプト]** を押し、次のスクリプトをエディタに貼り付けます。

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

7. スクリプトを保存ステータスという名前で **保存します**。

8. ここで、フローを作成する必要があります。 [Power Automate](https://flow.microsoft.com/)を開きます。

    > [!TIP]
    > フローを作成したことがない場合は、チュートリアル[「Power Automateを含むスクリプトの使用を開始](../../tutorials/excel-power-automate-manual.md)する」を参照して基本を学習してください。

9. 新しい **インスタントフロー** を作成する:

10. オプションから **[フローを手動でトリガー** ] を選択し、[ **作成]** を押します。

11. フローは、空のステータスフィールドを持つすべての従業員を取得するために **Get People** スクリプトを呼び出す必要があります。 **[新規ステップ]** を押して、[**オンライン (ビジネス)] Excel** 選択します。 [ **アクション]** で、[ **スクリプトの実行**] を選択します。 フローステップに対して以下のエントリを入力します。

    - **場所**: OneDrive for Business
    - **ドキュメント ライブラリ**: OneDrive
    - **ファイル**: task-reminders.xlsx *(ファイルブラウザを使用して選択)*
    - **スクリプト**: 人をゲット

    :::image type="content" source="../../images/scenario-task-reminders-first-flow-step.png" alt-text="スクリプトフローの最初の実行ステップを示すPower Automateフロー":::

12. 次に、フローは、スクリプトによって返される配列内の各従業員を処理する必要があります。 **[新規ステップ]** を押して、[**アダプティブ カードをTeamsユーザーにポストし、応答を待つ**] を選択します。

13. [**受信者**] フィールドに、動的コンテンツから **電子メール** を追加します (選択内容には、そのコンテンツによってExcelロゴが表示されます)。 **電子メール** を追加すると、フロー ステップは **各ブロックに適用** によって囲まれます。 つまり、配列はPower Automateによって反復処理されます。

14. アダプティブ カードを送信するには、カードの JSON を **メッセージ** として指定する必要があります。 [アダプティブ カード デザイナー](https://adaptivecards.io/designer/)を使用して、カスタム カードを作成できます。 このサンプルでは、次の JSON を使用します。  

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

    - **更新メッセージ**: 進捗レポートを提出していただきありがとうございます。 スプレッドシートに返信が追加されました。
    - **カードを更新する必要があります**: はい

16. **[各ブロックに適用]** で、[**アダプティブ カードをTeamsユーザーにポストして応答を待つ**] に続いて、[**アクションの追加]** をクリックします。 [**オンライン (ビジネス)] Excel** 選択します。 [ **アクション]** で、[ **スクリプトの実行**] を選択します。 フローステップに対して以下のエントリを入力します。

    - **場所**: OneDrive for Business
    - **ドキュメント ライブラリ**: OneDrive
    - **ファイル**: task-reminders.xlsx *(ファイルブラウザを使用して選択)*
    - **スクリプト**: 保存ステータス
    - **送信者電子メール**: 電子メール *(Excelからの動的コンテンツ)*
    - **ステータスレポート応答**: 応答 *(Teamsからの動的コンテンツ)*

    :::image type="content" source="../../images/scenario-task-reminders-last-flow-step.png" alt-text="適用先のステップを示すPower Automateフロー":::

17. フローを保存します。

## <a name="running-the-flow"></a>フローの実行

フローをテストするには、空白のステータスを持つテーブル行で、Teamsアカウントに関連付けられた電子メール アドレスが使用されていることを確認します (テスト中は、独自の電子メール アドレスを使用する必要があります)。

フロー デザイナーから **[テスト** ] を選択するか、[ **自分のフロー** ] ページからフローを実行できます。 フローを開始し、必要な接続の使用を受け入れた後、Power AutomateからTeamsまでのアダプティブ カードを受け取る必要があります。 カードのステータスフィールドに入力すると、フローが続行され、提供したステータスでスプレッドシートが更新されます。

### <a name="before-running-the-flow"></a>フローを実行する前に

:::image type="content" source="../../images/scenario-task-reminders-spreadsheet-before.png" alt-text="1 つの不足している状態エントリを含む進捗レポートを含むワークシート":::

### <a name="receiving-the-adaptive-card"></a>アダプティブカードの受け取り

:::image type="content" source="../../images/scenario-task-reminders-adaptive-card.png" alt-text="従業員にステータスの更新を求めるTeamsのアダプティブ カード":::

### <a name="after-running-the-flow"></a>フローの実行後

:::image type="content" source="../../images/scenario-task-reminders-spreadsheet-after.png" alt-text="現在入力された状態エントリを含む進捗レポートを含むワークシート":::
