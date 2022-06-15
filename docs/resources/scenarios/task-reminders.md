---
title: 'Office スクリプトのサンプル シナリオ: タスクリマインダーの自動化'
description: Power Automateカードとアダプティブ カードを使用するサンプルでは、プロジェクト管理スプレッドシートでタスクのリマインダーを自動化します。
ms.date: 06/29/2021
ms.localizationpriority: medium
ms.openlocfilehash: 08f3713210e83162f86d38bc8eb33d76bf8a7288
ms.sourcegitcommit: dd01979d34b3499360d2f79a56f8a8f24f480eed
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 06/15/2022
ms.locfileid: "66088114"
---
# <a name="office-scripts-sample-scenario-automated-task-reminders"></a>Office スクリプトのサンプル シナリオ: タスクリマインダーの自動化

このシナリオでは、プロジェクトを管理しています。 Excel ワークシートを使用して、従業員の状態を毎月追跡します。 多くの場合、ユーザーに状態を入力するように通知する必要があるため、そのリマインダー プロセスを自動化することにしました。

ステータス フィールドが見つからないユーザーにメッセージを送信するPower Automate フローを作成し、その応答をスプレッドシートに適用します。 これを行うには、ブックの操作を処理するスクリプトのペアを開発します。 最初のスクリプトは空白の状態のユーザーの一覧を取得し、2 番目のスクリプトは右側の行に状態文字列を追加します。 また、[アダプティブ カードTeams](/microsoftteams/platform/task-modules-and-cards/what-are-cards)使用して、従業員に通知から直接状態を入力してもらうこともできます。

## <a name="scripting-skills-covered"></a>スクリプティング スキルの説明

- Power Automateでフローを作成する
- スクリプトにデータを渡す
- スクリプトからデータを返す
- アダプティブ カードのTeams
- テーブル

## <a name="prerequisites"></a>前提条件

このシナリオでは[、Power Automate](https://flow.microsoft.com)とMicrosoft Teamsを使用[します](https://www.microsoft.com/microsoft-365/microsoft-teams/group-chat-software)。 両方とも、Office スクリプトの開発に使用するアカウントに関連付けられている必要があります。 Microsoft Developer サブスクリプションに無料でアクセスして、これらのアプリケーションについて学習し、これらのアプリケーションを操作するには、[Microsoft 365開発者プログラム](https://developer.microsoft.com/microsoft-365/dev-program)への参加を検討してください。

## <a name="setup-instructions"></a>セットアップ手順

1. <a href="task-reminders.xlsx"> OneDriveにtask-reminders.xlsx</a>をダウンロードします。

1. Excel on the webでブックを開きます。

1. まず、スプレッドシートに表示されない状態レポートを持つすべての従業員を取得するためのスクリプトが必要です。 [ **自動化** ] タブで [ **新しいスクリプト** ] を選択し、次のスクリプトをエディターに貼り付けます。

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

1. **Get People** という名前のスクリプトを保存します。

1. 次に、状態レポート カードを処理し、新しい情報をスプレッドシートに配置するための 2 番目のスクリプトが必要です。 [コード エディター] 作業ウィンドウで、[ **新しいスクリプト** ] を選択し、次のスクリプトをエディターに貼り付けます。

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

1. [状態の保存] という名前でスクリプト **を保存します**。

1. 次に、フローを作成する必要があります。 [Power Automate](https://flow.microsoft.com/)を開きます。

    > [!TIP]
    > フローを作成したことがない場合は、チュートリアル「Power Automate[を含むスクリプトの使用を開始](../../tutorials/excel-power-automate-manual.md)して基本を学習する」を参照してください。

1. 新しい **インスタント フロー** を作成します。

1. オプションから **[手動でフローをトリガー** する] を選択し、[ **作成**] を選択します。

1. フローでは、Get **People** スクリプトを呼び出して、空の状態フィールドを持つすべての従業員を取得する必要があります。 [**新しい手順**] を選択し、**オンライン (ビジネス) Excel** 選択します。 **[アクション]** で、**[スクリプトの実行]** を選択します。 フロー ステップに次のエントリを指定します。

    - **場所**: OneDrive for Business
    - **ドキュメント ライブラリ**: OneDrive
    - **ファイル**: task-reminders.xlsx *(ファイル ブラウザーから選択)*
    - **スクリプト**: ユーザーを取得する

    :::image type="content" source="../../images/scenario-task-reminders-first-flow-step.png" alt-text="最初のスクリプト フローの実行手順を示すPower Automate フロー。":::

1. 次に、このフローでは、スクリプトによって返された配列内の各 Employee を処理する必要があります。 [**新しい手順**] を選択し、[**アダプティブ カードをTeams ユーザーに投稿する] を選択し、応答を待ちます**。

1. **[受信者]** フィールドに、動的コンテンツから **電子メール** を追加します (選択範囲にはExcelロゴが表示されます)。 **電子メール** を追加すると、フロー ステップが各ブロックに **適用** されて囲まれます。 つまり、配列はPower Automateによって反復処理されます。

1. アダプティブ カードを送信するには、カードの [JSON](https://www.w3schools.com/whatis/whatis_json.asp) を **メッセージ** として提供する必要があります。 [アダプティブ カード デザイナー](https://adaptivecards.io/designer/)を使用して、カスタム カードを作成できます。 このサンプルでは、次の JSON を使用します。  

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

1. 残りのフィールドに次のように入力します。

    - **メッセージの更新**: 状態レポートを送信していただきありがとうございます。 応答がスプレッドシートに正常に追加されました。
    - **カードを更新する必要があります**: はい

1. **[各ブロックに適用**] で、[**アダプティブ カードをTeams ユーザーに投稿し、応答を待機** する] の下にある [**アクションの追加**] を選択します。 **Excel Online (Business)** を選択します。 **[アクション]** で、**[スクリプトの実行]** を選択します。 フロー ステップに次のエントリを指定します。

    - **場所**: OneDrive for Business
    - **ドキュメント ライブラリ**: OneDrive
    - **ファイル**: task-reminders.xlsx *(ファイル ブラウザーから選択)*
    - **スクリプト**: 状態の保存
    - **senderEmail**: 電子メール *(Excelからの動的コンテンツ)*
    - **statusReportResponse**: 応答 *(Teamsからの動的コンテンツ)*

    :::image type="content" source="../../images/scenario-task-reminders-last-flow-step.png" alt-text="各ステップへの適用を示すPower Automate フロー。":::

1. フローを保存します。

## <a name="running-the-flow"></a>フローの実行

フローをテストするには、空の状態のテーブル行で、Teams アカウントに関連付けられた電子メール アドレスが使用されていることを確認します (テスト中は自分のメール アドレスを使用する必要があります)。 フロー エディター ページの **[テスト** ] ボタンを使用するか、[ **マイ フロー** ] タブでフローを実行します。メッセージが表示されたら、必ずアクセスを許可してください。

Power AutomateからTeamsまでアダプティブ カードを受け取る必要があります。 カードの状態フィールドに入力すると、フローは続行され、指定した状態でスプレッドシートが更新されます。

### <a name="before-running-the-flow"></a>フローを実行する前に

:::image type="content" source="../../images/scenario-task-reminders-spreadsheet-before.png" alt-text="1 つの不足している状態エントリを含む状態レポートを含むワークシート。":::

### <a name="receiving-the-adaptive-card"></a>アダプティブ カードの受信

:::image type="content" source="../../images/scenario-task-reminders-adaptive-card.png" alt-text="従業員に状態の更新を求めるTeamsのアダプティブ カード。":::

### <a name="after-running-the-flow"></a>フローを実行した後

:::image type="content" source="../../images/scenario-task-reminders-spreadsheet-after.png" alt-text="現在入力された状態エントリを含む状態レポートを含むワークシート。":::
