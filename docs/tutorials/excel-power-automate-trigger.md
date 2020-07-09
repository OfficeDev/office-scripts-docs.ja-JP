---
title: 自動電源自動化フローを使用してスクリプトを自動的に実行する
description: 自動的な外部トリガー (Outlook 経由でメールを受信する) を使用して、Power automatic を使用して、web 上で Excel の Office スクリプトを実行する方法についてのチュートリアルです。
ms.date: 07/01/2020
localization_priority: Priority
ms.openlocfilehash: fc98fb36fd5a8c5ef10bc3b767d6f5add0306246
ms.sourcegitcommit: edf58aed3cd38f57e5e7227465a1ef5515e15703
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 07/08/2020
ms.locfileid: "45081629"
---
# <a name="automatically-run-scripts-with-automated-power-automate-flows-preview"></a>自動電源自動化フロー (プレビュー) を使用してスクリプトを自動的に実行する

このチュートリアルでは、自動[電源自動化](https://flow.microsoft.com)ワークフローを使用して web 上の Excel 用 Office スクリプトを使用する方法について説明します。 スクリプトは、電子メールを受信するたびに自動的に実行され、Excel ブックに電子メールの情報を記録します。

## <a name="prerequisites"></a>前提条件

[!INCLUDE [Tutorial prerequisites](../includes/tutorial-prerequisites.md)]

> [!IMPORTANT]
> このチュートリアルでは、 [「Power オートメーションチュートリアルを使用して web 上の Excel で Office スクリプトを実行する」](excel-power-automate-manual.md)を完了していることを前提としています。

## <a name="prepare-the-workbook"></a>ブックの準備

Power オートメーションは、ブックコンポーネントへのアクセスなどの[相対参照](../develop/power-automate-integration.md#avoid-using-relative-references)を使用できません `Workbook.getActiveWorksheet` 。 そのため、Power オートメーションが参照できるように、名前が一貫したブックとワークシートが必要です。

1. **Myworkbook**という名前の新しいブックを作成します。

2. [**自動化**] タブに移動して、[**コードエディター**] を選択します。

3. [**新しいスクリプト**] を選択します。

4. 既存のコードを次のスクリプトに置き換え、[**実行**] を押します。 これにより、ワークシート、テーブル、およびピボットテーブル名が一致するブックが設定されます。

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

## <a name="create-an-office-script-for-your-automated-workflow"></a>自動化されたワークフロー用の Office スクリプトを作成する

電子メールから情報をログに記録するスクリプトを作成してみましょう。 最もメールを受信する曜日と、そのメールを送信している一意の送信者の数を知りたいと考えています。 ブックに**は、****日付**、曜日、**電子メールアドレス**、および**件名**の列を持つテーブルがあります。 また、このワークシートに**は、曜日と****電子メールアドレス**(行階層) に対してピボットされたピボットテーブルもあります。 一意の**件名**の数は、表示される集計情報 (データ階層) です。 メールテーブルを更新した後、スクリプトによってピボットテーブルが更新されるようになります。

1. **コードエディター**で、[**新しいスクリプト**] を選択します。

2. このチュートリアルで後で作成するフローによって、受信した各電子メールについてのスクリプト情報が送信されます。 スクリプトは、関数内のパラメーターを使用して、その入力を受け入れる必要があり `main` ます。 既定のスクリプトを次のスクリプトに置き換えます。

    ```TypeScript
    function main(
      workbook: ExcelScript.Workbook,
      from: string,
      dateReceived: string,
      subject: string) {

    }
    ```

3. このスクリプトには、ブックのテーブルとピボットテーブルへのアクセス権が必要です。 次のコードをスクリプトの本文に追加します。その後、次のコードを開き `{` ます。

    ```TypeScript
    // Get the email table.
    let emailWorksheet = workbook.getWorksheet("Emails");
    let table = emailWorksheet.getTable("EmailTable");
  
    // Get the PivotTable.
    let pivotTableWorksheet = workbook.getWorksheet("SubjectPivot");
    let pivotTable = pivotTableWorksheet.getPivotTable("Pivot");
    ```

4. `dateReceived`パラメーターの型がである `string` 。 曜日を簡単に取得できるように、を[ `Date` オブジェクト](../develop/javascript-objects.md#date)に変換しましょう。 その後、その日の番号の値をより読みやすいバージョンにマップする必要があります。 次のコードをスクリプトの最後に追加してから、閉じる前にし `}` ます。

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

5. 文字列には `subject` 、"RE:" という返信タグを含めることができます。 これを文字列から削除して、同じスレッドの電子メールがテーブルの同じ件名を持つようにしましょう。 次のコードをスクリプトの最後に追加してから、閉じる前にし `}` ます。

    ```TypeScript
    // Remove the reply tag from the email subject to group emails on the same thread.
    let subjectText = subject.replace("Re: ", "");
    subjectText = subjectText.replace("RE: ", "");
    ```

6. これで、電子メールデータの形式が希望どおりになったので、電子メールの表に行を追加しましょう。 次のコードをスクリプトの最後に追加してから、閉じる前にし `}` ます。

    ```TypeScript
    // Add the parsed text to the table.
    table.addRow(-1, [dateReceived, dayText, from, subjectText]);
    ```

7. 最後に、ピボットテーブルが更新されていることを確認してみましょう。 次のコードをスクリプトの最後に追加してから、閉じる前にし `}` ます。

    ```TypeScript
    // Refresh the PivotTable to include the new row.
    pivotTable.refresh();
    ```

8. スクリプト**レコード**の名前を変更し、[**スクリプトの保存**] をクリックします。

これで、パワー自動化ワークフローのためのスクリプトの準備が整いました。 次のスクリプトのようになります。

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

## <a name="create-an-automated-workflow-with-power-automate"></a>Power 自動化を使用して自動化されたワークフローを作成する

1. [パワー自動化プレビューサイト](https://flow.microsoft.com)にサインインします。

2. 画面の左側に表示されるメニューで、[**作成**] を押します。 これにより、新しいワークフローを作成する方法の一覧が表示されます。

    ![パワー自動化の [作成] ボタン。](../images/power-automate-tutorial-1.png)

3. [**空白から開始**] セクションで、[**自動フロー**] を選択します。 これにより、電子メールの受信など、イベントによってトリガーされるワークフローが作成されます。

    ![電源自動化の [フローの自動化] オプション。](../images/power-automate-params-tutorial-1.png)

4. 表示されるダイアログウィンドウで、[**フロー名**] テキストボックスにフローの名前を入力します。 次に、[**フローのトリガーを選択して**ください] の一覧から **、新しい電子メールを受信するタイミング**を選択します。 検索ボックスを使用してオプションを検索する必要がある場合があります。 最後に、[**作成**] を押します。

    ![「新しい電子メールの到着」オプションを示す [パワー・自動化] の [自動フロー] ウィンドウの構築の一部。](../images/power-automate-params-tutorial-2.png)

    > [!NOTE]
    > このチュートリアルでは、Outlook を使用します。 代わりに、優先する電子メールサービスを自由に使用できますが、一部のオプションは異なる場合があります。

5. **新しい手順**を押します。

6. [**標準**] タブを選択し、[ **Excel Online (Business)**] を選択します。

    ![Excel Online (Business) の電源自動化オプション。](../images/power-automate-tutorial-4.png)

7. [**アクション**] で、[**スクリプトを実行する (プレビュー)**] を選択します。

    ![実行スクリプトのパワー自動処理オプション (プレビュー)。](../images/power-automate-tutorial-5.png)

8. **実行スクリプト**コネクタについて、次の設定を指定します。

    - **場所**: OneDrive for business
    - **ドキュメントライブラリ**: OneDrive
    - **ファイル**: MyWorkbook.xlsx
    - **スクリプト**: メールの録音
    - **from**: From *(Outlook の動的コンテンツ)*
    - **dateReceived**: 受信時刻 *(Outlook からの動的なコンテンツ)*
    - **件名**: 件名 *(Outlook の動的コンテンツ)*

    *スクリプトのパラメーターは、スクリプトが選択された後にのみ表示されることに注意してください。*

    ![実行スクリプトのパワー自動処理オプション (プレビュー)。](../images/power-automate-params-tutorial-3.png)

9. [**保存**します。

これでフローが有効になります。 Outlook を使用して電子メールを受信するたびに、スクリプトが自動的に実行されます。

## <a name="manage-the-script-in-power-automate"></a>パワー自動化でスクリプトを管理する

1. [メインパワーの自動化] ページで、[**マイフロー**] を選択します。

    ![パワー自動化の [マイフロー] ボタン。](../images/power-automate-tutorial-7.png)

2. フローを選択します。 ここに、実行履歴が表示されます。 ページを更新するか、[すべての**実行**の更新] ボタンをクリックすると、履歴を更新できます。 このフローは、電子メールの受信後すぐにトリガーされます。 自分のメールを送信してフローをテストします。

フローがトリガーされ、スクリプトが正常に実行されると、ブックのテーブルとピボットテーブルの更新が表示されます。

![フローが2回実行された後の電子メールテーブル。](../images/power-automate-params-tutorial-4.png)

![フローが2回実行された後のピボットテーブル。](../images/power-automate-params-tutorial-5.png)

## <a name="next-steps"></a>次の手順

Office スクリプトを Power オートメーションで接続する方法の詳細については、「 [Power オートメーションで Office スクリプトを実行](../develop/power-automate-integration.md)する」を参照してください。

また、[自動タスクリマインダーのサンプルシナリオ](../resources/scenarios/task-reminders.md)を参照して、Office スクリプトと Teams のアダプティブカードを組み合わせたパワーオートメーションを組み合わせる方法を確認することもできます。
