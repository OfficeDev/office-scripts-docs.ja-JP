---
title: 自動で実行される Power Automate フローにスクリプトからデータを返す
description: Power Automate を使用して Excel on the web 用の Office スクリプトを実行してリマインダー メールを送信する方法を示すチュートリアル。
ms.date: 06/29/2021
localization_priority: Priority
ms.openlocfilehash: c928ca612e08098c85307e22185e2240f47a5adcdfe028815c4dd837b4501c9f
ms.sourcegitcommit: 75f7ed8c2d23a104acc293f8ce29ea580b4fcdc5
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 08/11/2021
ms.locfileid: "57846460"
---
# <a name="return-data-from-a-script-to-an-automatically-run-power-automate-flow"></a>自動で実行される Power Automate フローにスクリプトからデータを返す

このチュートリアルでは、自動化された [Power Automate](https://flow.microsoft.com) ワークフローの一部として、Excel on the web 用の Office スクリプトから情報を返す方法について説明します。 スケジュールを確認し、フローに従ってリマインダー メールを送信するスクリプトを作成します。 このフローは定期的に実行され、ユーザーに代わってこれらのリマインダーを提供します。

> [!TIP]
> Office スクリプトを初めて使用する場合は、チュートリアルの「[Excel on the web で Office スクリプトを記録、編集、作成する](excel-tutorial.md)」から始めることをお勧めします。
>
> Power Automate を初めて使用する場合は、チュートリアルの「[手動 Power Automate フローからスクリプトを呼び出す](excel-power-automate-manual.md)」と「[自動で実行される Power Automate フロー内で、データをスクリプトに渡す](excel-power-automate-trigger.md)」から始めることを勧めします。
>
> [Office スクリプトは TypeScript を使用](../overview/code-editor-environment.md)します。このチュートリアルは、JavaScript や TypeScript について初級から中級レベルの知識を持つユーザーを対象としています。 JavaScript を使い慣れていない場合は、「[Mozilla の JavaScript チュートリアル](https://developer.mozilla.org/docs/Web/JavaScript/Guide/Introduction)」から始めることをお勧めします。

## <a name="prerequisites"></a>前提条件

[!INCLUDE [Tutorial prerequisites](../includes/power-automate-tutorial-prerequisites.md)]

## <a name="prepare-the-workbook"></a>ブックを準備する

1. ブック <a href="on-call-rotation.xlsx">on-call-rotation.xlsx</a> を 自分の OneDrive にダウンロードします。

1. Excel on the web で **on-call-rotation.xlsx** を開きます。

1. テーブルに行を追加して、自分の名前、メール アドレス、および現在の日付と重なるように開始日と終了日を入力します。

    > [!IMPORTANT]
    > これから作成するスクリプトは、テーブル内の最初に一致するエントリを使用するため、自分の名前が現在の週のどの行よりも上にあることを確認してください。

    :::image type="content" source="../images/power-automate-return-tutorial-1.png" alt-text="呼び出し時の回転テーブルのデータを含むワークシート。":::

## <a name="create-an-office-script"></a>Office スクリプトを作成する

1. **[オートメーション]** タブに移動して **[すべてのスクリプト]** を選択します。

1. **[新しいスクリプト]** を選択します。

1. スクリプトに **Get On-Call Person** という名前を付けます。

1. これで空のスクリプトができました。 スクリプトを使用して、スプレッドシートからメール アドレスを取得します。 文字列が返されるように、`main` を次のように変更します。

    ```TypeScript
    function main(workbook: ExcelScript.Workbook) : string {
    }
    ```

1. 続いて、テーブルからすべてのデータを取得する必要があります。 それにより、スクリプトを使用して各行を確認できます。 `main` 関数に次のコードを追加します。

    ```TypeScript
    // Get the H1 worksheet.
    let worksheet = workbook.getWorksheet("H1");

    // Get the first (and only) table in the worksheet.
    let table = worksheet.getTables()[0];

    // Get the data from the table.
    let tableValues = table.getRangeBetweenHeaderAndTotal().getValues();
    ```

1. テーブル内の日付は、[Excel の日付システム](https://support.microsoft.com/office/date-systems-in-excel-e7fe7167-48a9-4b96-bb53-5612a800b487)を使用して保存されます。 これらの日付は、比較できるように JavaScript の日付に変換する必要があります。 ヘルパー関数をスクリプトに追加します。 `main` 関数の外に次のコードを追加します。

    ```TypeScript
    // Convert the Excel date to a JavaScript Date object.
    function convertDate(excelDateValue: number) {
        let javaScriptDate = new Date(Math.round((excelDateValue - 25569) * 86400 * 1000));
        return javaScriptDate;
    }
    ```

1. 次に、現在誰が呼び出し期間中かを把握する必要があります。 それらの行では、開始日と終了日の間に現在の日付が含まれています。 ここでは、一度に 1 人だけが呼び出し期間であると想定してスクリプトを作成します。 スクリプトで配列を返して複数の値を処理することもできますが、現時点では、最初に一致するメール アドレスを返すようにします。 次の関数を `main` 関数の最後に追加します。

    ```TypeScript
    // Look for the first row where today's date is between the row's start and end dates.
    let currentDate = new Date();
    for (let row = 0; row < tableValues.length; row++) {
        let startDate = convertDate(tableValues[row][2] as number);
        let endDate = convertDate(tableValues[row][3] as number);
        if (startDate <= currentDate && endDate >= currentDate) {
            // Return the first matching email address.
            return tableValues[row][1].toString();
        }
    }
    ```

1. 最終的なスクリプトは、次のようになります。

    ```TypeScript
    function main(workbook: ExcelScript.Workbook) : string {
        // Get the H1 worksheet.
        let worksheet = workbook.getWorksheet("H1");

        // Get the first (and only) table in the worksheet.
        let table = worksheet.getTables()[0];
    
        // Get the data from the table.
        let tableValues = table.getRangeBetweenHeaderAndTotal().getValues();
    
        // Look for the first row where today's date is between the row's start and end dates.
        let currentDate = new Date();
        for (let row = 0; row < tableValues.length; row++) {
            let startDate = convertDate(tableValues[row][2] as number);
            let endDate = convertDate(tableValues[row][3] as number);
            if (startDate <= currentDate && endDate >= currentDate) {
                // Return the first matching email address.
                return tableValues[row][1].toString();
            }
        }
    }

    // Convert the Excel date to a JavaScript Date object.
    function convertDate(excelDateValue: number) {
        let javaScriptDate = new Date(Math.round((excelDateValue - 25569) * 86400 * 1000));
        return javaScriptDate;
    }
    ```

## <a name="create-an-automated-workflow-with-power-automate"></a>Power Automate を使用して自動化されたワークフローを作成する

1. [「Power Automate のサイト」](https://flow.microsoft.com)にサインインします。

1. 画面の左側に表示されるメニューで、**[作成]** を選択します。 これにより、新しいワークフローを作成する方法の一覧を表示できます。

    :::image type="content" source="../images/power-automate-tutorial-1.png" alt-text="Power Automate の [作成] ボタン。":::

1. **[空白から開始]** セクションで **[スケジュール済みクラウド フロー]** を選択します。

    :::image type="content" source="../images/power-automate-return-tutorial-2.png" alt-text="Power Automate の [スケジュール済みクラウド フロー] ボタン。":::

1. 続いて、このフローのスケジュールを設定します。 使用しているスプレッドシートには、2021 年前半の毎週月曜日から始まる新しい呼び出し期間の割り当てが含まれています。 月曜日の朝一番に実行するようにフローを設定します。 次のオプションを使用して、毎週月曜日に実行するようにフローを構成します。

    - **フロー名**: Notify On-Call Person
    - **開始**: 21/1/4 時間 1:00 AM
    - **繰り返し間隔**: 1 週
    - **設定曜日**: 月

    :::image type="content" source="../images/power-automate-return-tutorial-3.png" alt-text="オプションが表示された Power Automate の [スケジュールされたクラウド フローを作成する] ダイアログ。オプションには、フロー名、開始時刻、繰り返しの頻度、フローを実行する曜日が含まれます。":::

1. **[作成]** を選択します。

1. **[新しい手順]** を選択します。

1. **[標準]** タブを選択し、**Excel Online (ビジネス)** を選択します。

    :::image type="content" source="../images/power-automate-tutorial-4.png" alt-text="Power Automate の [Excel Online (Business)] オプション。":::

1. **[アクション]** で、**[スクリプトの実行]** を選択します。

    :::image type="content" source="../images/power-automate-tutorial-5.png" alt-text="Power Automate の [スクリプトの実行] アクションのオプション":::

1. 次に、フロー ステップで使用するブックとスクリプトを選択します。 自分の OneDrive で作成したブック **on-call-rotation.xlsx** を使用します。 **スクリプトの実行** コネクタには、次の設定を指定します。

    - **場所**: OneDrive for Business
    - **ドキュメント ライブラリ**: OneDrive
    - **ファイル**: on-call-rotation.xlsx *(ファイル ブラウザーを使用して選択されています)*
    - **スクリプト**: Get On-Call Person

    :::image type="content" source="../images/power-automate-return-tutorial-4.png" alt-text="スクリプトを実行するための Power Automate コネクタの設定。":::

1. **[新しい手順]** を選択します。

1. リマインダー メールを送信してフローを終了します。 コネクタの検索バーを使用して、**[メールの送信 (V2)]** を選択します。 スクリプトによって返されるメール アドレスを追加するために、**動的なコンテンツの追加** コントロールを使用します。 これは、**result** というラベル付きの Excel アイコンで示されます。 件名、本文は自由に入力できます。

    :::image type="content" source="../images/power-automate-return-tutorial-5.png" alt-text="メールを送信するための Power Automate Outlook コネクタの設定。オプションには、送信するファイル、メールの件名、メールの本文、および詳細オプションが含まれます。":::

    > [!NOTE]
    > このチュートリアルでは、Outlook を使用します。代わりに、お好きなメール サービスを自由に使用することもできますが、一部のオプションは異なる場合があります。

1. **[保存]** を選択します。

## <a name="test-the-script-in-power-automate"></a>Power Automate でスクリプトをテストする

作成したフローは毎週月曜日に実行されます。 画面の右上隅にある **[テスト]** ボタンを選択すると、スクリプトをテストできます。 **[手動]** を選択し、**[テストの実行]** を選択して直ちにフローを実行し、動作をテストします。 続行するには、Excel と Outlook にアクセス許可を付与する必要がある場合があります。

:::image type="content" source="../images/power-automate-return-tutorial-6.png" alt-text="Power Automate の [テスト] ボタン。":::

> [!TIP]
> フローでメールを送信できない場合は、スプレッドシートで、有効なメールが現在の日付範囲用としてテーブルの先頭にリストされていることを再確認してください。

## <a name="next-steps"></a>次の手順

Office スクリプトを Power Automate に接続する方法に関する詳細については、 [「Power Automate で Office スクリプトを実行する」](../develop/power-automate-integration.md)を参照してください。

[「自動タスク リマインダーのサンプル シナリオ」](../resources/scenarios/task-reminders.md)では、Office スクリプトと Power Automate を Teams アダプティブ カードと組み合わせる方法についても説明します。
