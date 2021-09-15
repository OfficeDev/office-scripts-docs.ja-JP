---
title: 手動 Power Automation フローからスクリプトを呼び出す
description: Power Automate の Office スクリプトで、手動のトリガーを使う方法を説明します。
ms.date: 06/29/2021
ms.localizationpriority: high
ms.openlocfilehash: 506481c8b5ee1ae94a4e0a7fc926abc62ba7c5f9
ms.sourcegitcommit: d3ed4bdeeba805d97c930394e172e8306a0cf484
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 09/15/2021
ms.locfileid: "59330864"
---
# <a name="call-scripts-from-a-manual-power-automate-flow"></a>手動 Power Automation フローからスクリプトを呼び出す

このチュートリアルでは、[Power Automate](https://flow.microsoft.com)を使用して、Excel on the web 用の Office スクリプトを実行する方法について説明します。 現在の時刻で 2 つのセルの値を更新するスクリプトを作成します。 次に、このスクリプトを手動でトリガーした Power Automate フローに接続し、Power Automate のボタンを選択したときにいつでもこのスクリプトが実行されるようにします。 基本的なパターンを理解したら、フローを拡大して他のアプリケーションを含めることができ、毎日のワークフローの自動化を進めることが可能です。

> [!TIP]
> Office スクリプトを初めて使用する場合は、チュートリアルの「[Excel on the web で Office スクリプトを記録、編集、作成する](excel-tutorial.md)」から始めることをお勧めします。 [Office スクリプトは TypeScript を使用](../overview/code-editor-environment.md)します。このチュートリアルは、JavaScript や TypeScript について初級から中級レベルの知識を持つユーザーを対象としています。 JavaScript を使い慣れていない場合は、「[Mozilla の JavaScript チュートリアル](https://developer.mozilla.org/docs/Web/JavaScript/Guide/Introduction)」から始めることをお勧めします。

## <a name="prerequisites"></a>前提条件

[!INCLUDE [Tutorial prerequisites](../includes/power-automate-tutorial-prerequisites.md)]

## <a name="prepare-the-workbook"></a>ブックを準備する

Power Automate では、ブック コンポーネントにアクセスするために `Workbook.getActiveWorksheet` などの[相対参照](../testing/power-automate-troubleshooting.md#avoid-relative-references)を使わないようにする必要があります。 したがって、Power Automate が参照できる、名前が統一されたワークブックとワークシートが必要です。

1. **MyWorkbook** という名前の新しいブックを作成します。

2. **MyWorkbook** というワークブック内に、**TutorialWorksheet** という名前のワークシートを作成します。

## <a name="create-an-office-script"></a>オフィス スクリプトを作成する

1. **[オートメーション]** タブに移動して **[すべてのスクリプト]** を選択します。

2. **[新しいスクリプト]** を選択します。

3. 既定のスクリプトを次のスクリプトに置き換えます。 このスクリプトは、**TutorialWorksheet** というワークシートの最初の 2 つのセルに現在の日付と時刻を追加します。

    ```TypeScript
    function main(workbook: ExcelScript.Workbook) {
      // Get the "TutorialWorksheet" worksheet from the workbook.
      let worksheet = workbook.getWorksheet("TutorialWorksheet");

      // Get the cells at A1 and B1.
      let dateRange = worksheet.getRange("A1");
      let timeRange = worksheet.getRange("B1");

      // Get the current date and time using the JavaScript Date object.
      let date = new Date(Date.now());

      // Add the date string to A1.
      dateRange.setValue(date.toLocaleDateString());

      // Add the time string to B1.
      timeRange.setValue(date.toLocaleTimeString());
    }
    ```

4. スクリプトの名前を **[日付と時刻の設定]** に変更します。 スクリプト名を選択して変更します。

5. スクリプトを保存するには **[スクリプトの保存]** を選択します。

## <a name="create-an-automated-workflow-with-power-automate"></a>Power Automate を使用して自動化されたワークフローを作成する

1. [「Power Automate のサイト」](https://flow.microsoft.com)にサインインします。

2. 画面の左側に表示されるメニューで、**[作成]** を選択します。 これにより、新しいワークフローを作成する方法の一覧を表示できます。

    :::image type="content" source="../images/power-automate-tutorial-1.png" alt-text="Power Automate の [作成] ボタン。":::

3. **[白紙から初める]** セクションで、**[インスタント フロー]** を選択します。 これで、手動でアクティベートされたワークフローが作成されます。

    :::image type="content" source="../images/power-automate-tutorial-2.png" alt-text="新しいワークフローを作成するための Power Automate インスタント フロー オプション":::

4. 表示されたダイアログ ウィンドウで、フローの名前を **[フロー名]** テキスト ボックスに入力し、**[フローをトリガーする方法の選択]** のオプションの一覧から **[手動でフローをトリガーする]** を選択し、**[作成]** を選択します。

    :::image type="content" source="../images/power-automate-tutorial-3.png" alt-text="Power Automate の[手動でフローをトリガーする] オプション。":::

    手動でトリガーするフローは、いくつかあるフローの種類のうちの 1 つです。 次のチュートリアルでは、メールを受信したときに自動的に実行されるフローを作成します。

5. **[新しいステップ]** を選択します。

6. **[標準]** タブを選択し、**Excel Online (ビジネス)** を選択します。

    :::image type="content" source="../images/power-automate-tutorial-4.png" alt-text="Power Automate の [Excel Online (Business)] オプション。":::

7. **[アクション]** で、**[スクリプトの実行]** を選択します。

    :::image type="content" source="../images/power-automate-tutorial-5.png" alt-text="Power Automate の [スクリプトの実行] アクションのオプション":::

8. 次に、フロー ステップで使用するブックおよびスクリプトを選択します。 このチュートリアルでは、OneDrive に作成したブックを使用しますが、OneDrive サイトまたは SharePoint サイトでは任意のブックを使用できます。 **スクリプトの実行** コネクタには、次の設定を指定します。

    - **場所**: OneDrive for Business
    - **ドキュメント ライブラリ**: OneDrive
    - **ファイル**: MyWorkbook.xlsx *(ファイル ブラウザーを使用して選択されています)*
    - **スクリプト**: 日時を設定

    :::image type="content" source="../images/power-automate-tutorial-6.png" alt-text="スクリプトを実行するための Power Automate コネクタの設定。":::

9. **[保存]** を選択します。

これで、フローは Power Automate で実行できるようになりました。 フロー エディターの **[テスト]** ボタンを使用してテストするか、チュートリアルの残りの手順に従って、フロー コレクションからフローを実行できます。

## <a name="run-the-script-through-power-automate"></a>Power Automate でスクリプトを実行する

1. Power Automate のメイン ページで、**[自分のフロー]** を選択します。

    :::image type="content" source="../images/power-automate-tutorial-7.png" alt-text="Power Automate の [自分のフロー] ボタン。":::

2. **[自分のフロー]** タブに表示されているフローの一覧から、**[自分のチュートリアル フロー]** を選択すると、以前に作成したフローの詳細が表示されます。

3. **[実行]** を選択します。

    :::image type="content" source="../images/power-automate-tutorial-8.png" alt-text="Power Automate の [実行] ボタン。":::

4. フローを実行するための作業ウィンドウが表示されます。 Excel Online への **サインイン** を要求された場合は、**[続行]** を選択します。

5. **[フローの実行]** を選択します。 これにより、関連する Office スクリプトを実行するフローが実行されます。

6. **[完了]** を選択します。 それに応じて **[実行]** セクションが更新されます。

7. ページを更新して、Power Automate の結果を表示します。 成功した場合は、ワークブックに移動して、更新されたセルを確認します。 エラーが発生した場合は、フローの設定を確認し、もう一度実行します。

    :::image type="content" source="../images/power-automate-tutorial-9.png" alt-text="正常にフローが発生したことを示す Power Automate 出力。":::

## <a name="next-steps"></a>次の手順

[「自動で実行される Power Automate フロー内で、データをスクリプトに渡す」](excel-power-automate-trigger.md)のチュートリアルを完了します。 このコースでは、ワークフロー サービスから Office スクリプトにデータを渡す方法と、特定のイベントが発生したときに Power Automate フローを実行する方法について説明します。
