---
title: Power Automate でスクリプトの使用を開始する
description: パワーで Office スクリプトを使用する方法についてのチュートリアルは、手動のトリガーを使用して自動化します。
ms.date: 07/01/2020
localization_priority: Priority
ms.openlocfilehash: 83e072a45fc724ff2aac5bf5f3893dcb64eaf2ff
ms.sourcegitcommit: edf58aed3cd38f57e5e7227465a1ef5515e15703
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 07/08/2020
ms.locfileid: "45081636"
---
# <a name="start-using-scripts-with-power-automate-preview"></a>Power 自動でのスクリプトの使用を開始する (プレビュー)

このチュートリアルでは、 [Power オートメーション](https://flow.microsoft.com)を使用して web 上で Excel の Office スクリプトを実行する方法について説明します。

## <a name="prerequisites"></a>前提条件

[!INCLUDE [Tutorial prerequisites](../includes/tutorial-prerequisites.md)]

> [!IMPORTANT]
> このチュートリアルでは、「web チュートリアルでの[Excel での Office スクリプトの記録、編集、および作成」](excel-tutorial.md)を完了していることを前提としています。

## <a name="prepare-the-workbook"></a>ブックの準備

Power オートメーションは `Workbook.getActiveWorksheet` 、ブックコンポーネントへのアクセスなどの相対参照を使用できません。 そのため、Power オートメーションが参照できる、一貫した名前を持つブックとワークシートが必要です。

1. **Myworkbook**という名前の新しいブックを作成します。

2. **Myworkbook**ブックで、 **TutorialWorksheet**という名前のワークシートを作成します。

## <a name="create-an-office-script"></a>Office スクリプトを作成する

1. [**自動化**] タブに移動して、[**コードエディター**] を選択します。

2. [**新しいスクリプト**] を選択します。

3. 既定のスクリプトを次のスクリプトに置き換えます。 このスクリプトは、 **TutorialWorksheet**ワークシートの最初の2つのセルに現在の日付と時刻を追加します。

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

4. スクリプトの名前を変更し**て、日付と時刻を設定**します。 スクリプト名を押して変更します。

5. [**保存スクリプト**を押してスクリプトを保存します。

## <a name="create-an-automated-workflow-with-power-automate"></a>Power 自動化を使用して自動化されたワークフローを作成する

1. [パワー自動化プレビューサイト](https://flow.microsoft.com)にサインインします。

2. 画面の左側に表示されるメニューで、[**作成**] を押します。 これにより、新しいワークフローを作成する方法の一覧が表示されます。

    ![パワー自動化の [作成] ボタン。](../images/power-automate-tutorial-1.png)

3. [**空白から開始**] セクションで、[**インスタントフロー**] を選択します。 これにより、手動でアクティブ化したワークフローが作成されます。

    ![新しいワークフローを作成するためのインスタントフローオプション。](../images/power-automate-tutorial-2.png)

4. 表示されるダイアログウィンドウで、[**フロー名**] テキストボックスにフローの名前を入力し、[フロー**の開始方法を選択**してください] で、オプションの一覧から [**フローを手動でトリガー**する] を選択して、[**作成**] をクリックします。

    ![新しいインスタントフローを作成するための手動トリガーオプション。](../images/power-automate-tutorial-3.png)

5. **新しい手順**を押します。

6. [**標準**] タブを選択し、[ **Excel Online (Business)**] を選択します。

    ![Excel Online (Business) の電源自動化オプション。](../images/power-automate-tutorial-4.png)

7. [**アクション**] で、[**スクリプトを実行する (プレビュー)**] を選択します。

    ![実行スクリプトのパワー自動処理オプション (プレビュー)。](../images/power-automate-tutorial-5.png)

8. **実行スクリプト**コネクタについて、次の設定を指定します。

    - **場所**: OneDrive for business
    - **ドキュメントライブラリ**: OneDrive
    - **ファイル**: MyWorkbook.xlsx
    - **スクリプト**: 日付と時刻を設定する

    ![パワー自動化でスクリプトを実行するためのコネクタの設定。](../images/power-automate-tutorial-6.png)

9. [**保存**します。

これで、電力の自動化を通じてフローを実行する準備が整いました。 フローエディターの [**テスト**] ボタンを使用してテストするか、チュートリアルの残りの手順に従ってフローコレクションからフローを実行することができます。

## <a name="run-the-script-through-power-automate"></a>電源自動化を使用してスクリプトを実行する

1. [メインパワーの自動化] ページで、[**マイフロー**] を選択します。

    ![パワー自動化の [マイフロー] ボタン。](../images/power-automate-tutorial-7.png)

2. [ **My** flow] タブに表示されるフローの一覧から [ **my チュートリアルフロー** ] を選択します。これで、以前に作成したフローの詳細が表示されます。

3. **Run**を押します。

    ![電源自動化の [実行] ボタン。](../images/power-automate-tutorial-8.png)

4. フローを実行するための作業ウィンドウが表示されます。 Excel Online に**サインイン**するように求めるメッセージが表示されたら、[**続行**] を押します。

5. **Run flow**を押します。 これにより、関連する Office スクリプトが実行されるフローが実行されます。

6. [**完了**します。 それに応じて、「**実行**」セクションの更新が表示されます。

7. ページを更新して、電力自動化の結果を表示します。 成功した場合は、ブックに移動して、更新されたセルを表示します。 失敗した場合は、フローの設定を確認し、2回目に実行します。

    ![フローが正常に実行されたことを示す電力を自動で出力します。](../images/power-automate-tutorial-9.png)

## <a name="next-steps"></a>次の手順

[自動電源自動化フローに関するチュートリアルを使用して、自動実行スクリプトを](excel-power-automate-trigger.md)完了します。 この章では、ワークフローサービスから Office スクリプトにデータを渡す方法について説明します。
