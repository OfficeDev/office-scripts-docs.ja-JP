---
title: Excel on the web で Office スクリプトを記録、編集、作成する
description: 操作レコーダーを使用したスクリプトの記録、ブックへのデータの書き込みなど、Office スクリプトの基本について説明したチュートリアル。
ms.date: 01/06/2021
localization_priority: Priority
ms.openlocfilehash: ae864cc08453a9c8a2538f15ceee1275e131725d
ms.sourcegitcommit: 45ffe3dbd2c834b78592ad35928cf8096f5e80bc
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 04/14/2021
ms.locfileid: "51754846"
---
# <a name="record-edit-and-create-office-scripts-in-excel-on-the-web"></a>Excel on the web で Office スクリプトを記録、編集、作成する

このチュートリアルでは、Excel on the web の Office スクリプトの基本となる記録、編集、書き込みについて説明します。 売上記録ワークシートにいくつか書式設定を適用するスクリプトを記録します。 記録されたスクリプトを編集して、より多くの書式設定を適用し、テーブルを作成して、そのテーブルを並べ替えます。 記録して編集するこのパターンは、Excel のアクションがコードとしてどのように表示されるか確認するための重要なツールです。

## <a name="prerequisites"></a>前提条件

[!INCLUDE [Tutorial prerequisites](../includes/tutorial-prerequisites.md)]

> [!IMPORTANT]
> このチュートリアルは、JavaScript や TypeScript について初級から中級レベルの知識を持つユーザーを対象としています。 JavaScript を使い慣れていない場合は、「[Mozilla の JavaScript チュートリアル](https://developer.mozilla.org/docs/Web/JavaScript/Guide/Introduction)」から始めることをお勧めします。 スクリプト環境の詳細については、「[Office スクリプト コード エディターの環境](../overview/code-editor-environment.md)」を参照してください。

## <a name="add-data-and-record-a-basic-script"></a>データを追加し、基本スクリプトを記録する

まず、いくらかのデータと、最初の小さなスクリプトが必要です。

1. Excel for the Web で新しいブックを作成します。
2. 次の果物売上データをコピーし、ワークシートのセル **A1** から始まるセル範囲に貼り付けます。

    |果物 |2018 |2019 |
    |:---|:---|:---|
    |オレンジ |1000 |1200 |
    |レモン |800 |900 |
    |ライム |600 |500 |
    |グレープフルーツ |900 |700 |

3. **[自動化]** タブを開きます。**[自動化]** タブが表示されていない場合は、ドロップダウン矢印を押して、リボンのオーバーフローを確認します。
4. **[操作を記録する]** ボタンを押します。
5. セル **A2:C2** ("オレンジ" 行) を選択し、塗りつぶしの色をオレンジ色に設定します。
6. **[停止]** ボタンを押して、記録を停止します。
7. **[スクリプト名]** フィールドに覚えやすい名前を入力します。
8. *オプション:* **[説明]** フィールドにわかりやすい説明を入力します。 このフィールドは、スクリプトの動作に関するコンテキストを提供するために使用します。 このチュートリアルでは、「テーブルの色コード行」を使用できます。

   > [!TIP]
   > スクリプトの説明は、**[スクリプトの詳細]** ウィンドウで後から編集できます。これは、コード エディターの **[...]** メニューの下にあります。

9. **[保存]** ボタンを押して、スクリプトを保存します。

    ワークシートは次のようになります (色が違っていても問題ありません)。

    :::image type="content" source="../images/tutorial-1.png" alt-text="&quot;オレンジ&quot; を含む行がオレンジ色で強調表示された、フルーツの売上データ行を示すワークシート。":::

## <a name="edit-an-existing-script"></a>既存のスクリプトを編集する

前のスクリプトでは、"オレンジ" の行がオレンジ色になります。 "レモン" の行に黄色を追加しましょう。

1. [**詳細**] ウィンドウを開き、[**編集**] ボタンを押します。
2. 次のようなコードが表示されるはずです。

    ```TypeScript
    function main(workbook: ExcelScript.Workbook) {
      // Set fill color to FFC000 for range Sheet1!A2:C2
      let selectedSheet = workbook.getActiveWorksheet();
      selectedSheet.getRange("A2:C2").getFormat().getFill().setColor("FFC000");
    }
    ```

    このコードは、ブックから現在のワークシートを取得します。 次に、**A2:C2** の範囲の塗りつぶしの色を設定します。

    範囲は、Excel on the web の Office スクリプトの基本となる部分です。 範囲とは、隣接するセルからなる四角形のブロックで、値、数式、書式設定が含まれます。 範囲はセルの基本構造であり、スクリプト タスクの大部分は範囲を指定することにより実行します。

3. 次の行をスクリプトの最後 (`color` の設定箇所と末尾の `}` の間) に追加します。

    ```TypeScript
    selectedSheet.getRange("A3:C3").getFormat().getFill().setColor("yellow");
    ```

4. **[実行]** を押して、スクリプトをテストします。 ブックは次のように表示されるはずです。

    :::image type="content" source="../images/tutorial-2.png" alt-text="&quot;オレンジ&quot; の行はオレンジ色、&quot;レモン&quot; の行は黄色で強調表示されている果物売上データの行を示すワークシート。":::

## <a name="create-a-table"></a>テーブルを作成する

この果物売上データをテーブルに変換しましょう。 プロセス全体でスクリプトを使用します。

1. 次の行をスクリプトの最後 (末尾の `}` の前) に追加します。

    ```TypeScript
    let table = selectedSheet.addTable("A1:C5", true);
    ```

2. この呼び出しは `Table` オブジェクトを返します。 そのテーブルを使用して、データを並べ替えましょう。 "果物" 列の値に基づいて、データを昇順で並べ替えます。 次の行を、テーブル作成の後に追加します。

    ```TypeScript
    table.getSort().apply([{ key: 0, ascending: true }]);
    ```

    スクリプトは次のようになります。

    ```TypeScript
    function main(workbook: ExcelScript.Workbook) {
        // Set fill color to FFC000 for range Sheet1!A2:C2
        let selectedSheet = workbook.getActiveWorksheet();
        selectedSheet.getRange("A2:C2").getFormat().getFill().setColor("FFC000");
        selectedSheet.getRange("A3:C3").getFormat().getFill().setColor("yellow");
        let table = selectedSheet.addTable("A1:C5", true);
        table.getSort().apply([{ key: 0, ascending: true }]);
    }
    ```

    テーブルには `TableSort` オブジェクトがあり、`Table.getSort` メソッドを使用してアクセスできます。 そのオブジェクトに並べ替え条件を適用できます。 `apply` メソッドは、`SortField` オブジェクトの配列を受け取ります。 今回は、並べ替え条件が 1 つだけなので、`SortField` を 1 つだけ使用します。 `key: 0` は、並べ替えを定義する値を含む列を "0" (テーブルの 1 列目。この例では **A**) に設定します。 `ascending: true` は、昇順 (降順ではなく) にデータを並べ替えます。

3. スクリプトを実行します。 テーブルが次のように表示されます。

    :::image type="content" source="../images/tutorial-3.png" alt-text="並べ替えされたフルーツの販売テーブルを示すワークシート。":::

    > [!NOTE]
    > スクリプトを再実行すると、エラーが表示されます。 これは、テーブルの上に別のテーブルを重ねて作成することはできないためです。 ただし、別のワークシートやブックでスクリプトを実行することはできます。

### <a name="re-run-the-script"></a>スクリプトを再実行する

1. 現在のブックに新しいワークシートを作成します。
2. このチュートリアルの最初にある果物のデータをコピーし、新しいワークシートのセル **A1** から始まるセル範囲に貼り付けます。
3. スクリプトを実行します。

## <a name="next-steps"></a>次の手順

チュートリアルの「[Excel on the web で Office スクリプトを使用してブックのデータを読み取る](excel-read-tutorial.md)」を完了します。 このチュートリアルでは、Office スクリプトを使用してブックのデータを読み取る方法について説明します。
