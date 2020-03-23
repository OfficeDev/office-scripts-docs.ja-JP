---
title: Excel on the web で Office スクリプトを使用してブックのデータを読み取る
description: ブックのデータを読み取り、スクリプトでそのデータを評価する方法について説明した Office スクリプトのチュートリアル。
ms.date: 01/27/2020
localization_priority: Priority
ms.openlocfilehash: 42ed0fe5843a78692f9660b873211e3668702164
ms.sourcegitcommit: b075eed5a6f275274fbbf6d62633219eac416f26
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 03/10/2020
ms.locfileid: "42700323"
---
# <a name="read-workbook-data-with-office-scripts-in-excel-on-the-web"></a>Excel on the web で Office スクリプトを使用してブックのデータを読み取る

このチュートリアルでは、Excel on the web 用の Office スクリプトを使用してブックのデータを読み取る方法について説明します。 その後、読み取ったデータを編集し、ブックに戻します。

> [!TIP]
> Office スクリプトを初めて使用する場合は、チュートリアルの「[Excel on the web で Office スクリプトを記録、編集、作成する](excel-tutorial.md)」から始めることをお勧めします。

## <a name="prerequisites"></a>前提条件

[!INCLUDE [Preview note](../includes/preview-note.md)]

このチュートリアルを開始するには、Office スクリプトへのアクセスが必要です。これには次のものが必要です。

- [Excel on the web](https://www.office.com/launch/excel)。
- [組織に対して Office スクリプトを許可する](https://support.office.com/article/office-scripts-settings-in-m365-19d3c51a-6ca2-40ab-978d-60fa49554dcf)よう管理者に依頼します。これにより、リボンに **[自動化]** タブが追加されます。

> [!IMPORTANT]
> このチュートリアルは、JavaScript や TypeScript について初級から中級レベルの知識を持つユーザーを対象としています。 JavaScript を使い慣れていない場合は、[Mozilla の JavaScript チュートリアル](https://developer.mozilla.org/docs/Web/JavaScript/Guide/Introduction)をご覧になることをお勧めします。 スクリプト環境の詳細については、「[Excel on the web の Office スクリプト](../overview/excel.md)」を参照してください。

## <a name="read-a-cell"></a>セルを読み取る。

操作レコーダーで作成したスクリプトは、ブックに情報を書き込む操作のみを実行できます。 コード エディターを使用すると、ブックのデータを読み取ることも可能なスクリプトの編集と作成ができます。

データを読み取り、読み取った内容に基づいて動作するスクリプトを作成しましょう。 今回は、サンプルの銀行取引明細書を使用します。 この明細書は、支払いと貸方がまとまった明細書です。 残念ながら、残高の変化が異なる仕方で報告されています。 支払い明細では、収入を負の貸方として記録し、支出を負の借方として記録しています。 貸方明細ではその逆になっています。

チュートリアルの残りの部分で、スクリプトを使用してこのデータを正規化します。 まず、ブックからデータを読み取る方法について説明します。

1. チュートリアルの残りの部分で使用したブックに新しいワークシートを作成します。
2. 次のデータをコピーし、新しいワークシートのセル **A1** から始まるセル範囲に貼り付けます。

    |日付 |取引 |説明 |借方 |貸方 |
    |:--|:--|:--|:--|:--|
    |2019/10/10 |支払い |Coho Vineyard |-20.05 | |
    |2019/10/11 |貸方 |The Phone Company |99.95 | |
    |2019/10/13 |貸方 |Coho Vineyard |154.43 | |
    |2019/10/15 |支払い |外部預金 | |1000 |
    |2019/10/20 |貸方 |Coho Vineyard - 返金 | |-35.45 |
    |2019/10/25 |支払い |Best For You Organics Company | -85.64 | |
    |2019/11/01 |支払い |外部預金 | |1000 |

3. **[コード エディター]** を開き、**[新しいスクリプト]** を選択します。
4. 書式設定をクリーンアップします。 これは財務ドキュメントなので、**[借方]** 列と **[貸方]** 列の数値の書式設定を変更して、値がドル金額として表示されるようにします。 さらに、列幅をデータに合わせます。

    スクリプトの内容を次のコードで置き換えます。

    ```TypeScript
    async function main(context: Excel.RequestContext) {
      // Get the current worksheet.
      let workbook = context.workbook;
      let worksheets = workbook.worksheets;
      let selectedSheet = worksheets.getActiveWorksheet();

      // Format the range to display numerical dollar amounts.
      selectedSheet.getRange("D2:E8").numberFormat = [["$#,##0.00"]];

      // Fit the width of all the used columns to the data.
      selectedSheet.getUsedRange().format.autofitColumns();
    }
    ```

5. では、いずれかの数値列の値を読み取ってみましょう。 次のコードをスクリプトの最後に追加します。

    ```TypeScript
    // Get the value of cell D2.
    let range = selectedSheet.getRange("D2");
    range.load("values");
    await context.sync();
  
    // Print the value of D2.
    console.log(range.values);
    ```

    `load` と `sync` への呼び出しに注目してください。 これらのメソッドの詳細については、「[Excel on the web での Office スクリプトのスクリプトの基本事項](../develop/scripting-fundamentals.md#sync-and-load)」で説明します。 ここでは、データの読み取りを要求し、スクリプトとブックを同期してそのデータを読み取る必要があることを覚えておいてください。

6. スクリプトを実行します。
7. コンソールを開きます。 **省略記号**のメニューを選択し、**[Logs...](ログ...)** を押します。
8. コンソールに `[Array[1]]` が表示されます。 範囲は 2 次元のデータ配列であるため、これは数値ではありません。 この 2 次元の範囲は、コンソールに直接ログ記録されます。 コード エディターを使用すると、この配列の内容を表示できます。
9. 2 次元の配列がコンソールにログ記録すると、各行の下に列の値がグループ化されます。 青い三角形を押して、配列のログを展開します。
10. 新たに表示された青い三角形を押して、配列の第 2 レベルを展開します。 次のように表示されるはずです。

    ![出力 "-20.05" が 2 つの配列の下に入れ子になって表示されているコンソール ログ。](../images/tutorial-4.png)

## <a name="modify-the-value-of-a-cell"></a>セルの値を変更する

データを読み取れたので、そのデータを使用してブックを変更しましょう。 セル **D2** の値を、`Math.abs` 関数を使用して正の値にします。 [Math](https://developer.mozilla.org/docs/web/javascript/reference/global_objects/math) オブジェクトには、スクリプトでアクセスできる多くの関数が含まれています。 `Math` および他の組み込みオブジェクトの詳細については、「[Office スクリプトでの組み込みの JavaScript オブジェクトの使用](../develop/javascript-objects.md)」を参照してください。

1. 次のコードをスクリプトの最後に追加します。

    ```TypeScript
    // Run the `Math.abs` function with the value at D2 and apply that value back to D2.
    let positiveValue = Math.abs(range.values[0][0]);
    range.values = [[positiveValue]];
    ```

2. セル **D2** の値が正の値になります。

## <a name="modify-the-values-of-a-column"></a>列の値を変更する

1 つのセルの読み取り方法と書き込み方法がわかったので、スクリプトを一般化して、**[借方]** 列と **[貸方]** 列全体を操作できるようにしましょう。

1. 1 つのセルにのみ影響するコード (前述の絶対値コード) を削除します。すると、スクリプトは次のようになります。

    ```TypeScript
    async function main(context: Excel.RequestContext) {
      // Get the current worksheet.
      let workbook = context.workbook;
      let worksheets = workbook.worksheets;
      let selectedSheet = worksheets.getActiveWorksheet();

      // Format the range to display numerical dollar amounts.
      selectedSheet.getRange("D2:E8").numberFormat = [["$#,##0.00"]];

      // Fit the width of all the used columns to the data.
      selectedSheet.getUsedRange().format.autofitColumns();
    }
    ```

2. 最後の 2 つの列の行を反復処理するループを追加します。 スクリプトにより、各セルの値が現在の値の絶対値に設定されます。

    セルの位置を定義する配列は 0 から始まることにご注意ください。 したがって、セル **A1** は `range[0][0]` になります。

    ```TypeScript
    // Get the values of the used range.
    let range = selectedSheet.getUsedRange();
    range.load("rowCount,values");
    await context.sync();

    // Iterate over the fourth and fifth columns and set their values to their absolute value.
    for (let i = 1; i < range.rowCount; i++) {
      // The column at index 3 is column "4" in the worksheet.
      if (range.values[i][3] != 0) {
        let positiveValue = Math.abs(range.values[i][3]);
        selectedSheet.getCell(i, 3).values = [[positiveValue]];
      }

      // The column at index 4 is column "5" in the worksheet.
      if (range.values[i][4] != 0) {
        let positiveValue = Math.abs(range.values[i][4]);
        selectedSheet.getCell(i, 4).values = [[positiveValue]];
      }
    }
    ```

    スクリプトのこの部分は、いくつかの重要なタスクを実行します。 まず、指定された範囲の値と行数を読み込みます。 これにより、値が表示され、いつ停止すればよいかを確認できます。 次に、指定された範囲を反復処理し、**[借方]** 列と **[貸方]** 列の各セルをチェックします。 最後に、セルの値が 0 ではない場合、その値が絶対値で置き換えられます。 0 は使用しないので、空のセルはそのままにしておきます。

3. スクリプトを実行します。

    銀行取引明細書は次のように表示されるはずです。

    ![書式設定された正の値のみを含むテーブル形式の銀行取引明細書。](../images/tutorial-5.png)

## <a name="next-steps"></a>次の手順

コード エディターを開き、「[Excel on the web での Office スクリプトのサンプル スクリプト](../resources/excel-samples.md)」をいくつか試してみます。 Office スクリプトの作成について詳しくは、「[Excel on the web での Office スクリプトのスクリプトの基本事項](../develop/scripting-fundamentals.md)」も参照してください。
