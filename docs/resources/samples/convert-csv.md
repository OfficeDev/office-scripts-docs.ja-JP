---
title: CSV ファイルをブックにExcelする
description: スクリプトとスクリプトを使用してOfficeファイルPower Automateファイル.xlsx作成する.csvします。
ms.date: 03/28/2022
ms.localizationpriority: medium
ms.openlocfilehash: 52619c1867b654fae3fce1a383a612f81f80d868
ms.sourcegitcommit: 7023b9e23499806901a5ecf8ebc460b76887cca6
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 03/31/2022
ms.locfileid: "64585591"
---
# <a name="convert-csv-files-to-excel-workbooks"></a>CSV ファイルをブックにExcelする

多くのサービスは、データをコンマ区切り値 (CSV) ファイルとしてエクスポートします。 このソリューションでは、これらの CSV ファイルを、Excel形式のブック.xlsx自動化します。 Power Automate フロー[を](https://flow.microsoft.com)使用して、OneDrive フォルダー内の .csv 拡張子を持つファイルと Office スクリプトを使用して、.csv ファイルから新しい Excel ブックにデータをコピーします。

## <a name="solution"></a>ソリューション

1. 新しい.csvと空の "Template" ファイルを.xlsxフォルダーにOneDriveします。
1. CSV データOfficeを解析するスクリプトを作成します。
1. ファイルをPower Automateし、その内容をスクリプトに渡.csvフローを作成します。

## <a name="sample-files"></a>サンプル ファイル

ダウンロード <a href="https://github.com/OfficeDev/office-scripts-docs/blob/master/docs/resources/samples/convert-csv-example.zip?raw=true">convert-csv-example.zip</a> ファイルと 2 つのサンプル Template.xlsxファイルを取得.csvします。 フォルダー内のフォルダーにファイルを抽出OneDrive。 このサンプルでは、フォルダーの名前が "output" と見なされます。

次のスクリプトを追加し、サンプルを自分で試す手順を使用してフローを作成します。

## <a name="sample-code-insert-comma-separated-values-into-a-workbook"></a>サンプル コード: ブックにコンマ区切りの値を挿入する

```TypeScript
/**
 * Convert incoming CSV data into a range and add it to the workbook.
 */
function main(workbook: ExcelScript.Workbook, csv: string) {
  let sheet = workbook.getWorksheet("Sheet1");

  // Remove any Windows \r characters.
  csv = csv.replace(/\r/g, "");

  // Split each line into a row.
  let rows = csv.split("\n");
  /*
   * For each row, match the comma-separated sections.
   * For more information on how to use regular expressions to parse CSV files,
   * see this Stack Overflow post: https://stackoverflow.com/a/48806378/9227753
   */
  const csvMatchRegex = /(?:,|\n|^)("(?:(?:"")*[^"]*)*"|[^",\n]*|(?:\n|$))/g
  rows.forEach((value, index) => {
    if (value.length > 0) {
        let row = value.match(csvMatchRegex);
    
        // Check for blanks at the start of the row.
        if (row[0].charAt(0) === ',') {
          row.unshift("");
        }
    
        // Remove the preceding comma.
        row.forEach((cell, index) => {
          row[index] = cell.indexOf(",") === 0 ? cell.substr(1) : cell;
        });
    
        // Create a 2D array with one row.
        let data: string[][] = [];
        data.push(row);
    
        // Put the data in the worksheet.
        let range = sheet.getRangeByIndexes(index, 0, 1, data[0].length);
        range.setValues(data);
    }
  });

  // Add any formatting or table creation that you want.
}
```

## <a name="power-automate-flow-create-new-xlsx-files"></a>Power Automateフロー: 新しいファイルファイル.xlsxする

1. 新しいクラウド [Power Automate](https://flow.microsoft.com)にサインインし、新しい **スケジュールされたクラウド フローを作成します**。
1. フローを [1 **日] ごとに** 繰り返しに設定し、[作成] を選択 **します**。
1. テンプレート ファイルを取得Excelします。 これは、すべての変換されたファイルの.csvです。 [ファイル コンテンツ **の取得]** **アクションと [** OneDrive for Businessを使用する新しい **手順を追加** します。 "Template.xlsx" ファイルへのファイル パスを指定します。
    * **ファイル**: /output/Template.xlsx
1. [ **ファイル コンテンツ** の取得] 手順の名前を変更するには、その手順の [ファイルコンテンツの取得] **メニュー (....)** (コネクタの右上隅) に移動し、[名前の変更] オプション **を** 選択します。 手順名を "Get Excel" に変更します。

     :::image type="content" source="../../images/convert-csv-flow-1.png" alt-text="完了したOneDrive for BusinessコネクタPower Automate、名前を [Get Excel] に変更しました。":::
1. "出力" フォルダー内のすべてのファイルを取得します。 [フォルダー内 **のファイルの一****覧表示] OneDrive for Business** を使用する新 **しい手順を追加** します。 フォルダー ファイルを含むフォルダー パス.csvします。
    * **フォルダー**: /output

    :::image type="content" source="../../images/convert-csv-flow-2.png" alt-text="完了したOneDrive for BusinessコネクタをPower Automate。":::
1. フローが一部のファイルでのみ動作.csvします。 Condition コントロール **である新しい** ステップ **を追加** します。 Condition には次の値を使用 **します**。
    * **[名前] (***フォルダー内のリスト* ファイルの動的 **コンテンツ) の値を選択します**。 この動的コンテンツには複数の結果が含まれるので、 **各値に適用** する *コントロールは Condition* を囲む点に注意 **してください**。
    * **で終** わる (ドロップダウン リストから)
    * **値を選択する**: .csv

    :::image type="content" source="../../images/convert-csv-flow-3.png" alt-text="完了した Condition コントロールとその周囲の各コントロールに適用します。":::
1. フローの残りの部分は[ **は** い] セクションの下に表示されます。これは、ファイルの処理のみを行.csvです。 新しい手順を.csvコネクタと [ファイル コンテンツの取得] アクションを使用して、OneDrive for Businessファイル **を取得** します。 フォルダー内 **のリスト** ファイルの動的コンテンツ **の ID を使用します**。
    * **ファイル**: *Id* (フォルダーステップの **リスト ファイルからの動的** コンテンツ)
1. 新しいファイル コンテンツ **の取得手順の名前を** "Get .csv" に変更します。 これにより、このファイルをテンプレートからExcelできます。
1. 基本コンテンツとして .xlsxテンプレートを使用してExcelファイルを作成します。 [ファイルの **作成] コネクタと** [ファイル **OneDrive for Businessを使用** する新しい **手順を追加** します。 次の値を使用します。
    * **フォルダー パス**: /output
    * **ファイル名**: *拡張子* のない名前.xlsx (フォルダー内のリスト ファイルから拡張動的コンテンツのない名前を選択し、その後に手動で ".xlsx" と入力します)
    * **ファイル コンテンツ**: *ファイル コンテンツ* (Get Excel **テンプレートから動的コンテンツ**)

     :::image type="content" source="../../images/convert-csv-flow-4.png" alt-text="[ファイルの.csv] および [ファイルの作成] の手順は、Power Automateフローです。":::
1. スクリプトを実行して、新しいブックにデータをコピーします。 [スクリプトの **Excel] アクションを使用して、オンライン (Business)** コネクタ **を追加** します。 アクションには、次の値を使用します。
    * **場所**: OneDrive for Business
    * **ドキュメント ライブラリ**: OneDrive
    * **ファイル**: *Id* (ファイルの作成から動的 **コンテンツ**)
    * **スクリプト**: CSV の変換
    * **csv**: *ファイル コンテンツ* (Get .csv **ファイルから動的コンテンツ**)

    :::image type="content" source="../../images/convert-csv-flow-5.png" alt-text="オンライン (Excel) コネクタの完成Power Automate。":::
1. フローを保存します。 [フロー エディター **] ページ** の [テスト] ボタンを使用するか、[マイ フロー] タブでフロー **を実行** します。メッセージが表示されたら、必ずアクセスを許可してください。
1. "output" フォルダー.xlsx、元のファイルと一緒に新しいファイル.csvがあります。 新しいブックには、CSV ファイルと同じデータが含まれています。

## <a name="troubleshooting"></a>トラブルシューティング

### <a name="script-testing"></a>スクリプトのテスト

スクリプトを使用せずにテストするには、Power Automate前に値`csv`を割り当てる必要があります。 関数の最初の行として次のコードを追加し `main` 、Run キーを押 **してみてください**。

```TypeScript
  csv = `1, 2, 3
         4, 5, 6
         7, 8, 9`;
```

### <a name="semicolon-separated-files-and-other-alternative-separators"></a>セミコロンで区切られたファイルと他の代替の区切り記号

一部の地域ではセミコロンを使用して(';')、コンマではなくセル値を区切ります。 この場合、スクリプト内の次の行を変更する必要があります。

1. 正規表現ステートメントのコンマをセミコロンに置き換える。 これはで始まります `let row = value.match`。

    ```TypeScript
    let row = value.match(/(?:;|\n|^)("(?:(?:"")*[^"]*)*"|[^";\n]*|(?:\n|$))/g);
    ```

1. 空白の最初のセルのチェックで、コンマをセミコロンに置き換える。 これはで始まります `if (row[0].charAt(0)`。

    ```TypeScript
    if (row[0].charAt(0) === ';') {
    ```

1. 表示されるテキストから区切り文字を削除する行のコンマをセミコロンに置き換える。 これはで始まります `row[index] = cell.indexOf`。

   ```TypeScript
      row[index] = cell.indexOf(";") === 0 ? cell.substr(1) : cell;
    ```

> [!NOTE]
> ファイルでタブなどの文字`;``\t`を使用して値を分離する場合は、上記の置換内の文字を使用するか、使用されている文字に置き換える必要があります。

### <a name="large-csv-files"></a>大きな CSV ファイル

ファイルに数十万個のセルがある場合は、データ転送の制限Excel[に達する可能性があります](../../testing/platform-limits.md#excel)。 スクリプトと定期的に同期するスクリプトExcel必要があります。 これを行う最も簡単な方法は、行の `console.log` バッチが処理された後に呼び出す方法です。 これを実行するには、次のコード行を追加します。

1. 前に `rows.forEach((value, index) => {`、次の行を追加します。

    ```TypeScript
      let rowCount = 0;
    ```

1. 後 `range.setValues(data);`に、次のコードを追加します。 列の数によっては、小 `5000` さい数値に減らす必要がある場合があります。

    ```TypeScript
      rowCount++;
      if (rowCount % 5000 === 0) {
        console.log("Syncing 5000 rows.");
      }
    ```

> [!WARNING]
> CSV ファイルが非常に大きい場合は、ファイルのサイズが非常に大きい場合に、[Power Automate。](../../testing/platform-limits.md#power-automate) CSV データを別のブックに変換する前に、CSV データを複数のExcelがあります。
