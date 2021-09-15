---
title: CSV ファイルをブックExcel変換する
description: スクリプトとスクリプトを使用Office、Power Automateファイルから.xlsxファイルを.csvします。
ms.date: 07/19/2021
ms.localizationpriority: medium
ms.openlocfilehash: ecfc4d143cbaf10b9ea5f02881751f2c4fa28853
ms.sourcegitcommit: d3ed4bdeeba805d97c930394e172e8306a0cf484
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 09/15/2021
ms.locfileid: "59333434"
---
# <a name="convert-csv-files-to-excel-workbooks"></a>CSV ファイルをブックExcel変換する

多くのサービスは、データをコンマ区切り値 (CSV) ファイルとしてエクスポートします。 このソリューションは、これらの CSV ファイルを、Excel形式のブック.xlsx自動化します。 Power Automate フロー[を](https://flow.microsoft.com)使用して、OneDrive フォルダー内の .csv 拡張子を持つファイルと Office スクリプトを使用して、.csv ファイルから新しい Excel ブックにデータをコピーします。

## <a name="solution"></a>ソリューション

1. 新しい.csvファイルと空の "Template" ファイルを.xlsxフォルダーにOneDriveします。
1. CSV データOfficeを解析するスクリプトを作成します。
1. ファイルをPower Automateし、その内容をスクリプトに渡.csvフローを作成します。

## <a name="sample-files"></a>サンプル ファイル

この <a href="https://github.com/OfficeDev/office-scripts-docs/blob/master/docs/resources/samples/convert-csv-example.zip?raw=true">convert-csv-example.zipをダウンロード </a> して、Template.xlsxファイルと 2 つのサンプル ファイル.csvします。 フォルダー内のフォルダーにファイルを抽出OneDrive。 このサンプルでは、フォルダーの名前が "output" と見なされます。

次のスクリプトを追加し、サンプルを自分で試す手順を使用してフローを作成します。

## <a name="sample-code-insert-comma-separated-values-into-a-workbook"></a>サンプル コード: ブックにコンマ区切りの値を挿入する

```TypeScript
function main(workbook: ExcelScript.Workbook, csv: string) {
  /* Convert the CSV data into a 2D array. */
  // Trim the trailing new line.
  csv = csv.trim();

  // Split each line into a row.
  let rows = csv.split("\r\n");
  let data : string[][] = [];
  rows.forEach((value) => {
    /*
     * For each row, match the comma-separated sections.
     * For more information on how to use regular expressions to parse CSV files,
     * see this Stack Overflow post: https://stackoverflow.com/a/48806378/9227753
     */
    let row = value.match(/(?:,|\n|^)("(?:(?:"")*[^"]*)*"|[^",\n]*|(?:\n|$))/g);
    
    // Remove the preceding comma.
    row.forEach((cell, index) => {
      row[index] = cell.indexOf(",") === 0 ? cell.substr(1) : cell;
    });
    data.push(row);
  });

  // Put the data in the worksheet.
  let sheet = workbook.getWorksheet("Sheet1");
  let range = sheet.getRangeByIndexes(0, 0, data.length, data[0].length);
  range.setValues(data);

  // Add any formatting or table creation that you want.
}
```

## <a name="power-automate-flow-create-new-xlsx-files"></a>Power Automateフロー: 新しいファイルファイル.xlsxする

1. 新しいクラウド [Power Automate](https://flow.microsoft.com)にサインインし、新しい **スケジュールされたクラウド フローを作成します**。
1. フローを [すべての "1" "日" を繰り返す] に設定し、[作成] を **選択します**。
1. テンプレート ファイルを取得Excelします。 これは、すべての変換されたファイルの.csvです。 [ファイル コンテンツ **の取得]** **アクションと**[OneDrive for Business] コネクタを使用する **新しい手順を追加** します。 "Template.xlsx" ファイルへのファイル パスを指定します。
    * **ファイル**: /output/Template.xlsx
1. (コネクタ **の右上隅** にある) その手順 **の ...** メニューに移動し、[名前の変更] オプションを選択して、ファイル コンテンツの取得手順の名前 **を変更** します。 手順名を "Get Excel" に変更します。

     :::image type="content" source="../../images/convert-csv-flow-1.png" alt-text="完了したOneDrive for Businessコネクタを Power Automate Get Excelに変更しました。":::
1. "出力" フォルダー内のすべてのファイルを取得します。 [フォルダー内 **のファイルの一** 覧] **OneDrive for Businessを使用** する新 **しい手順を追加** します。 フォルダー ファイルを含むフォルダー パス.csvします。
    * **フォルダー**: /output

    :::image type="content" source="../../images/convert-csv-flow-2.png" alt-text="完了したOneDrive for BusinessコネクタをPower Automate。":::
1. フローが一部のファイルでのみ動作.csvします。 Condition コントロール **である新しい** ステップ **を追加** します。 Condition には次の値を **使用します**。
    * **[名前] (***フォルダー内のリスト* ファイルから動的コンテンツ)**の値を選択します**。 この動的コンテンツには複数の結果が含まれるので、 **各** 値に適用する *コントロールは* Condition を囲 **む点** に注意してください。
    * **で終** わる (ドロップダウン リストから)
    * **値を選択する**: .csv

    :::image type="content" source="../../images/convert-csv-flow-3.png" alt-text="完了した Condition コントロールとその周囲の各コントロールに適用します。":::
1. フローの残りの部分は、[ **は** い] セクションの下に表示されます。これは、ファイルの処理のみを行う.csvです。 新しい手順を.csvコネクタを使用し、[ファイルコンテンツの取得]**アクションを** OneDrive for Businessして、個々のファイル ファイル **を取得** します。 フォルダー内 **のリスト** ファイルの動的コンテンツ **の ID を使用します**。
    * **ファイル**: *Id* (フォルダーステップのリスト ファイル **からの動的** コンテンツ)
1. 新しいファイル コンテンツ **の取得手順の名前を** "Get .csv" に変更します。 これにより、このファイルをテンプレートからExcelできます。
1. 基本コンテンツとして.xlsxテンプレートを使用してExcelファイルを作成します。 [ファイルの **作成] コネクタと**[ファイル **OneDrive for Business] アクションを使用する** 新しい **手順を追加** します。 次の値を使用します。
    * **フォルダーパス**: /output
    * **ファイル名**:*拡張子のない名前.xlsx* (フォルダー内のリスト ファイルから拡張動的コンテンツのない名前を選択し、その後に手動で ".xlsx" と入力します)
    * **ファイル コンテンツ**:*ファイル コンテンツ*(Get **Excel テンプレートから動的コンテンツ**)

     :::image type="content" source="../../images/convert-csv-flow-4.png" alt-text="[ファイルの.csv] および [ファイルの作成] の手順は、Power Automateフローです。":::
1. スクリプトを実行して、新しいブックにデータをコピーします。 [スクリプトの **Excel] アクションを使用して、オンライン (Business)** コネクタ **を追加** します。 アクションには、次の値を使用します。
    * **場所**: OneDrive for Business
    * **ドキュメント ライブラリ**: OneDrive
    * **ファイル**: *Id* (ファイルの作成から動的 **コンテンツ**)
    * **スクリプト**: CSV の変換
    * **csv**: *ファイル コンテンツ* (Get **.csv ファイルから動的コンテンツ**)

    :::image type="content" source="../../images/convert-csv-flow-5.png" alt-text="オンライン (Excel) コネクタの完成Power Automate。":::
1. フローを保存します。 [フロー エディター **] ページ** の [テスト] ボタンを使用するか、[マイ フロー] タブでフロー **を実行** します。メッセージが表示されたら、必ずアクセスを許可してください。
1. "output" フォルダー.xlsx、元のファイルと一緒に新しいファイル.csvがあります。 新しいブックには、CSV ファイルと同じデータが含まれています。
