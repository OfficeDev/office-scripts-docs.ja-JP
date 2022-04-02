---
title: ブックを 1 つのブックに結合する
description: 他のブックから 1 OfficeブックにPower Automateワークシートを作成する方法について説明します。
ms.date: 09/17/2021
ms.localizationpriority: medium
ms.openlocfilehash: f90980f2e2d1f125f4ca2ffb80822f13ecdeed0e
ms.sourcegitcommit: 7023b9e23499806901a5ecf8ebc460b76887cca6
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 03/31/2022
ms.locfileid: "64585885"
---
# <a name="combine-worksheets-into-a-single-workbook"></a>ワークシートを 1 つのブックに結合する

このサンプルでは、複数のブックから 1 つの集中型ブックにデータをプルする方法を示します。 2 つのスクリプトを使用します。1 つはブックから情報を取得し、もう 1 つはその情報を含む新しいワークシートを作成します。 このスクリプトは、Power Automateフォルダー全体に作用するスクリプトをOneDriveします。

> [!IMPORTANT]
> このサンプルでは、他のブックの値のみをコピーします。 書式設定、グラフ、表、その他のオブジェクトは保持されません。

## <a name="scenario"></a>シナリオ

1. 新しいスクリプト ファイルExcel作成し、OneDriveサンプルから 2 つのスクリプトを追加します。
1. フォルダーをフォルダーに作成OneDriveデータを含む 1 つ以上のブックを追加します。
1. フローを作成して、そのフォルダーのすべてのファイルを取得します。
1. 各ブック **内のすべてのワークシートから** データを取得するには、[ワークシートデータの取得] スクリプトを使用します。
1. [ワークシート **の追加] スクリプトを** 使用して、他のすべてのファイルのすべてのワークシートに対して 1 つのブックに新しいワークシートを作成します。

## <a name="sample-code-return-worksheet-data"></a>サンプル コード: ワークシート のデータを返す

```TypeScript
/**
 * This script returns the values from the used ranges on each worksheet.
 */
function main(workbook: ExcelScript.Workbook): WorksheetData[]
{
  // Create an object to return the data from each worksheet.
  let worksheetInformation: WorksheetData[] = [];

  // Get the data from every worksheet, one at a time.
  workbook.getWorksheets().forEach((sheet) => {
    let values = sheet.getUsedRange()?.getValues();
    worksheetInformation.push({
       name: sheet.getName(),
       data: values as string[][]
    });
  });

  return worksheetInformation;
}

// An interface to pass the worksheet name and cell values through a flow.
interface WorksheetData {
  name: string;
  data: string[][];
}
```

## <a name="sample-code-add-worksheets"></a>サンプル コード: ワークシートの追加

```TypeScript
/**
 * This script creates a new worksheet in the current workbook for each WorksheetData object provided.
 */
function main(workbook: ExcelScript.Workbook, workbookName: string, worksheetInformation: WorksheetData[])
{
  // Add each new worksheet.
  worksheetInformation.forEach((value) => {
    let sheet = workbook.addWorksheet(`${workbookName}.${value.name}`);

    // If there was any data in the worksheet, add it to a new range.
    if (value.data) {
      let range = sheet.getRangeByIndexes(0, 0, value.data.length, value.data[0].length);
      range.setValues(value.data);
    }
  });
}

// An interface to pass the worksheet name and cell values through a flow.
interface WorksheetData {
  name: string;
  data: string[][];
}
```

## <a name="power-automate-flow-combine-worksheets-into-a-single-workbook"></a>Power Automateフロー: ワークシートを 1 つのブックに結合する

1. 新しいインスタント [Power Automate](https://flow.microsoft.com)にサインインして、**新しいインスタント クラウド フローを作成します**。
1. [フロー **を手動でトリガーする] を選択し、[** 作成] を **選択します**。
1. フォルダー内のすべてのファイルを取得します。 この例では、"output" という名前のフォルダーを使用します。 [フォルダー内 **のファイルの一****覧表示] OneDrive for Business** を使用する新 **しい手順を追加** します。 フォルダー ファイルを含むフォルダー パス.csvします。
    * **フォルダー**: /output

    :::image type="content" source="../../images/combine-worksheets-flow-1.png" alt-text="完了したOneDrive for BusinessコネクタをPower Automate。":::
1. ワークシートの **戻り値のデータ** スクリプトを実行して、各ブックからすべてのデータを取得します。 [スクリプトの **Excel] アクションを使用して、オンライン (Business)** コネクタ **を追加** します。 アクションには、次の値を使用します。 ファイルの *ID* を追加すると、Power Automate は各コントロールに **適用** でアクションをラップし、すべてのファイルに対してアクションが実行されます。
    * **場所**: OneDrive for Business
    * **ドキュメント ライブラリ**: OneDrive
    * **ファイル**: *Id* (フォルダー内のリスト ファイル **からの動的コンテンツ**)
    * **スクリプト**: ワークシート のデータを返す
1. 作成した **新しいワークシート** ファイルでワークシートExcel実行します。 これにより、他のすべてのブックのデータが追加されます。 前の **[スクリプトの実行]** アクションと  [各コントロールに適用] の内側に、[スクリプトの実行] アクションExcel **オンライン (Business)** コネクタ **を追加** します。 アクションには、次の値を使用します。
    * **場所**: OneDrive for Business
    * **ドキュメント ライブラリ**: OneDrive
    * **ファイル**: ファイル
    * **スクリプト**: ワークシートの追加
    * **workbookName**: *Name* (フォルダー内のリスト ファイル **からの動的コンテンツ**)
    * **worksheetInformation** ([配列全体の入力に切り替える] ボタンを選択した後、次の画像に続くメモを参照してください):*結果* (Run **スクリプト** からの動的コンテンツ)

    :::image type="content" source="../../images/combine-worksheets-flow-2.png" alt-text="各コントロールに適用する 2 つのスクリプトアクションを実行します。":::
    > [!NOTE]
    > 配列の **個々の項目ではなく** 、配列オブジェクトを直接追加するには、[配列全体を入力する切り替え] ボタンを選択します。
    >
    > :::image type="content" source="../../images/combine-worksheets-flow-3.png" alt-text="コントロール フィールド入力ボックスに配列全体を入力するために切り替えるボタン。":::
1. フローを保存します。 [フロー エディター **] ページ** の [テスト] ボタンを使用するか、[マイ フロー] タブでフロー **を実行** します。メッセージが表示されたら、必ずアクセスを許可してください。
1. これでExcelワークシートが作成されます。
