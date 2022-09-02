---
title: Excel で日々の変更を記録し、Power Automate フローを使用してレポートする
description: Office スクリプトと Power Automate を使用してブック内の値の変更を追跡する方法について説明します
ms.date: 08/25/2022
ms.localizationpriority: medium
ms.openlocfilehash: 083ca08573db060aa4788aea58fc67e50d004a4b
ms.sourcegitcommit: a6504f8b0d6b717457c6e0b5306c35ad3900914e
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 09/02/2022
ms.locfileid: "67572664"
---
# <a name="record-day-to-day-changes-in-excel-and-report-them-with-a-power-automate-flow"></a>Excel で日々の変更を記録し、Power Automate フローを使用してレポートする

Power Automate と Office スクリプトは組み合わせて、繰り返しのタスクを処理します。 このサンプルでは、毎日 1 つの数値読み取り値をブックに記録し、昨日からの変更を報告する必要があります。 その読み取りを取得し、ブックに記録し、電子メールで変更を報告するフローを作成します。

## <a name="sample-excel-file"></a>Excel ファイルのサンプル

すぐに使用できるブックの [daily-readings.xlsx](daily-readings.xlsx) をダウンロードします。 サンプルを自分で試すには、次のスクリプトを追加します。

## <a name="sample-code-record-and-report-daily-readings"></a>サンプル コード: 毎日の測定値を記録して報告する

```TypeScript
function main(workbook: ExcelScript.Workbook, newData: string): string {
  // Get the table by its name.
  const table = workbook.getTable("ReadingTable");

  // Read the current last entry in the Reading column.
  const readingColumn = table.getColumnByName("Reading");
  const readingColumnValues = readingColumn.getRange().getValues();
  const previousValue = readingColumnValues[readingColumnValues.length - 1][0] as number;

  // Add a row with the date, new value, and a formula calculating the difference.
  const currentDate = new Date(Date.now()).toLocaleDateString();
  const newRow = [currentDate, newData, "=[@Reading]-OFFSET([@Reading],-1,0)"];
  table.addRow(-1, newRow,);

  // Return the difference between the newData and the previous entry.
  const difference = Number.parseFloat(newData) - previousValue;
  console.log(difference);
  return difference;
}
```

## <a name="sample-flow-report-day-to-day-changes"></a>サンプル フロー: 毎日の変更を報告する

サンプルの [Power Automate](https://powerautomate.microsoft.com/) フローを構築するには、次の手順に従います。

1. 新しい **スケジュールされたクラウド フロー** を作成します。
1. **フローを 1 日** ごとに繰り返すスケジュールを設定します。

    :::image type="content" source="../../images/day-to-day-changes-flow-1.png" alt-text="それを示すフロー作成手順は毎日繰り返されます。":::
1. **[作成]** を選択します。
1. 実際のフローでは、データを取得するステップを追加します。 データは、別のブック、Teams アダプティブ カード、またはその他のソースから取得できます。 サンプルをテストするには、テスト番号を作成します。 **変数の初期化** アクションを使用して新しいステップを追加します。 次の値を指定します。
    1. **名前**: 入力
    1. **型**: 整数
    1. **値**: 190000

    :::image type="content" source="../../images/day-to-day-changes-flow-2.png" alt-text="指定された値を持つ変数の初期化アクション。":::
1. **スクリプトの実行** アクションを使用して **、Excel Online (Business)** コネクタを使用して新しいステップを追加します。 アクションには次の値を使用します。
    1. **場所**: OneDrive for Business
    1. **ドキュメント ライブラリ**: OneDrive
    1. **ファイル**: daily-readings.xlsx *(ファイル ブラウザーから選択)*
    1. **スクリプト**: スクリプト名
    1. **newData**: 入力 *(動的コンテンツ)*

    :::image type="content" source="../../images/day-to-day-changes-flow-3.png" alt-text="指定された値を使用したスクリプトの実行アクション。":::
1. このスクリプトは、"result" という名前の動的コンテンツとして、毎日の読み取りの違いを返します。 サンプルの場合は、自分に情報を電子メールで送信できます。 **電子メールの送信 (V2) アクション (** または任意の電子メール クライアント) で **Outlook** コネクタを使用する新しい手順を作成します。 アクションを完了するには、次の値を使用します。
    1. **宛先**: メール アドレス
    1. **件名**: 毎日の読み取り変更
    1. **本文**: "昨日との違い" の結果 *(Excel の動的コンテンツ)*

    :::image type="content" source="../../images/day-to-day-changes-flow-4.png" alt-text="Power Automate で完成した Outlook コネクタ。":::
1. フローを保存して試してください。フロー エディター ページの **[テスト** ] ボタンを使用します。 メッセージが表示されたら、必ずアクセスを許可してください。
