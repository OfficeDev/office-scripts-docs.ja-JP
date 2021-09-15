---
title: ファイルとファイルExcel相互参照Power Automate
description: スクリプトとスクリプトを使用Office、Power Automateファイルを相互参照して書式設定するExcelします。
ms.date: 06/29/2021
ms.localizationpriority: medium
ms.openlocfilehash: adeb84140cb9884309c9f37854a29fc4d59b17ed
ms.sourcegitcommit: d3ed4bdeeba805d97c930394e172e8306a0cf484
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 09/15/2021
ms.locfileid: "59332979"
---
# <a name="cross-reference-excel-files-with-power-automate"></a>ファイルとファイルExcel相互参照Power Automate

このソリューションでは、2 つのファイル間でデータを比較Excel不一致を見つける方法を示します。 このスクリプトはOfficeを使用してデータを分析し、Power Automate間の通信を行います。

## <a name="example-scenario"></a>シナリオ例

今後の会議にスピーカーをスケジュールしているイベント コーディネーターです。 イベント データは 1 つのスプレッドシートに、スピーカーの登録は別のスプレッドシートに保持します。 2 つのブックの同期を確実に行う場合は、Officeスクリプトを使用して、潜在的な問題を強調表示します。

## <a name="sample-excel-files"></a>サンプル Excel ファイル

次のファイルをダウンロードして、サンプルのすぐに使用できるブックを取得します。

1. <a href="event-data.xlsx">event-data.xlsx</a>
1. <a href="speaker-registrations.xlsx">speaker-registrations.xlsx</a>

次のスクリプトを追加して、サンプルを自分で試してみてください。

## <a name="sample-code-get-event-data"></a>サンプル コード: イベント データの取得

```TypeScript
function main(workbook: ExcelScript.Workbook): string {
  // Get the first table in the "Keys" worksheet.
  let table = workbook.getWorksheet('Keys').getTables()[0];

  // Get the rows in the event table.
  let range = table.getRangeBetweenHeaderAndTotal();
  let rows = range.getValues();

  // Save each row as an EventData object. This lets them be passed through Power Automate.
  let records: EventData[] = [];
  for (let row of rows) {
    let [eventId, date, location, capacity] = row;
    records.push({
      eventId: eventId as string,
      date: date as number,
      location: location as string,
      capacity: capacity as number
    })
  }

  // Log the event data to the console and return it for a flow.
  let stringResult = JSON.stringify(records);
  console.log(stringResult);
  return stringResult;
}

// An interface representing a row of event data.
interface EventData {
  eventId: string
  date: number
  location: string
  capacity: number
}
```

## <a name="sample-code-validate-speaker-registrations"></a>サンプル コード: スピーカー登録の検証

```TypeScript
function main(workbook: ExcelScript.Workbook, keys: string): string {
  // Get the first table in the "Transactions" worksheet.
  let table = workbook.getWorksheet('Transactions').getTables()[0];

  // Clear the existing formatting in the table.
  let range = table.getRangeBetweenHeaderAndTotal();
  range.clear(ExcelScript.ClearApplyTo.formats);

  // Compare the data in the table to the keys passed into the script.
  let keysObject = JSON.parse(keys) as EventData[];
  let speakerSlotsRemaining = keysObject.map(value => value.capacity);
  let overallMatch = true;

  // Iterate over every row looking for differences from the other worksheet.
  let rows = range.getValues();
  for (let i = 0; i < rows.length; i++) {
    let row = rows[i];
    let [eventId, date, location, capacity] = row;
    let match = false;

    // Look at each key provided for a matching Event ID.
    for (let keyIndex = 0; keyIndex < keysObject.length; keyIndex++) {
      let event = keysObject[keyIndex];
      if (event.eventId === eventId) {
        match = true;
        speakerSlotsRemaining[keyIndex]--;
        // If there's a match on the event ID, look for things that don't match and highlight them.
        if (event.date !== date) {
          overallMatch = false;
          range.getCell(i, 1).getFormat()
            .getFill()
            .setColor("FFFF00");
        }
        if (event.location !== location) {
          overallMatch = false;
          range.getCell(i, 2).getFormat()
            .getFill()
            .setColor("FFFF00");
        }

        break;
      }
    }

    // If no matching Event ID is found, highlight the Event ID's cell.
    if (!match) {
      overallMatch = false;
      range.getCell(i, 0).getFormat()
        .getFill()
        .setColor("FFFF00");
    }
  }

  

  // Choose a message to send to the user.
  let returnString = "All the data is in the right order.";
  if (overallMatch === false) {
    returnString = "Mismatch found. Data requires your review.";
  } else if (speakerSlotsRemaining.find(remaining => remaining < 0)){
    returnString = "Event potentially overbooked. Please review."
  }

  console.log("Returning: " + returnString);
  return returnString;
}

// An interface representing a row of event data.
interface EventData {
  eventId: string
  date: number
  location: string
  capacity: number
}
```

## <a name="power-automate-flow-check-for-inconsistencies-across-the-workbooks"></a>Power Automateフロー: ブック全体の不整合を確認する

このフローは、最初のブックからイベント情報を抽出し、そのデータを使用して 2 番目のブックを検証します。

1. 新しいインスタント [Power Automate](https://flow.microsoft.com)にサインインし、**新しいインスタント クラウド フローを作成します**。
1. [フロー **を手動でトリガーする] を選択し** 、[作成] を **選択します**。
1. [スクリプト **の実行]** アクションを使用して、Excel **(Business)** コネクタを使用する新しい **手順を追加** します。 アクションには、次の値を使用します。
    * **場所**: OneDrive for Business
    * **ドキュメント ライブラリ**: OneDrive
    * **ファイル**: event-data.xlsx ([ファイル選択で選択)](../../testing/power-automate-troubleshooting.md#select-workbooks-with-the-file-browser-control)
    * **スクリプト**: イベント データの取得

    :::image type="content" source="../../images/cross-reference-flow-1.png" alt-text="最初のスクリプトExcelオンライン (Business) コネクタの完成Power Automate。":::

1. [スクリプトの実行 **] アクション** を使用して、Excel **(Business)** コネクタを使用する 2 番目の新しい **手順を追加** します。 アクションには、次の値を使用します。
    * **場所**: OneDrive for Business
    * **ドキュメント ライブラリ**: OneDrive
    * **ファイル**: speaker-registration.xlsx ([ファイル選択で選択)](../../testing/power-automate-troubleshooting.md#select-workbooks-with-the-file-browser-control)
    * **スクリプト**: スピーカー登録の検証

    :::image type="content" source="../../images/cross-reference-flow-2.png" alt-text="2 番目Excelのオンライン (Business) コネクタの完成Power Automate。":::
1. このサンプルでは、Outlookクライアントとして使用します。 サポートされている任意の電子メール コネクタPower Automate使用できます。 新しい **手順を追加** して、Office 365 Outlook **および電子** メール **(V2) アクションを使用** します。 アクションには、次の値を使用します。
    * **To**: テスト用メール アカウント (または個人用メール)
    * **件名**: イベントの検証結果
    * **本文**: result (_Run スクリプト 2 からの **動的コンテンツ**_)

    :::image type="content" source="../../images/cross-reference-flow-3.png" alt-text="Office 365 OutlookでPower Automate。":::
1. フローを保存します。 [フロー エディター **] ページ** の [テスト] ボタンを使用するか、[マイ フロー] タブでフロー **を実行** します。メッセージが表示されたら、必ずアクセスを許可してください。
1. "不一致が見つかりました" というメールを受信する必要があります。 データにはレビューが必要です。 これは、グループ内の行と **speaker-registrations.xlsx行の** 間に違 **いevent-data.xlsx。** [speaker-registrations.xlsxを **開** き、スピーカー登録リストに潜在的な問題があるいくつかの強調表示されたセルを表示します。
