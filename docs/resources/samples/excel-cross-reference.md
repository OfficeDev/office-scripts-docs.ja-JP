---
title: Power Automate を使用した Excel ファイルの相互参照
description: Office スクリプトと Power Automate を使用して Excel ファイルを相互参照および書式設定する方法について説明します。
ms.date: 06/06/2022
ms.localizationpriority: medium
ms.openlocfilehash: b32249dc7cb1e8c1b841a4db6caaff3b4d2998ec
ms.sourcegitcommit: a6504f8b0d6b717457c6e0b5306c35ad3900914e
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 09/02/2022
ms.locfileid: "67572676"
---
# <a name="cross-reference-excel-files-with-power-automate"></a>Power Automate を使用した Excel ファイルの相互参照

このソリューションでは、2 つの Excel ファイル間でデータを比較して不一致を見つける方法を示します。 Office スクリプトを使用してデータを分析し、Power Automate を使用してブック間で通信します。

このサンプルでは、 [JSON](https://www.w3schools.com/whatis/whatis_json.asp) オブジェクトを使用してブック間でデータを渡します。 JSON の操作の詳細については、「JSON を [使用して Office スクリプトとの間でデータを渡す」を参照](../../develop/use-json.md)してください。

## <a name="example-scenario"></a>シナリオ例

あなたは、今後の会議の講演者をスケジュールしているイベント コーディネーターです。 イベント データは 1 つのスプレッドシートに保持し、話者の登録は別のスプレッドシートに保持します。 2 つのブックが確実に同期されるようにするには、Office スクリプトでフローを使用して、潜在的な問題を強調表示します。

## <a name="sample-excel-files"></a>Excel ファイルのサンプル

サンプルのすぐに使用できるブックを取得するには、次のファイルをダウンロードします。

1. [event-data.xlsx](event-data.xlsx)
1. [speaker-registrations.xlsx](speaker-registrations.xlsx)

サンプルを自分で試すには、次のスクリプトを追加します。

## <a name="sample-code-get-event-data"></a>サンプル コード: イベント データを取得する

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

## <a name="sample-code-validate-speaker-registrations"></a>サンプル コード: 話者の登録を検証する

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

## <a name="power-automate-flow-check-for-inconsistencies-across-the-workbooks"></a>Power Automate フロー: ブック間の不整合を確認する

このフローは、最初のブックからイベント情報を抽出し、そのデータを使用して 2 番目のブックを検証します。

1. [Power Automate](https://flow.microsoft.com) にサインインし、新しい **インスタント クラウド フロー** を作成します。
1. [ **手動でフローをトリガーする** ] を選択し、[ **作成**] を選択します。
1. **スクリプトの実行** アクションで **Excel Online (Business)** コネクタを使用する **新しいステップ** を追加します。 アクションには次の値を使用します。
    * **場所**: OneDrive for Business
    * **ドキュメント ライブラリ**: OneDrive
    * **ファイル**: event-data.xlsx ([ファイル選択子で選択)](../../testing/power-automate-troubleshooting.md#select-workbooks-with-the-file-browser-control)
    * **スクリプト**: イベント データを取得する

    :::image type="content" source="../../images/cross-reference-flow-1.png" alt-text="Power Automate の最初のスクリプト用に完成した Excel Online (Business) コネクタ。":::

1. **スクリプトの実行** アクションで **Excel Online (Business)** コネクタを使用する 2 番目の **新しい手順** を追加します。 これにより、イベント データの **検証** スクリプトの入力として **、Get イベント データ** スクリプトから返された値が使用されます。 アクションには次の値を使用します。
    * **場所**: OneDrive for Business
    * **ドキュメント ライブラリ**: OneDrive
    * **ファイル**: speaker-registration.xlsx ([ファイル選択子で選択)](../../testing/power-automate-troubleshooting.md#select-workbooks-with-the-file-browser-control)
    * **スクリプト**: 話者の登録を検証する
    * **キー**: 結果 (_**実行スクリプト** からの動的コンテンツ_)

    :::image type="content" source="../../images/cross-reference-flow-2.png" alt-text="Power Automate の 2 番目のスクリプト用に完成した Excel Online (Business) コネクタ。":::
1. このサンプルでは、Outlook を電子メール クライアントとして使用します。 Power Automate でサポートされている任意の電子メール コネクタを使用できます。 **Office 365 Outlook** コネクタと **送信と電子メール (V2)** アクションを使用する **新しい手順** を追加します。 これは、電子メール本文のコンテンツとして **、スピーカー登録の検証** スクリプトから返された値を使用します。 アクションには次の値を使用します。
    * **To**: テスト用メール アカウント (または個人用メール)
    * **件名**: イベント検証の結果
    * **本文**: 結果 (_**実行スクリプト 2** の動的コンテンツ_)

    :::image type="content" source="../../images/cross-reference-flow-3.png" alt-text="Power Automate の完成したOffice 365 Outlook コネクタ。":::
1. フローを保存します。 フロー エディター ページの **[テスト** ] ボタンを使用するか、[ **マイ フロー** ] タブでフローを実行します。メッセージが表示されたら、必ずアクセスを許可してください。
1. "不一致が見つかりました。 データにはレビューが必要です。 これは、speaker-registrations.xlsxの行と **event-data.xlsx** **の行** の間に違いがあることを示しています。 **speaker-registrations.xlsx** を開いて、スピーカー登録の一覧に問題があるいくつかの強調表示されたセルを表示します。
