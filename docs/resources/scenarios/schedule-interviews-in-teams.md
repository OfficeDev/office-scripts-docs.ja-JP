---
title: Teams で面接をスケジュールする
description: Office スクリプトを使用して、Excel データからTeams会議を送信する方法について説明します。
ms.date: 06/29/2021
ms.localizationpriority: medium
ms.openlocfilehash: 1c07eed0ce8392cf6d08f7836970753194f54b05
ms.sourcegitcommit: dd01979d34b3499360d2f79a56f8a8f24f480eed
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 06/15/2022
ms.locfileid: "66088058"
---
# <a name="office-scripts-sample-scenario-schedule-interviews-in-teams"></a>Office スクリプトのサンプル シナリオ: Teamsで面接をスケジュールする

このシナリオでは、人事採用担当者が、Teamsの候補者との面接会議をスケジュールしています。 Excel ファイル内の候補者の面接スケジュールを管理します。 Teams会議出席依頼を候補者と面接官の両方に送信する必要があります。 その後、Excel ファイルを更新し、Teams会議が送信されたことを確認する必要があります。

このソリューションには、1 つのPower Automate フローで組み合わされた 3 つの手順があります。

1. スクリプトはテーブルからデータを抽出し、オブジェクトの配列を [JSON](https://www.w3schools.com/whatis/whatis_json.asp) データとして返します。
1. その後、Teams会議アクションを作成して招待を送信 **するTeams** にデータが送信されます。
1. 同じ JSON データが別のスクリプトに送信され、招待の状態が更新されます。

JSON の操作の詳細については、「[JSON を使用して、Office スクリプトとの間でデータを渡す](../../develop/use-json.md)」を参照してください。

## <a name="scripting-skills-covered"></a>スクリプティング スキルの説明

* フローのPower Automate
* Teams統合
* テーブルの解析

## <a name="sample-excel-file"></a>サンプル Excel ファイル

このソリューションで使用 <a href="hr-schedule.xlsx">hr-schedule.xlsx</a> ファイルをダウンロードし、自分で試してみてください。 招待を受け取ることができるように、少なくとも 1 つのメール アドレスを変更してください。

## <a name="sample-code-extract-table-data-to-schedule-invites"></a>サンプル コード: テーブル データを抽出して招待をスケジュールする

このスクリプトをスクリプト コレクションに追加します。 フローの **インタビューのスケジュール** に名前を付けます。

```TypeScript
function main(workbook: ExcelScript.Workbook): InterviewInvite[] {
  const MEETING_DURATION = workbook.getWorksheet("Constants").getRange("B1").getValue() as number;
  const MESSAGE_TEMPLATE = workbook.getWorksheet("Constants").getRange("B2").getValue() as string;

  // Get the interview candidate information.
  const sheet = workbook.getWorksheet("Interviews");
  const table = sheet.getTables()[0];
  const dataRows = table.getRangeBetweenHeaderAndTotal().getValues();

  // Convert the table rows into InterviewInvite objects for the flow.
  let invites: InterviewInvite[] = [];
  dataRows.forEach((row) => {
    const inviteSent = row[1] as boolean;
    if (!inviteSent) {
      const startTime = new Date(Math.round(((row[6] as number) - 25569) * 86400 * 1000));
      const finishTime = new Date(startTime.getTime() + MEETING_DURATION * 60 * 1000);
      const candidateName = row[2] as string;
      const interviewerName = row[4] as string;

      invites.push({
        ID: row[0] as string,
        Candidate: candidateName,
        CandidateEmail: row[3] as string,
        Interviewer: row[4] as string,
        InterviewerEmail: row[5] as string,
        StartTime: startTime.toISOString(),
        FinishTime: finishTime.toISOString(),
        Message: generateInviteMessage(MESSAGE_TEMPLATE, candidateName, interviewerName)
      });
    }    
  });

  console.log(JSON.stringify(invites));
  return invites;
}

function generateInviteMessage(
  messageTemplate: string,
   candidate: string,
   interviewer: string) : string {
  return messageTemplate.replace("_Candidate_", candidate).replace("_Interviewer_", interviewer);
}

// The interview invite information.
interface InterviewInvite {
  ID: string
  Candidate: string
  CandidateEmail: string
  Interviewer: string
  InterviewerEmail: string
  StartTime: string
  FinishTime: string
  Message: string
}
```

## <a name="sample-code-mark-rows-as-invited"></a>サンプル コード: 行を招待済みとしてマークする

このスクリプトをスクリプト コレクションに追加します。 フローの **レコード送信済み招待に** 名前を付けます。

```TypeScript
function main(workbook: ExcelScript.Workbook, invites: InterviewInvite[]) {
  const table = workbook.getWorksheet("Interviews").getTables()[0];

  // Get the ID and Invite Sent columns from the table.
  const idColumn = table.getColumnByName("ID");
  const idRange = idColumn.getRangeBetweenHeaderAndTotal().getValues();
  const inviteSentColumn = table.getColumnByName("Invite Sent?");

  const dataRowCount = idRange.length;

  // Find matching IDs to mark the correct row.
  for (let row = 0; row < dataRowCount; row++){
    let inviteSent = invites.find((invite) => {
      return invite.ID == idRange[row][0] as string;
    });

    if (inviteSent) {
      inviteSentColumn.getRangeBetweenHeaderAndTotal().getCell(row, 0).setValue(true);
      console.log(`Invite for ${inviteSent.Candidate} has been sent.`);
    }
  } 
}

// The interview invite information.
interface InterviewInvite {
  ID: string
  Candidate: string
  CandidateEmail: string
  Interviewer: string
  InterviewerEmail: string
  StartTime: string
  FinishTime: string
  Message: string
}
```

## <a name="sample-flow-run-the-interview-scheduling-scripts-and-send-the-teams-meetings"></a>サンプル フロー: 面接スケジュール スクリプトを実行し、Teams会議を送信する

1. 新しい **インスタント クラウド フロー** を作成します。
1. [ **手動でフローをトリガーする** ] を選択し、[ **作成**] を選択します。
1. **Excel Online (Business)** コネクタと **スクリプトの実行** アクションを使用する **新しい手順** を追加します。 次の値を使用してコネクタを完了します。
    1. **場所**: OneDrive for Business
    1. **ドキュメント ライブラリ**: OneDrive
    1. **ファイル**: hr-interviews.xlsx *(ファイル ブラウザーから選択)*
    1. **スクリプト**: :::image type="content" source="../../images/schedule-interviews-1.png" alt-text="完了したExcel Online (Business) コネクタのインタビューのスケジュールスクリーンショット。Power Automateのブックからインタビュー データを取得します。":::
1. **[Teams会議の作成**] アクションを使用する **新しいステップ** を追加します。 Excel コネクタから動的コンテンツを選択すると、フロー **に対して各ブロックに適用** が生成されます。 次の値を使用してコネクタを完了します。
    1. **予定表 ID**: 予定表
    1. **件名**: Contoso インタビュー
    1. **メッセージ**: **メッセージ** (Excel値)
    1. **タイム ゾーン**: 太平洋標準時
    1. **開始時刻**: **StartTime** (Excel値)
    1. **終了時刻**: **FinishTime** (Excel値)
    1. **必要な出席者**: **CandidateEmail** ;**InterviewerEmail** (Excel値) :::image type="content" source="../../images/schedule-interviews-2.png" alt-text="Power Automateで会議をスケジュールするための完成したTeams コネクタのスクリーンショット。":::
1. **同じ [各ブロックに適用]** で、[**スクリプトの実行**] アクション **を使用して別の Excel Online (Business)** コネクタを追加します。 次の値を使用します。
    1. **場所**: OneDrive for Business
    1. **ドキュメント ライブラリ**: OneDrive
    1. **ファイル**: hr-interviews.xlsx *(ファイル ブラウザーから選択)*
    1. **スクリプト**: 送信された招待を記録する
    1. **invites**: **結果** (Excel値) Power Automateで :::image type="content" source="../../images/schedule-interviews-3.png" alt-text="送信された招待を記録する完了したExcel Online (Business) コネクタのスクリーンショット。":::
1. フローを保存して試してください。フロー エディター ページの **[テスト** ] ボタンを使用するか、[ **マイ フロー** ] タブでフローを実行します。メッセージが表示されたら、必ずアクセスを許可してください。

## <a name="training-video-send-a-teams-meeting-from-excel-data"></a>トレーニング ビデオ: Excel データからTeams会議を送信する

[YouTube でこのサンプルのバージョンを見て、スディ Ramamurthy が歩くのを見てください](https://youtu.be/HyBdx52NOE8)。 彼のバージョンでは、列の変更と古い会議時間を処理する、より堅牢なスクリプトを使用しています。
