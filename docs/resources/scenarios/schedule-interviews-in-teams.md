---
title: 面接のスケジュールを設定Teams
description: '[スクリプト] を使用してOfficeデータから会議Teams送信するExcelします。'
ms.date: 05/25/2021
localization_priority: Normal
ms.openlocfilehash: f93d9ceca6603ddb9e7123a393787fcf54597cca
ms.sourcegitcommit: 339ecbb9914d54f919e3475018888fb5d00abe89
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 05/28/2021
ms.locfileid: "52697785"
---
# <a name="office-scripts-sample-scenario-schedule-interviews-in-teams"></a>Officeスクリプトのサンプル シナリオ: スケジュールの面接のスケジュールを設定Teams

このシナリオでは、人事担当の採用担当者が、面接会議をスケジュールし、Teams。 候補者の面接スケジュールは、1 つのファイルExcelします。 候補者と面接官の両方にTeams会議の招待を送信する必要があります。 その後、会議が送信されたExcel確認して、Teamsファイルを更新する必要があります。

ソリューションには、1 つのフローで組み合わされる 3 つのPower Automateがあります。

1. スクリプトはテーブルからデータを抽出し、オブジェクトの配列を JSON データとして返します。
1. 次に、データが [会議の作成] Teamsに送信Teams **に** 送信されます。
1. 招待の状態を更新するために、同じ JSON データが別のスクリプトに送信されます。

## <a name="scripting-skills-covered"></a>スクリプティングのスキルをカバー

* Power Automateフロー
* Teams統合
* テーブルの解析

## <a name="sample-excel-file"></a>サンプル Excel ファイル

このソリューションで <a href="hr-schedule.xlsx">hr-schedule.xlsx</a> ファイルをダウンロードして、自分で試してみてください。 招待を受け取る電子メール アドレスを少なくとも 1 つ変更してください。

## <a name="sample-code-extract-table-data-to-schedule-invites"></a>サンプル コード: テーブル データを抽出して招待をスケジュールする

このスクリプトに、 **フローのインタビューのスケジュール** を指定します。

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

このスクリプトに **、フローの [送信された招待を記録する]** という名前を指定します。

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

## <a name="sample-flow-run-the-interview-scheduling-scripts-and-send-the-teams-meetings"></a>サンプル フロー: 面接スケジュール スクリプトを実行し、会議にTeamsする

1. 新しいインスタント クラウド **フローを作成します**。
1. [フロー **を手動でトリガーする] を選択し** 、[作成] を **押します**。
1. オンライン **(Business)** コネクタと [スクリプト **Excel実行]** アクションを使用する新しい **手順を追加** します。 コネクタに次の値を入力します。
    1. **場所**: OneDrive for Business
    1. **ドキュメント ライブラリ**: OneDrive
    1. **ファイル**: hr-interviews.xlsx *(ファイル ブラウザーから選択)*
    1. **スクリプト**: オンライン :::image type="content" source="../../images/schedule-interviews-1.png" alt-text="(Business)":::コネクタで完了したExcelのスケジュール のスクリーンショットを使用して、ブックからインタビュー データを取得Power Automate
1. [会議の **作成] アクション** を使用する **新しいTeams追加** します。 コネクタから動的コンテンツを選択Excel、フローに対して各ブロックに **適用** が生成されます。 コネクタに次の値を入力します。
    1. **予定表 ID**: Calendar
    1. **件名**: Contoso インタビュー
    1. **メッセージ**:**メッセージ**(Excel値)
    1. **タイム ゾーン**: 太平洋標準時
    1. **開始時刻**: **StartTime** (Excel値)
    1. **終了時刻**: **FinishTime** (Excel値)
    1. **必須の出席者**: **CandidateEmail** ;**InterviewerEmail** (Excel値) 完了したコネクタのスクリーンショットTeamsで会議をスケジュール :::image type="content" source="../../images/schedule-interviews-2.png" alt-text="Power Automate":::
1. 同じ [各 **ブロックに適用] で**、[スクリプトの実行] アクションExcel **オンライン (Business)** コネクタを **追加** します。 次の値を使用します。
    1. **場所**: OneDrive for Business
    1. **ドキュメント ライブラリ**: OneDrive
    1. **ファイル**: hr-interviews.xlsx *(ファイル ブラウザーから選択)*
    1. **スクリプト**: 送信された招待を記録する
    1. **invites**: **result** (Excel 値) 完了した :::image type="content" source="../../images/schedule-interviews-3.png" alt-text="Excel Online (Business)":::コネクタのスクリーンショットで、招待が送信されたと記録Power Automate
1. フローを保存し、試してみてください。

## <a name="training-video-send-a-teams-meeting-from-excel-data"></a>トレーニング ビデオ: データからTeams会議をExcelする

[Sudhi Ramamurthy が YouTube でこのサンプル](https://youtu.be/HyBdx52NOE8)のバージョンを見るをご覧ください。 彼のバージョンでは、列の変更や廃止された会議の時間を処理する、より堅牢なスクリプトを使用しています。
