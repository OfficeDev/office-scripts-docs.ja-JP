---
title: Teams で面接をスケジュールする
description: '[スクリプト] を使用してOfficeデータから会議Teams送信するExcelします。'
ms.date: 06/29/2021
localization_priority: Normal
ms.openlocfilehash: cb24da12637add805d86da4d07ce878509c6a5f6
ms.sourcegitcommit: 211c157ca746e266eeb079f5fa1925a1e35ab702
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 07/07/2021
ms.locfileid: "53313730"
---
# <a name="office-scripts-sample-scenario-schedule-interviews-in-teams"></a><span data-ttu-id="a4c26-103">Officeスクリプトのサンプル シナリオ: スケジュールの面接のスケジュールを設定Teams</span><span class="sxs-lookup"><span data-stu-id="a4c26-103">Office Scripts sample scenario: Schedule interviews in Teams</span></span>

<span data-ttu-id="a4c26-104">このシナリオでは、人事担当の採用担当者が、面接会議をスケジュールし、Teams。</span><span class="sxs-lookup"><span data-stu-id="a4c26-104">In this scenario, you're an HR recruiter scheduling interview meetings with candidates in Teams.</span></span> <span data-ttu-id="a4c26-105">候補者の面接スケジュールは、1 つのファイルExcelします。</span><span class="sxs-lookup"><span data-stu-id="a4c26-105">You manage the interview schedule of candidates in an Excel file.</span></span> <span data-ttu-id="a4c26-106">候補者と面接官の両方にTeams会議の招待を送信する必要があります。</span><span class="sxs-lookup"><span data-stu-id="a4c26-106">You'll need to send the Teams meeting invite to both the candidate and interviewers.</span></span> <span data-ttu-id="a4c26-107">その後、会議が送信されたExcel確認して、Teamsファイルを更新する必要があります。</span><span class="sxs-lookup"><span data-stu-id="a4c26-107">You then need to update the Excel file with the confirmation that Teams meetings have been sent.</span></span>

<span data-ttu-id="a4c26-108">ソリューションには、1 つのフローで組み合わされる 3 つのPower Automateがあります。</span><span class="sxs-lookup"><span data-stu-id="a4c26-108">The solution has three steps that are combined in a single Power Automate flow.</span></span>

1. <span data-ttu-id="a4c26-109">スクリプトはテーブルからデータを抽出し、オブジェクトの配列を JSON データとして返します。</span><span class="sxs-lookup"><span data-stu-id="a4c26-109">A script extracts data from a table and returns an array of objects as JSON data.</span></span>
1. <span data-ttu-id="a4c26-110">次に、データが [会議の作成] Teamsに送信Teams **に** 送信されます。</span><span class="sxs-lookup"><span data-stu-id="a4c26-110">The data is then sent to the Teams **Create a Teams meeting** action to send invites.</span></span>
1. <span data-ttu-id="a4c26-111">招待の状態を更新するために、同じ JSON データが別のスクリプトに送信されます。</span><span class="sxs-lookup"><span data-stu-id="a4c26-111">The same JSON data is sent to another script to update the status of the invitation.</span></span>

## <a name="scripting-skills-covered"></a><span data-ttu-id="a4c26-112">スクリプティングのスキルをカバー</span><span class="sxs-lookup"><span data-stu-id="a4c26-112">Scripting skills covered</span></span>

* <span data-ttu-id="a4c26-113">Power Automateフロー</span><span class="sxs-lookup"><span data-stu-id="a4c26-113">Power Automate flows</span></span>
* <span data-ttu-id="a4c26-114">Teams統合</span><span class="sxs-lookup"><span data-stu-id="a4c26-114">Teams integration</span></span>
* <span data-ttu-id="a4c26-115">テーブルの解析</span><span class="sxs-lookup"><span data-stu-id="a4c26-115">Table parsing</span></span>

## <a name="sample-excel-file"></a><span data-ttu-id="a4c26-116">サンプル Excel ファイル</span><span class="sxs-lookup"><span data-stu-id="a4c26-116">Sample Excel file</span></span>

<span data-ttu-id="a4c26-117">このソリューションで <a href="hr-schedule.xlsx">hr-schedule.xlsx</a> ファイルをダウンロードして、自分で試してみてください。</span><span class="sxs-lookup"><span data-stu-id="a4c26-117">Download the file <a href="hr-schedule.xlsx">hr-schedule.xlsx</a> used in this solution and try it out yourself!</span></span> <span data-ttu-id="a4c26-118">招待を受け取る電子メール アドレスを少なくとも 1 つ変更してください。</span><span class="sxs-lookup"><span data-stu-id="a4c26-118">Be sure to change at least one of the email addresses so that you receive an invite.</span></span>

## <a name="sample-code-extract-table-data-to-schedule-invites"></a><span data-ttu-id="a4c26-119">サンプル コード: テーブル データを抽出して招待をスケジュールする</span><span class="sxs-lookup"><span data-stu-id="a4c26-119">Sample code: Extract table data to schedule invites</span></span>

<span data-ttu-id="a4c26-120">スクリプト コレクションにこのスクリプトを追加します。</span><span class="sxs-lookup"><span data-stu-id="a4c26-120">Add this script to your script collection.</span></span> <span data-ttu-id="a4c26-121">フローの **[面接のスケジュール]** という名前を付けします。</span><span class="sxs-lookup"><span data-stu-id="a4c26-121">Name it **Schedule Interviews** for the flow.</span></span>

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

## <a name="sample-code-mark-rows-as-invited"></a><span data-ttu-id="a4c26-122">サンプル コード: 行を招待済みとしてマークする</span><span class="sxs-lookup"><span data-stu-id="a4c26-122">Sample code: Mark rows as invited</span></span>

<span data-ttu-id="a4c26-123">スクリプト コレクションにこのスクリプトを追加します。</span><span class="sxs-lookup"><span data-stu-id="a4c26-123">Add this script to your script collection.</span></span> <span data-ttu-id="a4c26-124">フローの **[送信された招待を記録する** ] という名前を付けします。</span><span class="sxs-lookup"><span data-stu-id="a4c26-124">Name it **Record Sent Invites** for the flow.</span></span>

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

## <a name="sample-flow-run-the-interview-scheduling-scripts-and-send-the-teams-meetings"></a><span data-ttu-id="a4c26-125">サンプル フロー: 面接スケジュール スクリプトを実行し、会議にTeamsする</span><span class="sxs-lookup"><span data-stu-id="a4c26-125">Sample flow: Run the interview scheduling scripts and send the Teams meetings</span></span>

1. <span data-ttu-id="a4c26-126">新しいインスタント クラウド **フローを作成します**。</span><span class="sxs-lookup"><span data-stu-id="a4c26-126">Create a new **Instant cloud flow**.</span></span>
1. <span data-ttu-id="a4c26-127">[フロー **を手動でトリガーする] を選択し** 、[作成] を **選択します**。</span><span class="sxs-lookup"><span data-stu-id="a4c26-127">Choose **Manually trigger a flow** and select **Create**.</span></span>
1. <span data-ttu-id="a4c26-128">オンライン **(Business)** コネクタと [スクリプト **Excel実行]** アクションを使用する新しい **手順を追加** します。</span><span class="sxs-lookup"><span data-stu-id="a4c26-128">Add a **New step** that uses the **Excel Online (Business)** connector and the **Run script** action.</span></span> <span data-ttu-id="a4c26-129">コネクタに次の値を入力します。</span><span class="sxs-lookup"><span data-stu-id="a4c26-129">Complete the connector with the following values.</span></span>
    1. <span data-ttu-id="a4c26-130">**場所**: OneDrive for Business</span><span class="sxs-lookup"><span data-stu-id="a4c26-130">**Location**: OneDrive for Business</span></span>
    1. <span data-ttu-id="a4c26-131">**ドキュメント ライブラリ**: OneDrive</span><span class="sxs-lookup"><span data-stu-id="a4c26-131">**Document Library**: OneDrive</span></span>
    1. <span data-ttu-id="a4c26-132">**ファイル**: hr-interviews.xlsx *(ファイル ブラウザーから選択)*</span><span class="sxs-lookup"><span data-stu-id="a4c26-132">**File**: hr-interviews.xlsx *(Chosen through the file browser)*</span></span>
    1. **スクリプト**: オンライン :::image type="content" source="../../images/schedule-interviews-1.png" alt-text="(Business) コネクタExcel完了したインタビュー":::のスクリーンショットをスケジュールして、ブックからインタビュー データを取得Power Automate。
1. <span data-ttu-id="a4c26-134">[会議の **作成] アクション** を使用する **新しいTeams追加** します。</span><span class="sxs-lookup"><span data-stu-id="a4c26-134">Add a **New step** that uses the **Create a Teams meeting** action.</span></span> <span data-ttu-id="a4c26-135">コネクタから動的コンテンツを選択Excel、フローに対して各ブロックに **適用** が生成されます。</span><span class="sxs-lookup"><span data-stu-id="a4c26-135">As you select dynamic content from the Excel connector, an **Apply to each** block will be generated for your flow.</span></span> <span data-ttu-id="a4c26-136">コネクタに次の値を入力します。</span><span class="sxs-lookup"><span data-stu-id="a4c26-136">Complete the connector with the following values.</span></span>
    1. <span data-ttu-id="a4c26-137">**予定表 ID**: Calendar</span><span class="sxs-lookup"><span data-stu-id="a4c26-137">**Calendar id**: Calendar</span></span>
    1. <span data-ttu-id="a4c26-138">**件名**: Contoso インタビュー</span><span class="sxs-lookup"><span data-stu-id="a4c26-138">**Subject**: Contoso Interview</span></span>
    1. <span data-ttu-id="a4c26-139">**メッセージ**:**メッセージ**(Excel値)</span><span class="sxs-lookup"><span data-stu-id="a4c26-139">**Message**: **Message** (the Excel value)</span></span>
    1. <span data-ttu-id="a4c26-140">**タイム ゾーン**: 太平洋標準時</span><span class="sxs-lookup"><span data-stu-id="a4c26-140">**Time zone**: Pacific Standard Time</span></span>
    1. <span data-ttu-id="a4c26-141">**開始時刻**: **StartTime** (Excel値)</span><span class="sxs-lookup"><span data-stu-id="a4c26-141">**Start time**: **StartTime** (the Excel value)</span></span>
    1. <span data-ttu-id="a4c26-142">**終了時刻**: **FinishTime** (Excel値)</span><span class="sxs-lookup"><span data-stu-id="a4c26-142">**End time**: **FinishTime** (the Excel value)</span></span>
    1. **必須の出席者**: **CandidateEmail** ;**InterviewerEmail** (Excel値) 完了したコネクタのスクリーンショットTeamsで会議をスケジュール :::image type="content" source="../../images/schedule-interviews-2.png" alt-text="Power Automate。":::
1. <span data-ttu-id="a4c26-144">同じ [各 **ブロックに適用] で**、[スクリプトの実行] アクションExcel **オンライン (Business)** コネクタを **追加** します。</span><span class="sxs-lookup"><span data-stu-id="a4c26-144">In the same **Apply to each** block, add another **Excel Online (Business)** connector with the **Run script** action.</span></span> <span data-ttu-id="a4c26-145">次の値を使用します。</span><span class="sxs-lookup"><span data-stu-id="a4c26-145">Use the following values.</span></span>
    1. <span data-ttu-id="a4c26-146">**場所**: OneDrive for Business</span><span class="sxs-lookup"><span data-stu-id="a4c26-146">**Location**: OneDrive for Business</span></span>
    1. <span data-ttu-id="a4c26-147">**ドキュメント ライブラリ**: OneDrive</span><span class="sxs-lookup"><span data-stu-id="a4c26-147">**Document Library**: OneDrive</span></span>
    1. <span data-ttu-id="a4c26-148">**ファイル**: hr-interviews.xlsx *(ファイル ブラウザーから選択)*</span><span class="sxs-lookup"><span data-stu-id="a4c26-148">**File**: hr-interviews.xlsx *(Chosen through the file browser)*</span></span>
    1. <span data-ttu-id="a4c26-149">**スクリプト**: 送信された招待を記録する</span><span class="sxs-lookup"><span data-stu-id="a4c26-149">**Script**: Record Sent Invites</span></span>
    1. **invites**:**結果**(Excel 値) Excel Online (Business) コネクタのスクリーンショットで、招待が Power Automate で送信 :::image type="content" source="../../images/schedule-interviews-3.png" alt-text="されたレコードを記録します。":::
1. <span data-ttu-id="a4c26-151">フローを保存し、試してみてください。[フロー エディター **] ページ** の [テスト] ボタンを使用するか、[マイ フロー] タブでフロー **を実行** します。メッセージが表示されたら、必ずアクセスを許可してください。</span><span class="sxs-lookup"><span data-stu-id="a4c26-151">Save the flow and try it out. Use the **Test** button on the flow editor page or run the flow through your **My flows** tab. Be sure to allow access when prompted.</span></span>

## <a name="training-video-send-a-teams-meeting-from-excel-data"></a><span data-ttu-id="a4c26-152">トレーニング ビデオ: データからTeams会議をExcelする</span><span class="sxs-lookup"><span data-stu-id="a4c26-152">Training video: Send a Teams meeting from Excel data</span></span>

<span data-ttu-id="a4c26-153">[Sudhi Ramamurthy が YouTube でこのサンプル](https://youtu.be/HyBdx52NOE8)のバージョンを見るをご覧ください。</span><span class="sxs-lookup"><span data-stu-id="a4c26-153">[Watch Sudhi Ramamurthy walk through a version of this sample on YouTube](https://youtu.be/HyBdx52NOE8).</span></span> <span data-ttu-id="a4c26-154">彼のバージョンでは、列の変更や廃止された会議の時間を処理する、より堅牢なスクリプトを使用しています。</span><span class="sxs-lookup"><span data-stu-id="a4c26-154">His version uses a more robust script that handles changing columns and obsolete meeting times.</span></span>
