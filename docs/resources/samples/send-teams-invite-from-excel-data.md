---
title: ExcelデータからTeams会議を送信する
description: Officeスクリプトを使用して、データからTeams会議を送信する方法Excel説明します。
ms.date: 05/06/2021
localization_priority: Normal
ROBOTS: NOINDEX
ms.openlocfilehash: 85b39d7e3d1008dee01e7fe9c690116be1d7e5d8
ms.sourcegitcommit: 4687693f02fc90a57ba30c461f35046e02e6f5fb
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 05/19/2021
ms.locfileid: "52545631"
---
# <a name="send-teams-meeting-from-excel-data"></a><span data-ttu-id="fd1fa-103">Excelデータから会議Teams送信</span><span class="sxs-lookup"><span data-stu-id="fd1fa-103">Send Teams meeting from Excel data</span></span>

<span data-ttu-id="fd1fa-104">このソリューションでは、OfficeスクリプトとPower Automateアクションを使用してファイルから行Excel選択し、それを使用してTeamsの会議の招待を送信し、Excelを更新する方法を示します。</span><span class="sxs-lookup"><span data-stu-id="fd1fa-104">This solution shows how to use Office Scripts and Power Automate actions to select rows from Excel file and use it to send a Teams meeting invite then update Excel.</span></span>

## <a name="example-scenario"></a><span data-ttu-id="fd1fa-105">シナリオ例</span><span class="sxs-lookup"><span data-stu-id="fd1fa-105">Example scenario</span></span>

* <span data-ttu-id="fd1fa-106">人事採用担当者は、Excelファイルで候補者の面接スケジュールを管理します。</span><span class="sxs-lookup"><span data-stu-id="fd1fa-106">An HR recruiter manages the interview schedule of candidates in an Excel file.</span></span>
* <span data-ttu-id="fd1fa-107">採用担当者は、候補者と面接官にTeams会議の招待を送信する必要があります。</span><span class="sxs-lookup"><span data-stu-id="fd1fa-107">The recruiter needs to send the Teams meeting invite to the candidate and interviewers.</span></span> <span data-ttu-id="fd1fa-108">ビジネス ルールは、次の項目を選択します。</span><span class="sxs-lookup"><span data-stu-id="fd1fa-108">The business rules are to select:</span></span>

    <span data-ttu-id="fd1fa-109">(a) 招待がファイル列に記録されているとおりに送信されていない人にのみ招待します。</span><span class="sxs-lookup"><span data-stu-id="fd1fa-109">(a) Invites to only those for whom the invite isn't already sent as recorded in the file column.</span></span>

    <span data-ttu-id="fd1fa-110">(b) 今後の面接日(過去の日付なし)</span><span class="sxs-lookup"><span data-stu-id="fd1fa-110">(b) Interview dates in the future (no past dates).</span></span>

* <span data-ttu-id="fd1fa-111">採用担当者は、対象となるレコードに対してすべてのTeams会議が送信されたことを確認して、Excel ファイルを更新する必要があります。</span><span class="sxs-lookup"><span data-stu-id="fd1fa-111">The recruiter needs to update the Excel file with the confirmation that all Teams meetings have been sent for the eligible records.</span></span>

<span data-ttu-id="fd1fa-112">このソリューションには、次の 3 つの部分があります。</span><span class="sxs-lookup"><span data-stu-id="fd1fa-112">The solution has 3 parts:</span></span>

1. <span data-ttu-id="fd1fa-113">Office条件に基づいてテーブルからデータを抽出し、オブジェクトの配列を JSON データとして返すスクリプト。</span><span class="sxs-lookup"><span data-stu-id="fd1fa-113">Office Script to extract data from a table based on conditions and returns an array of objects as JSON data.</span></span>
1. <span data-ttu-id="fd1fa-114">その後、データは Teamsに送信されます Teams **会議を作成** するアクションを招待を送信します。</span><span class="sxs-lookup"><span data-stu-id="fd1fa-114">The data is then sent to the Teams **Create a Teams meeting** action to send invites.</span></span> <span data-ttu-id="fd1fa-115">JSON 配列内のインスタンスごとに 1 つのTeams会議を送信します。</span><span class="sxs-lookup"><span data-stu-id="fd1fa-115">Send one Teams meeting per instance in the JSON array.</span></span>
1. <span data-ttu-id="fd1fa-116">同じ JSON データを別のOfficeスクリプトに送信して、招待の状態を更新します。</span><span class="sxs-lookup"><span data-stu-id="fd1fa-116">Send the same JSON data to another Office Script to update the status of the invitation.</span></span>

## <a name="sample-excel-file"></a><span data-ttu-id="fd1fa-117">サンプル Excel ファイル</span><span class="sxs-lookup"><span data-stu-id="fd1fa-117">Sample Excel file</span></span>

<span data-ttu-id="fd1fa-118">このソリューションで使用 <a href="hr-schedule.xlsx">hr-schedule.xlsx</a> ファイルをダウンロードして、自分で試してみてください!</span><span class="sxs-lookup"><span data-stu-id="fd1fa-118">Download the file <a href="hr-schedule.xlsx">hr-schedule.xlsx</a> used in this solution and try it out yourself!</span></span>

## <a name="sample-code-select-filtered-rows-from-table-as-json"></a><span data-ttu-id="fd1fa-119">サンプル コード: JSON としてテーブルからフィルター処理された行を選択します。</span><span class="sxs-lookup"><span data-stu-id="fd1fa-119">Sample code: Select filtered rows from table as JSON</span></span>

```TypeScript
function main(workbook: ExcelScript.Workbook): InterviewInvite[] {
  console.log("Current date time: " + new Date().toUTCString());
  const MEETING_DURATION = workbook.getNamedItem('MeetingDuration').getRange().getValue() as number;

  // Get the interview candidate information.
  const sheet = workbook.getWorksheet('Interviews');
  const table = sheet.getTables()[0];
  const dataRows: string[][] = table.getRangeBetweenHeaderAndTotal().getTexts();

  // Convert the table rows into InterviewInvite objects for the flow.
  const recordDetails: RecordDetail[] = returnObjectFromValues(dataRows);
  const inviteRecords = generateInterviewRecords(recordDetails, MEETING_DURATION);
  console.log(JSON.stringify(inviteRecords));
  return inviteRecords;
}

/**
 * Converts table values into a RecordDetail array.
 */
function returnObjectFromValues(values: string[][]): RecordDetail[] {
  let objectArray: BasicObj[] = [];
  let objectKeys: string[] = [];
  for (let i = 0; i < values.length; i++) {
    if (i === 0) {
      objectKeys = values[i]
      continue;
    }

    let object = {}
    for (let j = 0; j < values[i].length; j++) {
      object[objectKeys[j]] = values[i][j]
    }
    objectArray.push(object);
  }
  return objectArray as RecordDetail[];
}

/**
 * Generate interview records by selecting required columns.
 * @param records Input records from the table of interviews.
 * @param mins Number of minutes to add to the start date-time.
 */
function generateInterviewRecords(records: RecordDetail[], mins: number): InterviewInvite[] {
  const interviewInvites: InterviewInvite[] = [];

  records.forEach((record) => {
    // Interviewer 1
    // If the start date-time is greater than current date-time, add to output records.
    if ((new Date(record['Start time1'])) > new Date()) {
      console.log("selected " + new Date(record['Start time1']).toUTCString());
      let startTime = new Date(record['Start time1']).toISOString();
      // Compute the finish time of the meeting.
      let finishTime = addMins(new Date(record['Start time1']), mins).toISOString();
      interviewInvites.push({
        ID: record.ID,
        Candidate: record.Candidate,
        CandidateEmail: record['Candidate email'] as string,
        CandidateContact: record['Candidate contact'] as string,
        Interviewer: record.Interviewer1,
        InterviewerEmail: record['Interviewer1 email'],
        StartTime: startTime,
        FinishTime: finishTime
      });
    } else {
      console.log("Rejected " + (new Date(record['Start time1']).toUTCString()));
    }
    // Interviewer 2 
    // If the start date-time is greater than current date-time, add to output records.
    if ((new Date(record['Start time2'])) > new Date()) {
      console.log("selected " + new Date(record['Start time2']).toUTCString());


      let startTime = new Date(record['Start time2']).toISOString();
      // Compute the finish time of the meeting.
      let finishTime = addMins(new Date(record['Start time2']), mins).toISOString();
      interviewInvites.push({
        ID: record.ID,
        Candidate: record.Candidate,
        CandidateEmail: record['Candidate email'] as string,
        CandidateContact: record['Candidate contact'] as string,
        Interviewer: record.Interviewer2,
        InterviewerEmail: record['Interviewer2 email'],
        StartTime: startTime,
        FinishTime: finishTime
      })
    } else {
      console.log("Rejected " + (new Date(record['Start time2']).toUTCString()))

    }
  })
  return interviewInvites;
}

/**
 * Add minutes to start date-time.
 * @param startDateTime Start date-time
 * @param mins Minutes to add to the start date-time
 */
function addMins(startDateTime: Date, mins: number) {
  return new Date(startDateTime.getTime() + mins * 60 * 1000);
}

// Basic key-value pair object.
interface BasicObj {
  [key: string]: string | number | boolean
}

// Input record that matches the table data.
interface RecordDetail extends BasicObj {
  ID: string
  'Invite to interview': string
  Candidate: string
  'Candidate email': string
  'Candidate contact': string
  Interviewer1: string
  'Interviewer1 email': string
  Interviewer2: string
  'Interviewer2 email': string
  'Start time1': string
  'Start time2': string
}

// Output record.
interface InterviewInvite extends BasicObj {
  ID: string
  Candidate: string
  CandidateEmail: string
  CandidateContact: string
  Interviewer: string
  InterviewerEmail: string
  StartTime: string
  FinishTime: string
}
```

## <a name="sample-code-mark-as-invited"></a><span data-ttu-id="fd1fa-120">サンプル コード: 招待済みとしてマークする</span><span class="sxs-lookup"><span data-stu-id="fd1fa-120">Sample code: Mark as invited</span></span>

```TypeScript
function main(workbook: ExcelScript.Workbook, completedInvitesString: string) {
    completedInvitesString = `[
      {
        "ID": "10",
        "Candidate": "Adele ",
        "CandidateEmail": "AdeleV@M365x904181.OnMicrosoft.com",
        "CandidateContact": "1234567899",
        "Interviewer": "Megan",
        "InterviewerEmail": "MeganB@M365x904181.OnMicrosoft.com",
        "StartTime": "2020-11-03T18:30:00Z",
        "FinishTime": "2020-11-03T22:45:00Z"
      },
      {
        "ID": "30",
        "Candidate": "Allan ",
        "CandidateEmail": "AllanD@M365x904181.OnMicrosoft.com",
        "CandidateContact": "1234567978",
        "Interviewer": "Raul",
        "InterviewerEmail": "RaulR@M365x904181.OnMicrosoft.com",
        "StartTime": "2020-11-03T23:00:00Z",
        "FinishTime": "2020-11-03T23:45:00Z"
      }
    ]`;
    let completedInvites = JSON.parse(completedInvitesString) as InterviewInvite[];
    const sheet = workbook.getWorksheet('Interviews');
    const range = sheet.getTables()[0].getRange();
    const dataRows = range.getValues();
    for (let i=0; i < dataRows.length; i++) {
        for (let invite of completedInvites) {
            if (String(dataRows[i][0]) === invite.ID) {
                range.getCell(i,1).setValue(true);
            }
        }
    }
    return;
}


// Invite record.
interface InterviewInvite  {
    ID: string
    Candidate: string
    CandidateEmail: string
    CandidateContact: string
    Interviewer: string
    InterviewerEmail: string
    StartTime: string
    FinishTime: string
}
```

## <a name="training-video-send-a-teams-meeting-from-excel-data"></a><span data-ttu-id="fd1fa-121">トレーニング ビデオ: Excel データからTeams会議を送信する</span><span class="sxs-lookup"><span data-stu-id="fd1fa-121">Training video: Send a Teams meeting from Excel data</span></span>

<span data-ttu-id="fd1fa-122">[スーディ・ラマムルティがこのサンプルをYouTubeで歩くのを見てください](https://youtu.be/HyBdx52NOE8)。</span><span class="sxs-lookup"><span data-stu-id="fd1fa-122">[Watch Sudhi Ramamurthy walk through this sample on YouTube](https://youtu.be/HyBdx52NOE8).</span></span>
