---
title: 'Office スクリプトサンプルシナリオ: 成績計算ツール'
description: 学生のクラスのパーセンテージおよび文字成績を決定するサンプルです。
ms.date: 02/20/2020
localization_priority: Normal
ms.openlocfilehash: 0db6f7c116594f7655bfc0adc8f5a79dbbf2a0af
ms.sourcegitcommit: b075eed5a6f275274fbbf6d62633219eac416f26
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 03/10/2020
ms.locfileid: "42700401"
---
# <a name="office-scripts-sample-scenario-grade-calculator"></a><span data-ttu-id="8b634-103">Office スクリプトサンプルシナリオ: 成績計算ツール</span><span class="sxs-lookup"><span data-stu-id="8b634-103">Office Scripts sample scenario: Grade calculator</span></span>

<span data-ttu-id="8b634-104">このシナリオでは、生徒のすべての学生の期末等級を集計しています。</span><span class="sxs-lookup"><span data-stu-id="8b634-104">In this scenario, you're an instructor tallying every student's end-of-term grades.</span></span> <span data-ttu-id="8b634-105">自分の割り当てとテストのスコアを入力したことがあります。</span><span class="sxs-lookup"><span data-stu-id="8b634-105">You've been entering the scores for their assignments and tests as you go.</span></span> <span data-ttu-id="8b634-106">ここで、学生の fates を決定します。</span><span class="sxs-lookup"><span data-stu-id="8b634-106">Now, it is time to determine the students' fates.</span></span>

<span data-ttu-id="8b634-107">各ポイントカテゴリの成績を合計するスクリプトを開発します。</span><span class="sxs-lookup"><span data-stu-id="8b634-107">You'll develop a script that totals the grades for each point category.</span></span> <span data-ttu-id="8b634-108">その後、合計数に基づいて各生徒にレター成績を割り当てます。</span><span class="sxs-lookup"><span data-stu-id="8b634-108">It will then assign a letter grade to each student based on the total.</span></span> <span data-ttu-id="8b634-109">精度を高めるために、個別のスコアが低すぎるか高すぎるかを確認するための2つのチェックを追加します。</span><span class="sxs-lookup"><span data-stu-id="8b634-109">To help ensure accuracy, you'll add a couple checks to see if any individual scores are too low or high.</span></span> <span data-ttu-id="8b634-110">生徒のスコアが0未満であるか、または可能な point 値よりも大きい場合、スクリプトは、そのセルに赤の塗りつぶしを設定し、生徒のポイントを合計しないようにします。</span><span class="sxs-lookup"><span data-stu-id="8b634-110">If a student's score is less than zero or more than the possible point value, the script will flag the cell with a red fill and not total that student's points.</span></span> <span data-ttu-id="8b634-111">これは、どのレコードが再チェックする必要があるかを明確に示します。</span><span class="sxs-lookup"><span data-stu-id="8b634-111">This will be a clear indication of which records you need to double-check.</span></span> <span data-ttu-id="8b634-112">また、クラスの上部と下部をすばやく表示できるように、いくつかの基本的な書式設定を成績に追加します。</span><span class="sxs-lookup"><span data-stu-id="8b634-112">You'll also add some basic formatting to the grades so you can quickly view the top and bottom of the class.</span></span>

## <a name="scripting-skills-covered"></a><span data-ttu-id="8b634-113">スクリプト作成スキルの説明</span><span class="sxs-lookup"><span data-stu-id="8b634-113">Scripting skills covered</span></span>

- <span data-ttu-id="8b634-114">セルの書式設定</span><span class="sxs-lookup"><span data-stu-id="8b634-114">Cell formatting</span></span>
- <span data-ttu-id="8b634-115">エラーチェック</span><span class="sxs-lookup"><span data-stu-id="8b634-115">Error checking</span></span>
- <span data-ttu-id="8b634-116">正規表現</span><span class="sxs-lookup"><span data-stu-id="8b634-116">Regular expressions</span></span>

## <a name="setup-instructions"></a><span data-ttu-id="8b634-117">セットアップの手順</span><span class="sxs-lookup"><span data-stu-id="8b634-117">Setup instructions</span></span>

1. <span data-ttu-id="8b634-118"><a href="grade-calculator.xlsx">Grade-calculator</a>を OneDrive にダウンロードします。</span><span class="sxs-lookup"><span data-stu-id="8b634-118">Download <a href="grade-calculator.xlsx">grade-calculator.xlsx</a> to your OneDrive.</span></span>

2. <span data-ttu-id="8b634-119">Web 用の Excel でブックを開きます。</span><span class="sxs-lookup"><span data-stu-id="8b634-119">Open the workbook with Excel for the web.</span></span>

3. <span data-ttu-id="8b634-120">[**自動化**] タブで、**コードエディター**を開きます。</span><span class="sxs-lookup"><span data-stu-id="8b634-120">Under the **Automate** tab, open the **Code Editor**.</span></span>

4. <span data-ttu-id="8b634-121">[**コードエディター** ] 作業ウィンドウで、[**新しいスクリプト**] をクリックし、次のスクリプトをエディターに貼り付けます。</span><span class="sxs-lookup"><span data-stu-id="8b634-121">In the **Code Editor** task pane, press **New Script** and paste the following script into the editor.</span></span>

    ```TypeScript
    async function main(context: Excel.RequestContext) {
      // Get the number of student record rows.
      let sheet = context.workbook.worksheets.getActiveWorksheet();
      let studentsRange = sheet.getUsedRange().load("values, rowCount");
      await context.sync();
      console.log("Total students: " + (studentsRange.rowCount - 1));

      // Clean up any formatting from previous runs of the script.
      studentsRange.clear(Excel.ClearApplyTo.formats);
      studentsRange.getColumn(4).getCell(0, 0).getRowsBelow(studentsRange.rowCount - 1).clear(Excel.ClearApplyTo.all);
      studentsRange.getColumn(5).getCell(0, 0).getRowsBelow(studentsRange.rowCount - 1).clear(Excel.ClearApplyTo.all);
      await context.sync();

      // Parse the headers for the maximum possible scores for each category.
      // The format is `category (score)`.
      let assignmentsMax = studentsRange.values[0][1].match(/\d+/)[0];
      let midTermMax = studentsRange.values[0][2].match(/\d+/)[0];
      let finalsMax = studentsRange.values[0][3].match(/\d+/)[0];
      console.log("Assignments max score:" + assignmentsMax);
      console.log("Mid-term max score: " + midTermMax);
      console.log("Final max score: " + finalsMax);

      // Look at every student row.
      for (let i = 1; i < studentsRange.values.length; i++) {
        let row = studentsRange.values[i];
        let total = row[1] + row[2] + row[3];
        let valid = true;

        // Look for any records that are too low or too high.
        if (row[1] < 0 || row[1] > assignmentsMax) {
          studentsRange.getCell(i, 1).format.fill.color = "Red";
          valid = false;
        }
        if (row[2] < 0 || row[2] > midTermMax) {
          studentsRange.getCell(i, 2).format.fill.color = "Red";
          valid = false;
        }
        if (row[3] < 0 || row[3] > finalsMax) {
          studentsRange.getCell(i, 3).format.fill.color = "Red";
          valid = false;
        }

        // If the scores are valid, total that student's points and assign them a letter grade.
        if (valid) {
          let grade: string;
          switch (true) {
            case total < 60:
              grade = "E";
              break;
            case total < 70:
              grade = "D";
              break;
            case total < 80:
              grade = "C";
              break;
            case total < 90:
              grade = "B";
              break;
            default:
              grade = "A";
              break;
          }

          studentsRange.getCell(i, 4).values = [[total]];
          studentsRange.getCell(i, 5).values = [[grade]];

          // Highlight excellent students and those in need of attention.
          if (grade === "A") {
            studentsRange.getCell(i, 5).format.fill.color = "Green";
          } else if (grade === "E" || grade === "D") {
            studentsRange.getCell(i, 5).format.fill.color = "Orange";
          }
        }
      }

      studentsRange.getColumn(5).format.horizontalAlignment = "Center";
    }
    ```

5. <span data-ttu-id="8b634-122">スクリプトの名前を [**成績電卓**] に変更し、保存します。</span><span class="sxs-lookup"><span data-stu-id="8b634-122">Rename the script to **Grade Calculator** and save it.</span></span>

## <a name="running-the-script"></a><span data-ttu-id="8b634-123">スクリプトを実行する</span><span class="sxs-lookup"><span data-stu-id="8b634-123">Running the script</span></span>

<span data-ttu-id="8b634-124">唯一のワークシートで**成績計算ツール**のスクリプトを実行します。</span><span class="sxs-lookup"><span data-stu-id="8b634-124">Run the **Grade Calculator** script on the only worksheet.</span></span> <span data-ttu-id="8b634-125">このスクリプトは成績を合計し、各学生にレターの成績を割り当てます。</span><span class="sxs-lookup"><span data-stu-id="8b634-125">The script will total the grades and assign each student a letter grade.</span></span> <span data-ttu-id="8b634-126">個々の成績の数が、割り当てまたはテストの価値を超える場合は、問題のある成績が赤で示され、合計が計算されません。</span><span class="sxs-lookup"><span data-stu-id="8b634-126">If any individual grades have more points than the assignment or test is worth, then the offending grade is marked red and the total is not calculated.</span></span>

### <a name="before-running-the-script"></a><span data-ttu-id="8b634-127">スクリプトを実行する前に</span><span class="sxs-lookup"><span data-stu-id="8b634-127">Before running the script</span></span>

![生徒のスコアの行を示すワークシート。](../../images/scenario-grade-calculator-before.png)

### <a name="after-running-the-script"></a><span data-ttu-id="8b634-129">スクリプトを実行した後</span><span class="sxs-lookup"><span data-stu-id="8b634-129">After running the script</span></span>

![有効な生徒の行の赤の合計で、無効なセルの生徒スコアデータを示すワークシート。](../../images/scenario-grade-calculator-after.png)
