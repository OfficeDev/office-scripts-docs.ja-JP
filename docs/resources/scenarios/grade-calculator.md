---
title: 'Officeスクリプトのサンプル シナリオ: 成績計算'
description: 学生のクラスのパーセンテージとレターの成績を決定するサンプル。
ms.date: 12/17/2020
localization_priority: Normal
ms.openlocfilehash: e2ef6e7522fc88219bf6ba40900a1ecceecb263b
ms.sourcegitcommit: f7a7aebfb687f2a35dbed07ed62ff352a114525a
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 05/06/2021
ms.locfileid: "52232698"
---
# <a name="office-scripts-sample-scenario-grade-calculator"></a><span data-ttu-id="51d1e-103">Officeスクリプトのサンプル シナリオ: 成績計算</span><span class="sxs-lookup"><span data-stu-id="51d1e-103">Office Scripts sample scenario: Grade calculator</span></span>

<span data-ttu-id="51d1e-104">このシナリオでは、すべての学生の学期終了の成績を集計する講師です。</span><span class="sxs-lookup"><span data-stu-id="51d1e-104">In this scenario, you're an instructor tallying every student's end-of-term grades.</span></span> <span data-ttu-id="51d1e-105">課題とテストのスコアを入力しています。</span><span class="sxs-lookup"><span data-stu-id="51d1e-105">You've been entering the scores for their assignments and tests as you go.</span></span> <span data-ttu-id="51d1e-106">それでは、学生の運命を決定する時間です。</span><span class="sxs-lookup"><span data-stu-id="51d1e-106">Now, it is time to determine the students' fates.</span></span>

<span data-ttu-id="51d1e-107">各ポイント カテゴリの成績を合計するスクリプトを開発します。</span><span class="sxs-lookup"><span data-stu-id="51d1e-107">You'll develop a script that totals the grades for each point category.</span></span> <span data-ttu-id="51d1e-108">その後、合計に基づいて各学生にレターグレードを割り当てる。</span><span class="sxs-lookup"><span data-stu-id="51d1e-108">It will then assign a letter grade to each student based on the total.</span></span> <span data-ttu-id="51d1e-109">精度を確保するために、複数のチェックを追加して、個々のスコアが低すぎるか高すぎるか確認します。</span><span class="sxs-lookup"><span data-stu-id="51d1e-109">To help ensure accuracy, you'll add a couple checks to see if any individual scores are too low or high.</span></span> <span data-ttu-id="51d1e-110">学生のスコアが 0 未満または指定できるポイント値より大きい場合、スクリプトはセルに赤い塗りつぶしを付け、その学生のポイントの合計にはフラグを設定します。</span><span class="sxs-lookup"><span data-stu-id="51d1e-110">If a student's score is less than zero or more than the possible point value, the script will flag the cell with a red fill and not total that student's points.</span></span> <span data-ttu-id="51d1e-111">これは、ダブル チェックする必要があるレコードを明確に示します。</span><span class="sxs-lookup"><span data-stu-id="51d1e-111">This will be a clear indication of which records you need to double-check.</span></span> <span data-ttu-id="51d1e-112">また、クラスの上部と下部をすばやく表示するために、いくつかの基本的な書式を成績に追加します。</span><span class="sxs-lookup"><span data-stu-id="51d1e-112">You'll also add some basic formatting to the grades so you can quickly view the top and bottom of the class.</span></span>

## <a name="scripting-skills-covered"></a><span data-ttu-id="51d1e-113">スクリプティングのスキルをカバー</span><span class="sxs-lookup"><span data-stu-id="51d1e-113">Scripting skills covered</span></span>

- <span data-ttu-id="51d1e-114">セルの書式設定</span><span class="sxs-lookup"><span data-stu-id="51d1e-114">Cell formatting</span></span>
- <span data-ttu-id="51d1e-115">エラー チェック</span><span class="sxs-lookup"><span data-stu-id="51d1e-115">Error checking</span></span>
- <span data-ttu-id="51d1e-116">正規表現</span><span class="sxs-lookup"><span data-stu-id="51d1e-116">Regular expressions</span></span>
- <span data-ttu-id="51d1e-117">条件付き書式</span><span class="sxs-lookup"><span data-stu-id="51d1e-117">Conditional formatting</span></span>

## <a name="setup-instructions"></a><span data-ttu-id="51d1e-118">セットアップ手順</span><span class="sxs-lookup"><span data-stu-id="51d1e-118">Setup instructions</span></span>

1. <span data-ttu-id="51d1e-119">ユーザー <a href="grade-calculator.xlsx">grade-calculator.xlsx</a>にダウンロードOneDrive。</span><span class="sxs-lookup"><span data-stu-id="51d1e-119">Download <a href="grade-calculator.xlsx">grade-calculator.xlsx</a> to your OneDrive.</span></span>

2. <span data-ttu-id="51d1e-120">Web 用のブックExcel開きます。</span><span class="sxs-lookup"><span data-stu-id="51d1e-120">Open the workbook with Excel for the web.</span></span>

3. <span data-ttu-id="51d1e-121">[自動化] **タブで** 、[すべてのスクリプト] **を開きます**。</span><span class="sxs-lookup"><span data-stu-id="51d1e-121">Under the **Automate** tab, open **All Scripts**.</span></span>

4. <span data-ttu-id="51d1e-122">[コード **エディター] 作業ウィンドウ** で、[新しいスクリプト] **を押** して、次のスクリプトをエディターに貼り付けます。</span><span class="sxs-lookup"><span data-stu-id="51d1e-122">In the **Code Editor** task pane, press **New Script** and paste the following script into the editor.</span></span>

    ```TypeScript
    function main(workbook: ExcelScript.Workbook) {
      // Get the worksheet and validate the data.
      let studentsRange = workbook.getActiveWorksheet().getUsedRange();
      if (studentsRange.getColumnCount() !== 6) {
        throw new Error(`The required columns are not present. Expected column headers: "Student ID | Assignment score | Mid-term | Final | Total | Grade"`);
      }

      let studentData = studentsRange.getValues();

      // Clear the total and grade columns.
      studentsRange.getColumn(4).getCell(1, 0).getAbsoluteResizedRange(studentData.length - 1, 2).clear();

      // Clear all conditional formatting.
      workbook.getActiveWorksheet().getUsedRange().clearAllConditionalFormats();

      // Use regular expressions to read the max score from the assignment, mid-term, and final scores columns.
      let maxScores: string[] = [];
      const assignmentMaxMatches = (studentData[0][1] as string).match(/\d+/);
      const midtermMaxMatches = (studentData[0][2] as string).match(/\d+/);
      const finalMaxMatches = (studentData[0][3] as string).match(/\d+/);

      // Check the matches happened before proceeding.
      if (!(assignmentMaxMatches && midtermMaxMatches && finalMaxMatches)) {
        throw new Error(`The scores are not present in the column headers. Expected format: "Assignments (n)|Mid-term (n)|Final (n)"`);
      }

      // Use the first (and only) match from the regular expressions as the max scores.
      maxScores = [assignmentMaxMatches[0], midtermMaxMatches[0], finalMaxMatches[0]];

      // Set conditional formatting for each of the assignment, mid-term, and final scores columns.
      maxScores.forEach((score, i) => {
        let range = studentsRange.getColumn(i + 1).getCell(0, 0).getRowsBelow(studentData.length - 1);
        setCellValueConditionalFormatting(
          score,
          range,
          "#9C0006",
          "#FFC7CE",
          ExcelScript.ConditionalCellValueOperator.greaterThan
        )
      });

      // Store the current range information to avoid calling the workbook in the loop.
      let studentsRangeFormulas = studentsRange.getColumn(4).getFormulasR1C1();
      let studentsRangeValues = studentsRange.getColumn(5).getValues();

      /* Iterate over each of the student rows and compute the total score and letter grade.
      * Note that iterator starts at index 1 to skip first (header) row.
      */
      for (let i = 1; i < studentData.length; i++) {
        // If any of the scores are invalid, skip processing it.
        if (studentData[i][1] > maxScores[0] ||
          studentData[i][2] > maxScores[1] ||
          studentData[i][3] > maxScores[2]) {
          continue;
        }
        const total = (studentData[i][1] as number) + (studentData[i][2] as number) + (studentData[i][3] as number);
        let grade: string;
        switch (true) {
          case total < 60:
            grade = "F";
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
    
        // Set total score formula.
        studentsRangeFormulas[i][0] = '=RC[-2]+RC[-1]';
        // Set grade cell.
        studentsRangeValues[i][0] = grade;
      }

      // Set the formulas and values outside the loop.
      studentsRange.getColumn(4).setFormulasR1C1(studentsRangeFormulas);
      studentsRange.getColumn(5).setValues(studentsRangeValues);

      // Put a conditional formatting on the grade column.
      let totalRange = studentsRange.getColumn(5).getCell(0, 0).getRowsBelow(studentData.length - 1);
      setCellValueConditionalFormatting(
        "A",
        totalRange,
        "#001600",
        "#C6EFCE",
        ExcelScript.ConditionalCellValueOperator.equalTo
      );
      ["D", "F"].forEach((grade) => {
        setCellValueConditionalFormatting(
          grade,
          totalRange,
          "#443300",
          "#FFEE22",
          ExcelScript.ConditionalCellValueOperator.equalTo
        );
      })
      // Center the grade column.
      studentsRange.getColumn(5).getFormat().setHorizontalAlignment(ExcelScript.HorizontalAlignment.center);
    }

    /**
     * Helper function to apply conditional formatting.
     * @param value Cell value to use in conditional formatting formula1.
     * @param range Target range.
     * @param fontColor Font color to use.
     * @param fillColor Fill color to use.
     * @param operator Operator to use in conditional formatting.
     */
    function setCellValueConditionalFormatting(
      value: string,
      range: ExcelScript.Range,
      fontColor: string,
      fillColor: string,
      operator: ExcelScript.ConditionalCellValueOperator) {
      // Determine the formula1 based on the type of value parameter.
      let formula1: string;
      if (isNaN(Number(value))) {
        // For cell value equalTo rule, use this format: formula1: "=\"A\"",
        formula1 = `=\"${value}\"`;
      } else {
        // For number input (greater-than or less-than rules), just append '='.
        formula1 = `=${value}`;
      }

      // Apply conditional formatting.
      let conditionalFormatting: ExcelScript.ConditionalFormat;
      conditionalFormatting = range.addConditionalFormat(ExcelScript.ConditionalFormatType.cellValue);
      conditionalFormatting.getCellValue().getFormat().getFont().setColor(fontColor);
      conditionalFormatting.getCellValue().getFormat().getFill().setColor(fillColor);
      conditionalFormatting.getCellValue().setRule({ formula1, operator });
    }
    ```

5. <span data-ttu-id="51d1e-123">スクリプトの名前を [ **成績計算] に変更** し、保存します。</span><span class="sxs-lookup"><span data-stu-id="51d1e-123">Rename the script to **Grade Calculator** and save it.</span></span>

## <a name="running-the-script"></a><span data-ttu-id="51d1e-124">スクリプトを実行する</span><span class="sxs-lookup"><span data-stu-id="51d1e-124">Running the script</span></span>

<span data-ttu-id="51d1e-125">唯一の **ワークシートで成績計算** スクリプトを実行します。</span><span class="sxs-lookup"><span data-stu-id="51d1e-125">Run the **Grade Calculator** script on the only worksheet.</span></span> <span data-ttu-id="51d1e-126">スクリプトは、成績を合計し、各学生にレターグレードを割り当てる。</span><span class="sxs-lookup"><span data-stu-id="51d1e-126">The script will total the grades and assign each student a letter grade.</span></span> <span data-ttu-id="51d1e-127">個々の成績に割り当てまたはテストよりも多くのポイントがある場合、問題のある成績は赤でマークされ、合計は計算されません。</span><span class="sxs-lookup"><span data-stu-id="51d1e-127">If any individual grades have more points than the assignment or test is worth, then the offending grade is marked red and the total is not calculated.</span></span> <span data-ttu-id="51d1e-128">また、'A' の成績は緑色で強調表示され、'D' と 'F' の成績は黄色で強調表示されます。</span><span class="sxs-lookup"><span data-stu-id="51d1e-128">Also, any 'A' grades are highlighted in green, while 'D' and 'F' grades are highlighted in yellow.</span></span>

### <a name="before-running-the-script"></a><span data-ttu-id="51d1e-129">スクリプトを実行する前に</span><span class="sxs-lookup"><span data-stu-id="51d1e-129">Before running the script</span></span>

:::image type="content" source="../../images/scenario-grade-calculator-before.png" alt-text="学生のスコアの行を表示するワークシート":::

### <a name="after-running-the-script"></a><span data-ttu-id="51d1e-131">スクリプトの実行後</span><span class="sxs-lookup"><span data-stu-id="51d1e-131">After running the script</span></span>

:::image type="content" source="../../images/scenario-grade-calculator-after.png" alt-text="有効な学生行の赤い合計で無効なセルを持つ学生スコア データを示すワークシート":::
