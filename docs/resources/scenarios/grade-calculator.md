---
title: 'Office スクリプトサンプルシナリオ: 成績計算ツール'
description: 学生のクラスのパーセンテージおよび文字成績を決定するサンプルです。
ms.date: 07/24/2020
localization_priority: Normal
ms.openlocfilehash: 4e488c6cc67bda9122b88c55070654632d9c7fa2
ms.sourcegitcommit: ff7fde04ce5a66d8df06ed505951c8111e2e9833
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 08/11/2020
ms.locfileid: "46616743"
---
# <a name="office-scripts-sample-scenario-grade-calculator"></a><span data-ttu-id="f5695-103">Office スクリプトサンプルシナリオ: 成績計算ツール</span><span class="sxs-lookup"><span data-stu-id="f5695-103">Office Scripts sample scenario: Grade calculator</span></span>

<span data-ttu-id="f5695-104">このシナリオでは、生徒のすべての学生の期末等級を集計しています。</span><span class="sxs-lookup"><span data-stu-id="f5695-104">In this scenario, you're an instructor tallying every student's end-of-term grades.</span></span> <span data-ttu-id="f5695-105">自分の割り当てとテストのスコアを入力したことがあります。</span><span class="sxs-lookup"><span data-stu-id="f5695-105">You've been entering the scores for their assignments and tests as you go.</span></span> <span data-ttu-id="f5695-106">ここで、学生の fates を決定します。</span><span class="sxs-lookup"><span data-stu-id="f5695-106">Now, it is time to determine the students' fates.</span></span>

<span data-ttu-id="f5695-107">各ポイントカテゴリの成績を合計するスクリプトを開発します。</span><span class="sxs-lookup"><span data-stu-id="f5695-107">You'll develop a script that totals the grades for each point category.</span></span> <span data-ttu-id="f5695-108">その後、合計数に基づいて各生徒にレター成績を割り当てます。</span><span class="sxs-lookup"><span data-stu-id="f5695-108">It will then assign a letter grade to each student based on the total.</span></span> <span data-ttu-id="f5695-109">精度を高めるために、個別のスコアが低すぎるか高すぎるかを確認するための2つのチェックを追加します。</span><span class="sxs-lookup"><span data-stu-id="f5695-109">To help ensure accuracy, you'll add a couple checks to see if any individual scores are too low or high.</span></span> <span data-ttu-id="f5695-110">生徒のスコアが0未満であるか、または可能な point 値よりも大きい場合、スクリプトは、そのセルに赤の塗りつぶしを設定し、生徒のポイントを合計しないようにします。</span><span class="sxs-lookup"><span data-stu-id="f5695-110">If a student's score is less than zero or more than the possible point value, the script will flag the cell with a red fill and not total that student's points.</span></span> <span data-ttu-id="f5695-111">これは、どのレコードが再チェックする必要があるかを明確に示します。</span><span class="sxs-lookup"><span data-stu-id="f5695-111">This will be a clear indication of which records you need to double-check.</span></span> <span data-ttu-id="f5695-112">また、クラスの上部と下部をすばやく表示できるように、いくつかの基本的な書式設定を成績に追加します。</span><span class="sxs-lookup"><span data-stu-id="f5695-112">You'll also add some basic formatting to the grades so you can quickly view the top and bottom of the class.</span></span>

## <a name="scripting-skills-covered"></a><span data-ttu-id="f5695-113">スクリプト作成スキルの説明</span><span class="sxs-lookup"><span data-stu-id="f5695-113">Scripting skills covered</span></span>

- <span data-ttu-id="f5695-114">セルの書式設定</span><span class="sxs-lookup"><span data-stu-id="f5695-114">Cell formatting</span></span>
- <span data-ttu-id="f5695-115">エラーチェック</span><span class="sxs-lookup"><span data-stu-id="f5695-115">Error checking</span></span>
- <span data-ttu-id="f5695-116">正規表現</span><span class="sxs-lookup"><span data-stu-id="f5695-116">Regular expressions</span></span>
- <span data-ttu-id="f5695-117">条件付き書式</span><span class="sxs-lookup"><span data-stu-id="f5695-117">Conditional formatting</span></span>

## <a name="setup-instructions"></a><span data-ttu-id="f5695-118">セットアップの手順</span><span class="sxs-lookup"><span data-stu-id="f5695-118">Setup instructions</span></span>

1. <span data-ttu-id="f5695-119">OneDrive に<a href="grade-calculator.xlsx">grade-calculator.xlsx</a>をダウンロードします。</span><span class="sxs-lookup"><span data-stu-id="f5695-119">Download <a href="grade-calculator.xlsx">grade-calculator.xlsx</a> to your OneDrive.</span></span>

2. <span data-ttu-id="f5695-120">Web 用の Excel でブックを開きます。</span><span class="sxs-lookup"><span data-stu-id="f5695-120">Open the workbook with Excel for the web.</span></span>

3. <span data-ttu-id="f5695-121">[**自動化**] タブで、**コードエディター**を開きます。</span><span class="sxs-lookup"><span data-stu-id="f5695-121">Under the **Automate** tab, open the **Code Editor**.</span></span>

4. <span data-ttu-id="f5695-122">[**コードエディター** ] 作業ウィンドウで、[**新しいスクリプト**] をクリックし、次のスクリプトをエディターに貼り付けます。</span><span class="sxs-lookup"><span data-stu-id="f5695-122">In the **Code Editor** task pane, press **New Script** and paste the following script into the editor.</span></span>

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
      const assignmentMaxMatches = studentData[0][1].match(/\d+/);
      const midtermMaxMatches = studentData[0][2].match(/\d+/);
      const finalMaxMatches = studentData[0][3].match(/\d+/);

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
        const total = studentData[i][1] + studentData[i][2] + studentData[i][3];
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
      let conditionalFormatting : ExcelScript.ConditionalFormat;
      conditionalFormatting = range.addConditionalFormat(ExcelScript.ConditionalFormatType.cellValue);
      conditionalFormatting.getCellValue().getFormat().getFont().setColor(fontColor);
      conditionalFormatting.getCellValue().getFormat().getFill().setColor(fillColor);
      conditionalFormatting.getCellValue().setRule({formula1, operator});
    }
    ```

5. <span data-ttu-id="f5695-123">スクリプトの名前を [**成績電卓**] に変更し、保存します。</span><span class="sxs-lookup"><span data-stu-id="f5695-123">Rename the script to **Grade Calculator** and save it.</span></span>

## <a name="running-the-script"></a><span data-ttu-id="f5695-124">スクリプトを実行する</span><span class="sxs-lookup"><span data-stu-id="f5695-124">Running the script</span></span>

<span data-ttu-id="f5695-125">唯一のワークシートで**成績計算ツール**のスクリプトを実行します。</span><span class="sxs-lookup"><span data-stu-id="f5695-125">Run the **Grade Calculator** script on the only worksheet.</span></span> <span data-ttu-id="f5695-126">このスクリプトは成績を合計し、各学生にレターの成績を割り当てます。</span><span class="sxs-lookup"><span data-stu-id="f5695-126">The script will total the grades and assign each student a letter grade.</span></span> <span data-ttu-id="f5695-127">個々の成績の数が、割り当てまたはテストの価値を超える場合は、問題のある成績が赤で示され、合計が計算されません。</span><span class="sxs-lookup"><span data-stu-id="f5695-127">If any individual grades have more points than the assignment or test is worth, then the offending grade is marked red and the total is not calculated.</span></span> <span data-ttu-id="f5695-128">また、"A" という成績は緑色で強調表示されていますが、「d」と「F」の成績は黄色で強調表示されています。</span><span class="sxs-lookup"><span data-stu-id="f5695-128">Also, any 'A' grades are highlighted in green, while 'D' and 'F' grades are highlighted in yellow.</span></span>

### <a name="before-running-the-script"></a><span data-ttu-id="f5695-129">スクリプトを実行する前に</span><span class="sxs-lookup"><span data-stu-id="f5695-129">Before running the script</span></span>

![生徒のスコアの行を示すワークシート。](../../images/scenario-grade-calculator-before.png)

### <a name="after-running-the-script"></a><span data-ttu-id="f5695-131">スクリプトを実行した後</span><span class="sxs-lookup"><span data-stu-id="f5695-131">After running the script</span></span>

![有効な生徒の行の赤の合計で、無効なセルの生徒スコアデータを示すワークシート。](../../images/scenario-grade-calculator-after.png)
