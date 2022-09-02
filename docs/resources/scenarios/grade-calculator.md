---
title: 'Office スクリプトのサンプル シナリオ: 成績計算ツール'
description: 学生のクラスの割合と文字の成績を決定するサンプル。
ms.date: 06/29/2021
ms.localizationpriority: medium
ms.openlocfilehash: 7dda3ebe84dc3edd10998cbe2c4cd0806da11411
ms.sourcegitcommit: a6504f8b0d6b717457c6e0b5306c35ad3900914e
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 09/02/2022
ms.locfileid: "67572529"
---
# <a name="office-scripts-sample-scenario-grade-calculator"></a>Office スクリプトのサンプル シナリオ: 成績計算ツール

このシナリオでは、すべての学生の最終学年を集計するインストラクターです。 割り当てとテストのスコアを入力しています。 次に、学生の運運を判断します。

各ポイント カテゴリの成績を合計するスクリプトを開発します。 その後、合計に基づいて、各学生にレターの成績が割り当てられます。 精度を確保するために、いくつかのチェックを追加して、個々のスコアが低すぎるか高すぎるかを確認します。 学生のスコアが 0 未満または可能なポイント値より大きい場合、スクリプトは赤い塗りつぶしでセルにフラグを設定し、その学生のポイントの合計は設定しません。 これは、再確認する必要があるレコードを明確に示します。 また、クラスの上部と下部をすばやく表示できるように、いくつかの基本的な書式設定を成績に追加します。

## <a name="scripting-skills-covered"></a>スクリプティング スキルの説明

- セルの書式設定
- エラーチェック
- 正規表現
- 条件付き書式

## <a name="setup-instructions"></a>セットアップ手順

1. [grade-calculator.xlsx](grade-calculator.xlsx)を OneDrive にダウンロードします。

1. Excel for the webを使用してブックを開きます。

1. [ **自動化** ] タブで [ **新しいスクリプト** ] を選択し、次のスクリプトをエディターに貼り付けます。

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

1. スクリプトの名前を **成績計算ツール** に変更して保存します。

## <a name="running-the-script"></a>スクリプトを実行する

唯一のワークシートで **成績計算ツール** スクリプトを実行します。 このスクリプトは成績を合計し、各学生にレターの成績を割り当てます。 課題またはテストの価値を超えるポイントを持つ個々の成績がある場合、問題のある成績は赤でマークされ、合計は計算されません。 また、'A' の成績は緑色で強調表示され、'D' と 'F' の成績は黄色で強調表示されます。

### <a name="before-running-the-script"></a>スクリプトを実行する前に

:::image type="content" source="../../images/scenario-grade-calculator-before.png" alt-text="学生のスコアの行を示すワークシート。":::

### <a name="after-running-the-script"></a>スクリプトを実行した後

:::image type="content" source="../../images/scenario-grade-calculator-after.png" alt-text="有効な学生行の赤の合計に無効なセルを含む学生スコア データを示すワークシート。":::
