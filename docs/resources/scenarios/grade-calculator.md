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
# <a name="office-scripts-sample-scenario-grade-calculator"></a>Office スクリプトサンプルシナリオ: 成績計算ツール

このシナリオでは、生徒のすべての学生の期末等級を集計しています。 自分の割り当てとテストのスコアを入力したことがあります。 ここで、学生の fates を決定します。

各ポイントカテゴリの成績を合計するスクリプトを開発します。 その後、合計数に基づいて各生徒にレター成績を割り当てます。 精度を高めるために、個別のスコアが低すぎるか高すぎるかを確認するための2つのチェックを追加します。 生徒のスコアが0未満であるか、または可能な point 値よりも大きい場合、スクリプトは、そのセルに赤の塗りつぶしを設定し、生徒のポイントを合計しないようにします。 これは、どのレコードが再チェックする必要があるかを明確に示します。 また、クラスの上部と下部をすばやく表示できるように、いくつかの基本的な書式設定を成績に追加します。

## <a name="scripting-skills-covered"></a>スクリプト作成スキルの説明

- セルの書式設定
- エラーチェック
- 正規表現

## <a name="setup-instructions"></a>セットアップの手順

1. <a href="grade-calculator.xlsx">Grade-calculator</a>を OneDrive にダウンロードします。

2. Web 用の Excel でブックを開きます。

3. [**自動化**] タブで、**コードエディター**を開きます。

4. [**コードエディター** ] 作業ウィンドウで、[**新しいスクリプト**] をクリックし、次のスクリプトをエディターに貼り付けます。

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

5. スクリプトの名前を [**成績電卓**] に変更し、保存します。

## <a name="running-the-script"></a>スクリプトを実行する

唯一のワークシートで**成績計算ツール**のスクリプトを実行します。 このスクリプトは成績を合計し、各学生にレターの成績を割り当てます。 個々の成績の数が、割り当てまたはテストの価値を超える場合は、問題のある成績が赤で示され、合計が計算されません。

### <a name="before-running-the-script"></a>スクリプトを実行する前に

![生徒のスコアの行を示すワークシート。](../../images/scenario-grade-calculator-before.png)

### <a name="after-running-the-script"></a>スクリプトを実行した後

![有効な生徒の行の赤の合計で、無効なセルの生徒スコアデータを示すワークシート。](../../images/scenario-grade-calculator-after.png)
