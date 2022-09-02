---
title: 'Office スクリプトのサンプル シナリオ: パンチ クロック ボタン'
description: このサンプルでは、パンチ クロック ボタンを追加し、ユーザーが現在の時刻を使用して出勤および退勤できるようにします。
ms.date: 04/26/2022
ms.localizationpriority: medium
ms.openlocfilehash: ac128a33b653506b6168bd4acfe1713bf6d26759
ms.sourcegitcommit: a6504f8b0d6b717457c6e0b5306c35ad3900914e
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 09/02/2022
ms.locfileid: "67572683"
---
# <a name="office-scripts-sample-scenario-punch-clock-button"></a>Office スクリプトのサンプル シナリオ: パンチ クロック ボタン

このサンプルで使用されるシナリオのアイデアとスクリプトは、Office スクリプト コミュニティ メンバー [の Brian Gonzalez](https://github.com/b-gonzalez) によって提供されました。

このシナリオでは、 [従業員がボタン](../../develop/script-buttons.md)を押して開始時刻と終了時刻を記録できるタイム シートを作成します。 以前に記録された内容に基づいて、ボタンを押すと、1 日の開始 (クロックイン) または 1 日の終了 (退勤) が行われます。 このサンプルは、Excel on the webと Windows の両方で機能します。

:::image type="content" source="../../images/punch-clock-sample-3.png" alt-text="3 つの列 ('Clock In'、'Clock Out'、および 'Duration') と、ブック内の &quot;パンチ クロック&quot; というラベルの付いたボタンを含むテーブル。":::

## <a name="setup-instructions"></a>セットアップ手順

1. [punch-clock-sample.xlsx](punch-clock-sample.xlsx)を OneDrive にダウンロードします。

    :::image type="content" source="../../images/punch-clock-sample-1.png" alt-text="'Clock In'、'Clock Out'、'Duration' の 3 つの列を持つテーブル。":::

1. Excel on the webでブックを開きます。

1. [ **自動化** ] タブで [ **新しいスクリプト** ] を選択し、次のスクリプトをエディターに貼り付けます。

    ```typescript
    /**
     * This script records either the start or end time of a shift, 
     * depending on what is filled out in the table. 
     * It is intended to be used with a Script Button.
     */
    function main(workbook: ExcelScript.Workbook) {
      // Get the first table in the timesheet.
      const timeSheet = workbook.getWorksheet("MyTimeSheet");
      const timeTable = timeSheet.getTables()[0];
    
      // Get the appropriate table columns.
      const clockInColumn = timeTable.getColumnByName("Clock In");
      const clockOutColumn = timeTable.getColumnByName("Clock Out");
      const durationColumn = timeTable.getColumnByName("Duration");
    
      // Get the last rows for the Clock In and Clock Out columns.
      let clockInLastRow = clockInColumn.getRangeBetweenHeaderAndTotal().getLastRow();
      let clockOutLastRow = clockOutColumn.getRangeBetweenHeaderAndTotal().getLastRow();
    
      // Get the current date to use as the start or end time.
      let date: Date = new Date();
    
      // Add the current time to a column based on the state of the table.
      if (clockInLastRow.getValue() as string === "") {
        // If the Clock In column has an empty value in the table, add a start time.
        clockInLastRow.setValue(date.toLocaleString());
      } else if (clockOutLastRow.getValue() as string === "") {
        // If the Clock Out column has an empty value in the table, 
        // add an end time and calculate the shift duration.
        clockOutLastRow.setValue(date.toLocaleString());
        const clockInTime = new Date(clockInLastRow.getValue() as string);
        const clockOutTime  = new Date(clockOutLastRow.getValue() as string);
        const clockDuration = Math.abs((clockOutTime.getTime() - clockInTime.getTime()));
    
        let durationString = getDurationMessage(clockDuration);
        durationColumn.getRangeBetweenHeaderAndTotal().getLastRow().setValue(durationString);
      } else {
        // If both columns are full, add a new row, then add a start time.
        timeTable.addRow()
        clockInLastRow.getOffsetRange(1, 0).setValue(date.toLocaleString());
      }
    }
    
    /**
     * A function to write a time duration as a string.
     */
    function getDurationMessage(delta: number) {
      // Adapted from here:
      // https://stackoverflow.com/questions/13903897/javascript-return-number-of-days-hours-minutes-seconds-between-two-dates
    
      delta = delta / 1000;
      let durationString = "";
    
      let days = Math.floor(delta / 86400);
      delta -= days * 86400;
    
      let hours = Math.floor(delta / 3600) % 24;
      delta -= hours * 3600;
    
      let minutes = Math.floor(delta / 60) % 60;
    
      if (days >= 1) {
        durationString += days;
        durationString += (days > 1 ? " days" : " day");
    
        if (hours >= 1 && minutes >= 1) {
          durationString += ", ";
        }
        else if (hours >= 1 || minutes > 1) {
          durationString += " and ";
        }
      }
    
      if (hours >= 1) {
        durationString += hours;
        durationString += (hours > 1 ? " hours" : " hour");
        if (minutes >= 1) {
          durationString += " and ";
        }
      }
    
      if (minutes >= 1) {
        durationString += minutes;
        durationString += (minutes > 1 ? " minutes" : " minute");
      }
    
      return durationString;
    }
    ```

1. スクリプトの名前を "パンチ クロック" に変更します。

1. スクリプトを保存します。

1. ブックでセル **E2** を選択します。

1. スクリプト ボタンを追加します。 **[スクリプトの詳細**] ページの **[その他のオプション (...)]** メニューに移動し、[**追加] ボタン** を選択します。

    :::image type="content" source="../../images/punch-clock-sample-2.png" alt-text="[その他のオプション] メニューと [追加] ボタン。":::

1. ブックを保存します。

## <a name="run-the-script"></a>スクリプトを実行する

**[パンチ クロック**] ボタンを押してスクリプトを実行します。 以前に入力した内容に応じて、現在の時刻が "Clock In" または "Clock Out" に記録されます。

:::image type="content" source="../../images/punch-clock-sample-3.png" alt-text="ブック内のテーブルと [パンチ クロック] ボタン。":::

> [!NOTE]
> 期間は、1 分を超える場合にのみ記録されます。 "Clock In" 時間を手動で編集して、より長い期間をテストします。
