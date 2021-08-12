---
title: コメントを追加Excel
description: ワークシートにコメントを追加Officeスクリプトを使用する方法について説明します。
ms.date: 06/29/2021
localization_priority: Normal
ms.openlocfilehash: b57b4112f2092dd83872f63854a8156b2a19384ac1d86a44ab2d9df4e6b7ff7e
ms.sourcegitcommit: 75f7ed8c2d23a104acc293f8ce29ea580b4fcdc5
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 08/11/2021
ms.locfileid: "57847101"
---
# <a name="add-comments-in-excel"></a>コメントを追加Excel

このサンプルでは、同僚のコメントを含むセルに [コメント@mentioning](https://support.microsoft.com/office/90701709-5dc1-41c7-aa48-b01d4a46e8c7) 示します。

## <a name="example-scenario"></a>シナリオ例

* チーム リードはシフト スケジュールを維持します。 チーム リードは、シフト レコードに従業員 ID を割り当てる。
* チーム リードは、従業員に通知する必要があります。 従業員にコメントを@mentionsすると、ワークシートからカスタム メッセージが送信されます。
* その後、従業員はブックを表示し、都合の良い時点でコメントに応答できます。

## <a name="solution"></a>解決方法

1. スクリプトは、従業員ワークシートから従業員情報を抽出します。
1. スクリプトは、シフト レコードの適切なセルにコメント (関連する従業員の電子メールを含む) を追加します。
1. セル内の既存のコメントは、新しいコメントを追加する前に削除されます。

## <a name="sample-excel-file"></a>サンプル Excel ファイル

すぐに <a href="excel-comments.xlsx"> 使用excel-comments.xlsx</a> ブックのブックをダウンロードします。 次のスクリプトを追加して、サンプルを自分で試してみてください。

## <a name="sample-code-add-comments"></a>サンプル コード: コメントの追加

```TypeScript
function main(workbook: ExcelScript.Workbook) {
  // Get the list of employees.
  const employees = workbook.getWorksheet('Employees').getUsedRange().getTexts();
  console.log(employees); 
  
  // Get the schedule information from the schedule table.
  const scheduleSheet = workbook.getWorksheet('Schedule');
  const table = scheduleSheet.getTables()[0];
  const range = table.getRangeBetweenHeaderAndTotal();
  const scheduleData = range.getTexts();

  // Look through the schedule for a matching employee.
  for (let i = 0; i < scheduleData.length; i++) {
    let employeeId = scheduleData[i][3];

    // Compare the employee ID in the schedule against the employee information table.
    let employeeInfo = employees.find(employeeRow => employeeRow[0] === employeeId);
    if (employeeInfo) {
      console.log("Found a match " + employeeInfo);
      let adminNotes = scheduleData[i][4];

      // Look for and delete old comments, so we avoid conflicts.
      let comment = workbook.getCommentByCell(range.getCell(i, 5));
      if (comment) {
        comment.delete();
      }

      // Add a comment using the admin notes as the text.
      workbook.addComment(range.getCell(i,5), {
        mentions: [{
          email: employeeInfo[1],
          id: 0, // This ID maps this mention to the `id=0` text in the comment.
          name: employeeInfo[2]
        }],
        richContent: `<at id=\"0\">${employeeInfo[2]}</at> ${adminNotes}`
      }, ExcelScript.ContentType.mention);        
      
    } else {
      console.log("No match for: " + employeeId);
    }
  }
}
```

## <a name="training-video-add-comments"></a>トレーニング ビデオ: コメントの追加

[Sudhi Ramamurthy が YouTube でこのサンプルを歩くのを見る](https://youtu.be/CpR78nkaOFw).
