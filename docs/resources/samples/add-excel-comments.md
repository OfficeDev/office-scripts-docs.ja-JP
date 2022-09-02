---
title: Excel でコメントを追加する
description: Office スクリプトを使用してワークシートにコメントを追加する方法について説明します。
ms.date: 06/29/2021
ms.localizationpriority: medium
ms.openlocfilehash: 90f072805e6798a4f9d6e74889ccca15610c87bd
ms.sourcegitcommit: a6504f8b0d6b717457c6e0b5306c35ad3900914e
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 09/02/2022
ms.locfileid: "67572494"
---
# <a name="add-comments-in-excel"></a>Excel でコメントを追加する

このサンプルでは、同僚を含むセルにコメント [@mentioning](https://support.microsoft.com/office/90701709-5dc1-41c7-aa48-b01d4a46e8c7) 追加する方法を示します。

## <a name="example-scenario"></a>シナリオ例

* チーム リーダーはシフト スケジュールを維持します。 チーム リーダーは、従業員 ID をシフト レコードに割り当てます。
* チーム リーダーは、従業員に通知することを希望しています。 従業員を@mentionsコメントを追加すると、従業員にワークシートからのカスタム メッセージが電子メールで送信されます。
* その後、従業員はブックを表示し、都合のよくコメントに返信できます。

## <a name="solution"></a>ソリューション

1. このスクリプトは、従業員ワークシートから従業員情報を抽出します。
1. 次に、スクリプトによって、シフト レコードの適切なセルにコメント (関連する従業員のメールを含む) が追加されます。
1. セル内の既存のコメントは、新しいコメントを追加する前に削除されます。

## <a name="sample-excel-file"></a>Excel ファイルのサンプル

すぐに使用できるブックの [excel-comments.xlsx](excel-comments.xlsx) をダウンロードします。 サンプルを自分で試すには、次のスクリプトを追加します。

## <a name="sample-code-add-comments"></a>サンプル コード: コメントを追加する

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

## <a name="training-video-add-comments"></a>トレーニング ビデオ: コメントを追加する

[YouTube でこのサンプルを見る、スディ Ramamurthy のチュートリアルをご覧ください](https://youtu.be/CpR78nkaOFw)。
