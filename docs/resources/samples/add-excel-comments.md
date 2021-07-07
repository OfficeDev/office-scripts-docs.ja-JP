---
title: コメントを追加Excel
description: ワークシートにコメントを追加Officeスクリプトを使用する方法について説明します。
ms.date: 06/29/2021
localization_priority: Normal
ms.openlocfilehash: 77e308d020281c71751e2652f8dbaec00c263e44
ms.sourcegitcommit: 211c157ca746e266eeb079f5fa1925a1e35ab702
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 07/07/2021
ms.locfileid: "53313912"
---
# <a name="add-comments-in-excel"></a><span data-ttu-id="98038-103">コメントを追加Excel</span><span class="sxs-lookup"><span data-stu-id="98038-103">Add comments in Excel</span></span>

<span data-ttu-id="98038-104">このサンプルでは、同僚のコメントを含むセルに [コメント@mentioning](https://support.microsoft.com/office/90701709-5dc1-41c7-aa48-b01d4a46e8c7) 示します。</span><span class="sxs-lookup"><span data-stu-id="98038-104">This sample shows how to add comments to a cell including [@mentioning](https://support.microsoft.com/office/90701709-5dc1-41c7-aa48-b01d4a46e8c7) a colleague.</span></span>

## <a name="example-scenario"></a><span data-ttu-id="98038-105">シナリオ例</span><span class="sxs-lookup"><span data-stu-id="98038-105">Example scenario</span></span>

* <span data-ttu-id="98038-106">チーム リードはシフト スケジュールを維持します。</span><span class="sxs-lookup"><span data-stu-id="98038-106">The team lead maintains the shift schedule.</span></span> <span data-ttu-id="98038-107">チーム リードは、シフト レコードに従業員 ID を割り当てる。</span><span class="sxs-lookup"><span data-stu-id="98038-107">The team lead assigns an employee ID to the shift record.</span></span>
* <span data-ttu-id="98038-108">チーム リードは、従業員に通知する必要があります。</span><span class="sxs-lookup"><span data-stu-id="98038-108">The team lead wishes to notify the employee.</span></span> <span data-ttu-id="98038-109">従業員にコメントを@mentionsすると、ワークシートからカスタム メッセージが送信されます。</span><span class="sxs-lookup"><span data-stu-id="98038-109">By adding a comment that @mentions the employee, the employee is emailed with a custom message from the worksheet.</span></span>
* <span data-ttu-id="98038-110">その後、従業員はブックを表示し、都合の良い時点でコメントに応答できます。</span><span class="sxs-lookup"><span data-stu-id="98038-110">Subsequently, the employee can view the workbook and respond to the comment at their convenience.</span></span>

## <a name="solution"></a><span data-ttu-id="98038-111">ソリューション</span><span class="sxs-lookup"><span data-stu-id="98038-111">Solution</span></span>

1. <span data-ttu-id="98038-112">スクリプトは、従業員ワークシートから従業員情報を抽出します。</span><span class="sxs-lookup"><span data-stu-id="98038-112">The script extracts employee information from the employee worksheet.</span></span>
1. <span data-ttu-id="98038-113">スクリプトは、シフト レコードの適切なセルにコメント (関連する従業員の電子メールを含む) を追加します。</span><span class="sxs-lookup"><span data-stu-id="98038-113">The script then adds a comment (including the relevant employee email) to the appropriate cell in the shift record.</span></span>
1. <span data-ttu-id="98038-114">セル内の既存のコメントは、新しいコメントを追加する前に削除されます。</span><span class="sxs-lookup"><span data-stu-id="98038-114">Existing comments in the cell are removed before adding the new comment.</span></span>

## <a name="sample-excel-file"></a><span data-ttu-id="98038-115">サンプル Excel ファイル</span><span class="sxs-lookup"><span data-stu-id="98038-115">Sample Excel file</span></span>

<span data-ttu-id="98038-116">すぐに <a href="excel-comments.xlsx"> 使用excel-comments.xlsx</a> ブックのブックをダウンロードします。</span><span class="sxs-lookup"><span data-stu-id="98038-116">Download <a href="excel-comments.xlsx">excel-comments.xlsx</a> for a ready-to-use workbook.</span></span> <span data-ttu-id="98038-117">次のスクリプトを追加して、サンプルを自分で試してみてください。</span><span class="sxs-lookup"><span data-stu-id="98038-117">Add the following script to try the sample yourself!</span></span>

## <a name="sample-code-add-comments"></a><span data-ttu-id="98038-118">サンプル コード: コメントの追加</span><span class="sxs-lookup"><span data-stu-id="98038-118">Sample code: Add comments</span></span>

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

## <a name="training-video-add-comments"></a><span data-ttu-id="98038-119">トレーニング ビデオ: コメントの追加</span><span class="sxs-lookup"><span data-stu-id="98038-119">Training video: Add comments</span></span>

<span data-ttu-id="98038-120">[Sudhi Ramamurthy が YouTube でこのサンプルを歩くのを見る](https://youtu.be/CpR78nkaOFw).</span><span class="sxs-lookup"><span data-stu-id="98038-120">[Watch Sudhi Ramamurthy walk through this sample on YouTube](https://youtu.be/CpR78nkaOFw).</span></span>
