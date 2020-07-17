---
title: Excel on the web の Office スクリプト
description: Office スクリプト用の操作レコーダーとコード エディターの概要をご紹介します。
ms.date: 06/29/2020
localization_priority: Priority
ms.openlocfilehash: 046dd4eac0cce14117da75199841f0b2f72031bc
ms.sourcegitcommit: bf9f33c37c6f7805d6b408aa648bb9785a7cd133
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 07/06/2020
ms.locfileid: "45043406"
---
# <a name="office-scripts-in-excel-on-the-web-preview"></a><span data-ttu-id="29a0c-103">Excel on the web の Office スクリプト (プレビュー)</span><span class="sxs-lookup"><span data-stu-id="29a0c-103">Office Scripts in Excel on the web (preview)</span></span>

<span data-ttu-id="29a0c-104">Excel on the web の Office スクリプトを使用すると、日常のタスクを自動化できます。</span><span class="sxs-lookup"><span data-stu-id="29a0c-104">Office Scripts in Excel on the web let you automate your day-to-day tasks.</span></span> <span data-ttu-id="29a0c-105">Excel で行う操作を操作レコーダーで記録すると、スクリプトが作成されます。</span><span class="sxs-lookup"><span data-stu-id="29a0c-105">You can record your Excel actions with the Action Recorder, which creates a script.</span></span> <span data-ttu-id="29a0c-106">さらに、コード エディターでスクリプトの作成や編集をすることもできます。</span><span class="sxs-lookup"><span data-stu-id="29a0c-106">You can also create and edit scripts with the Code Editor.</span></span> <span data-ttu-id="29a0c-107">スクリプトは組織全体で共有できるため、同僚もワークフローを自動化できます。</span><span class="sxs-lookup"><span data-stu-id="29a0c-107">Your scripts can then be shared across your organization so your coworkers can also automate their workflows.</span></span>

<span data-ttu-id="29a0c-108">この一連のドキュメントで、これらのツールの使用方法について説明します。</span><span class="sxs-lookup"><span data-stu-id="29a0c-108">This series of documents teaches you how to use these tools.</span></span> <span data-ttu-id="29a0c-109">操作レコーダーの紹介では、頻繁に実行する Excel 操作の記録方法を説明します。</span><span class="sxs-lookup"><span data-stu-id="29a0c-109">You'll be introduced to the Action Recorder and see how to record your frequent Excel actions.</span></span> <span data-ttu-id="29a0c-110">また、コード エディターを使用して、独自のスクリプトを作成したり更新したりする方法についても説明します。</span><span class="sxs-lookup"><span data-stu-id="29a0c-110">You'll also learn how to make or update your own scripts with the Code Editor.</span></span>

<br>

> [!VIDEO https://www.microsoft.com/videoplayer/embed/RE4qdFF]

[!INCLUDE [Preview note](../includes/preview-note.md)]

## <a name="when-to-use-office-scripts"></a><span data-ttu-id="29a0c-111">Office スクリプトの使用に適した状況</span><span class="sxs-lookup"><span data-stu-id="29a0c-111">When to use Office Scripts</span></span>

<span data-ttu-id="29a0c-112">スクリプトを使用すると、自分が行った Excel の操作を記録して、さまざまなブックやワークシートに対してその操作を再現できます。</span><span class="sxs-lookup"><span data-stu-id="29a0c-112">Scripts allow you to record and replay your Excel actions on different workbooks and worksheets.</span></span> <span data-ttu-id="29a0c-113">同じ操作を何度も繰り返し行う必要がある場合は、Office スクリプトを使用すると、ワークフロー全体を 1 度ボタンを押すだけの操作に短縮できます。</span><span class="sxs-lookup"><span data-stu-id="29a0c-113">If you find yourself doing the same things over and over again, an Office Script can help you by reducing your whole workflow to a single button press.</span></span>

<span data-ttu-id="29a0c-114">たとえば、毎日仕事の始めに Excel で会計サイトから .csv ファイルを開いているとします。</span><span class="sxs-lookup"><span data-stu-id="29a0c-114">As an example, say you start your work day by opening a .csv file from an accounting site in Excel.</span></span> <span data-ttu-id="29a0c-115">それから数分かけて、不要な列を削除し、テーブルの書式を設定し、数式を追加し、新しいワークシートにピボットテーブルを作成します。</span><span class="sxs-lookup"><span data-stu-id="29a0c-115">You then spend several minutes deleting unnecessary columns, formatting a table, adding formulas, and creating a PivotTable in a new worksheet.</span></span> <span data-ttu-id="29a0c-116">毎日繰り返しているこのような操作を、操作レコーダーで 1 回記録できます。</span><span class="sxs-lookup"><span data-stu-id="29a0c-116">Those actions you repeat daily can be recorded once with the Action Recorder.</span></span> <span data-ttu-id="29a0c-117">それ以降は、スクリプトを実行するだけで、.csv の変換処理すべてが自動的に実行されます。</span><span class="sxs-lookup"><span data-stu-id="29a0c-117">From then on, running the script will take care of your entire .csv conversion.</span></span> <span data-ttu-id="29a0c-118">手順を忘れる危険がなくなるだけでなく、特に操作を教えなくても他の人とプロセスを共有することもできます。</span><span class="sxs-lookup"><span data-stu-id="29a0c-118">You'll not only remove the risk of forgetting steps, but be able to share your process with others without having to teach them anything.</span></span> <span data-ttu-id="29a0c-119">Office スクリプトを使用すると一般的なタスクを自動化できるので、自分自身と職場の作業効率や生産性を向上できます。</span><span class="sxs-lookup"><span data-stu-id="29a0c-119">Office Scripts automate your common tasks so you and your workplace can be more efficient and productive.</span></span>

## <a name="action-recorder"></a><span data-ttu-id="29a0c-120">操作レコーダー</span><span class="sxs-lookup"><span data-stu-id="29a0c-120">Action Recorder</span></span>

![いくつかの操作を記録した後の操作レコーダー。](../images/action-recorder-intro.png)

<span data-ttu-id="29a0c-122">操作レコーダーは、ユーザーが Excel で実行した操作を記録し、その操作をスクリプトに変換します。</span><span class="sxs-lookup"><span data-stu-id="29a0c-122">The Action Recorder records actions you take in Excel and translates them into a script.</span></span> <span data-ttu-id="29a0c-123">操作レコーダーを実行すると、セルの編集、書式の変更、テーブルの作成などの Excel の操作をキャプチャできます。</span><span class="sxs-lookup"><span data-stu-id="29a0c-123">With the Action recorder running, you can capture the Excel actions as you edit cells, change formatting, and create tables.</span></span> <span data-ttu-id="29a0c-124">作成されたスクリプトは、他のワークシートやブックで実行して、ユーザーが実行した元の操作を再現することもできます。</span><span class="sxs-lookup"><span data-stu-id="29a0c-124">The resulting script can be run on other worksheets and workbooks to recreate your original actions.</span></span>

## <a name="code-editor"></a><span data-ttu-id="29a0c-125">コード エディター</span><span class="sxs-lookup"><span data-stu-id="29a0c-125">Code Editor</span></span>

![上記のスクリプトのスクリプト コードを表示しているコード エディター。](../images/code-editor-intro.png)

<span data-ttu-id="29a0c-127">操作レコーダーで記録したすべてのスクリプトは、コード エディターで編集できます。</span><span class="sxs-lookup"><span data-stu-id="29a0c-127">All scripts recorded with the Action Recorder can be edited through the Code Editor.</span></span> <span data-ttu-id="29a0c-128">これにより、ニーズにぴったり合うようにスクリプトを微調整したり、カスタマイズしたりできます。</span><span class="sxs-lookup"><span data-stu-id="29a0c-128">This lets you tweak and customize the script to better suit your exact needs.</span></span> <span data-ttu-id="29a0c-129">また、条件付きステートメント (if/else) やループなど、Excel の UI からでは直接アクセスできないロジックや機能を追加することもできます。</span><span class="sxs-lookup"><span data-stu-id="29a0c-129">You can also add logic and functionality that is not directly accessible through the Excel UI, such as conditional statements (if/else) and loops.</span></span>

<span data-ttu-id="29a0c-130">Office スクリプトの機能を学習する簡単な方法の 1 つは、Excel on the web でスクリプトを記録し、作成されたコードを表示することです。</span><span class="sxs-lookup"><span data-stu-id="29a0c-130">One easy way to start learning the capabilities of Office Scripts is to record scripts in Excel on the web and view the resulting code.</span></span> <span data-ttu-id="29a0c-131">別の方法としては、用意されている[チュートリアル](../tutorials/excel-tutorial.md)に従うと、詳しいガイド付きで、より体系的に学習できます。</span><span class="sxs-lookup"><span data-stu-id="29a0c-131">Another option is to follow our [tutorials](../tutorials/excel-tutorial.md) to learn in a more guided and structured way.</span></span>

## <a name="sharing-scripts"></a><span data-ttu-id="29a0c-132">スクリプトの共有</span><span class="sxs-lookup"><span data-stu-id="29a0c-132">Sharing scripts</span></span>

![[このブックで他のユーザーと共有する] オプションを表示するスクリプトの詳細ページ。](../images/script-sharing.png)

<span data-ttu-id="29a0c-134">Office スクリプトは、Excel ブックの他のユーザーと共有できます。</span><span class="sxs-lookup"><span data-stu-id="29a0c-134">Office Scripts can be shared with other users of an Excel workbook.</span></span> <span data-ttu-id="29a0c-135">スクリプトをブック内の他のユーザーと共有すると、スクリプトはブックに添付されます。</span><span class="sxs-lookup"><span data-stu-id="29a0c-135">When you share a script with others in a workbook, the script is attached to the workbook.</span></span> <span data-ttu-id="29a0c-136">スクリプトは、OneDrive に保存され、共有すると、開いているブックにリンクが作成されます。</span><span class="sxs-lookup"><span data-stu-id="29a0c-136">Your scripts are stored in your OneDrive, and when you share one, you create a link to it in the workbook you have open.</span></span>

<span data-ttu-id="29a0c-137">共有および共有解除スクリプトの詳細については、「[Excel for the Web で Office スクリプトを共有する](https://support.microsoft.com/office/sharing-office-scripts-in-excel-for-the-web-226eddbc-3a44-4540-acfe-fccda3d1122b?storagetype=live&ui=en-US&rs=en-US&ad=US)」の記事を参照してください。</span><span class="sxs-lookup"><span data-stu-id="29a0c-137">More details about sharing and unsharing scripts can be in the article [Sharing Office Scripts in Excel for the Web](https://support.microsoft.com/office/sharing-office-scripts-in-excel-for-the-web-226eddbc-3a44-4540-acfe-fccda3d1122b?storagetype=live&ui=en-US&rs=en-US&ad=US).</span></span>

## <a name="connecting-office-scripts-to-power-automate"></a><span data-ttu-id="29a0c-138">Office スクリプトを Power Automate に接続する</span><span class="sxs-lookup"><span data-stu-id="29a0c-138">Connecting Office Scripts to Power Automate</span></span>

<span data-ttu-id="29a0c-139">[Power Automate](https://flow.microsoft.com/) は、複数のアプリとサービスの間のワークフローを自動化するためのサービスです。</span><span class="sxs-lookup"><span data-stu-id="29a0c-139">[Power Automate](https://flow.microsoft.com/) is a service that helps you create automated workflows between multiple apps and services.</span></span> <span data-ttu-id="29a0c-140">これらのワークフローでは、Office スクリプトを使用して、ブック外のスクリプトを制御できます。</span><span class="sxs-lookup"><span data-stu-id="29a0c-140">Office Scripts can be used in these workflows, giving you control of your scripts outside of the workbook.</span></span> <span data-ttu-id="29a0c-141">スケジュールに基づいてスクリプトを実行したり、メールに応じてスクリプトをトリガーしたりできます。</span><span class="sxs-lookup"><span data-stu-id="29a0c-141">You can run your scripts on a schedule, trigger them in response to emails, and much more.</span></span> <span data-ttu-id="29a0c-142">この自動化サービスに接続するための基本的な方法については、「[Power Automate を使用して Excel on the web で Office スクリプトを実行する](../tutorials/excel-power-automate-manual.md)」チュートリアルにアクセスします。</span><span class="sxs-lookup"><span data-stu-id="29a0c-142">Visit the [Run Office Scripts in Excel on the web with Power Automate](../tutorials/excel-power-automate-manual.md) tutorial to learn the basics of connecting these automation services.</span></span>

## <a name="next-steps"></a><span data-ttu-id="29a0c-143">次の手順</span><span class="sxs-lookup"><span data-stu-id="29a0c-143">Next steps</span></span>

<span data-ttu-id="29a0c-144">[Excel on the web の Office スクリプトに関するチュートリアル](../tutorials/excel-tutorial.md)を完了すると、Office スクリプトを初めて作成する方法を理解できます。</span><span class="sxs-lookup"><span data-stu-id="29a0c-144">Complete the [Office Scripts in Excel on the web tutorial](../tutorials/excel-tutorial.md) to learn how to create your first Office Scripts.</span></span>

## <a name="see-also"></a><span data-ttu-id="29a0c-145">関連項目</span><span class="sxs-lookup"><span data-stu-id="29a0c-145">See also</span></span>

- [<span data-ttu-id="29a0c-146">Excel on the web での Office スクリプトのスクリプトの基本事項</span><span class="sxs-lookup"><span data-stu-id="29a0c-146">Scripting fundamentals for Office Scripts in Excel on the web</span></span>](../develop/scripting-fundamentals.md)
- [<span data-ttu-id="29a0c-147">Office スクリプト API リファレンス</span><span class="sxs-lookup"><span data-stu-id="29a0c-147">Office Scripts API reference</span></span>](/javascript/api/office-scripts/overview)
- [<span data-ttu-id="29a0c-148">Office スクリプトのトラブルシューティング</span><span class="sxs-lookup"><span data-stu-id="29a0c-148">Troubleshooting Office Scripts</span></span>](../testing/troubleshooting.md)
- [<span data-ttu-id="29a0c-149">M365 での Office スクリプトの設定</span><span class="sxs-lookup"><span data-stu-id="29a0c-149">Office Scripts settings in M365</span></span>](https://support.office.com/article/office-scripts-settings-in-m365-19d3c51a-6ca2-40ab-978d-60fa49554dcf)
- [<span data-ttu-id="29a0c-150">Excel の Office スクリプトの概要 (support.office.com)</span><span class="sxs-lookup"><span data-stu-id="29a0c-150">Introduction to Office Scripts in Excel (on support.office.com)</span></span>](https://support.office.com/article/introduction-to-office-scripts-in-excel-9fbe283d-adb8-4f13-a75b-a81c6baf163a)
- [<span data-ttu-id="29a0c-151">Excel on the web での Office スクリプトの共有</span><span class="sxs-lookup"><span data-stu-id="29a0c-151">Sharing Office Scripts in Excel for the Web</span></span>](https://support.microsoft.com/office/sharing-office-scripts-in-excel-for-the-web-226eddbc-3a44-4540-acfe-fccda3d1122b?storagetype=live&ui=en-US&rs=en-US&ad=US)
