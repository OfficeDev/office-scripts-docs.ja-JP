---
title: Excel on the web の Office スクリプト
description: Office スクリプト用の操作レコーダーとコード エディターの概要をご紹介します。
ms.date: 07/21/2020
localization_priority: Priority
ms.openlocfilehash: 557a4f8fac8fbeafe150f425170c8e45f62986c6
ms.sourcegitcommit: ff7fde04ce5a66d8df06ed505951c8111e2e9833
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 08/11/2020
ms.locfileid: "46616727"
---
# <a name="office-scripts-in-excel-on-the-web-preview"></a><span data-ttu-id="02a01-103">Excel on the web の Office スクリプト (プレビュー)</span><span class="sxs-lookup"><span data-stu-id="02a01-103">Office Scripts in Excel on the web (preview)</span></span>

<span data-ttu-id="02a01-104">Excel on the web の Office スクリプトを使用すると、日常のタスクを自動化できます。</span><span class="sxs-lookup"><span data-stu-id="02a01-104">Office Scripts in Excel on the web let you automate your day-to-day tasks.</span></span> <span data-ttu-id="02a01-105">Excel で行う操作を操作レコーダーで記録すると、スクリプトが作成されます。</span><span class="sxs-lookup"><span data-stu-id="02a01-105">You can record your Excel actions with the Action Recorder, which creates a script.</span></span> <span data-ttu-id="02a01-106">さらに、コード エディターでスクリプトの作成や編集をすることもできます。</span><span class="sxs-lookup"><span data-stu-id="02a01-106">You can also create and edit scripts with the Code Editor.</span></span> <span data-ttu-id="02a01-107">スクリプトは組織全体で共有できるため、同僚もワークフローを自動化できます。</span><span class="sxs-lookup"><span data-stu-id="02a01-107">Your scripts can then be shared across your organization so your coworkers can also automate their workflows.</span></span>

<span data-ttu-id="02a01-108">この一連のドキュメントで、これらのツールの使用方法について説明します。</span><span class="sxs-lookup"><span data-stu-id="02a01-108">This series of documents teaches you how to use these tools.</span></span> <span data-ttu-id="02a01-109">操作レコーダーの紹介では、頻繁に実行する Excel 操作の記録方法を説明します。</span><span class="sxs-lookup"><span data-stu-id="02a01-109">You'll be introduced to the Action Recorder and see how to record your frequent Excel actions.</span></span> <span data-ttu-id="02a01-110">また、コード エディターを使用して、独自のスクリプトを作成したり更新したりする方法についても説明します。</span><span class="sxs-lookup"><span data-stu-id="02a01-110">You'll also learn how to make or update your own scripts with the Code Editor.</span></span>

<br>

> [!VIDEO https://www.microsoft.com/videoplayer/embed/RE4qdFF]

## <a name="requirements"></a><span data-ttu-id="02a01-111">要件</span><span class="sxs-lookup"><span data-stu-id="02a01-111">Requirements</span></span>

[!INCLUDE [Preview note](../includes/preview-note.md)]

<span data-ttu-id="02a01-112">Office スクリプトを使用するには、以下が必要です。</span><span class="sxs-lookup"><span data-stu-id="02a01-112">To use Office Scripts, you'll need the following.</span></span>

1. <span data-ttu-id="02a01-113">[Excel on the web](https://www.office.com/launch/excel) (デスクトップなどのその他のプラットフォームは、サポートされていません)。</span><span class="sxs-lookup"><span data-stu-id="02a01-113">[Excel on the web](https://www.office.com/launch/excel) (other platforms, such as desktop, are not supported).</span></span>
1. <span data-ttu-id="02a01-114">[管理者によって有効にされた](https://support.office.com/article/office-scripts-settings-in-m365-19d3c51a-6ca2-40ab-978d-60fa49554dcf) Office スクリプト。</span><span class="sxs-lookup"><span data-stu-id="02a01-114">Office Scripts [enabled by your administrator](https://support.office.com/article/office-scripts-settings-in-m365-19d3c51a-6ca2-40ab-978d-60fa49554dcf).</span></span>
1. <span data-ttu-id="02a01-115">Microsoft 365 Office デスクトップ アプリにアクセスできる、次のような商用または教育機関向けの Microsoft 365 ライセンス。</span><span class="sxs-lookup"><span data-stu-id="02a01-115">Any commercial or educational Microsoft 365 license with access to the Microsoft 365 Office desktop apps, such as:</span></span>

    - <span data-ttu-id="02a01-116">Office 365 Business</span><span class="sxs-lookup"><span data-stu-id="02a01-116">Office 365 Business</span></span>
    - <span data-ttu-id="02a01-117">Office 365 Business Premium</span><span class="sxs-lookup"><span data-stu-id="02a01-117">Office 365 Business Premium</span></span>
    - <span data-ttu-id="02a01-118">Office 365 ProPlus</span><span class="sxs-lookup"><span data-stu-id="02a01-118">Office 365 ProPlus</span></span>
    - <span data-ttu-id="02a01-119">Office 365 ProPlus デバイス用</span><span class="sxs-lookup"><span data-stu-id="02a01-119">Office 365 ProPlus for Devices</span></span>
    - <span data-ttu-id="02a01-120">Office 365 Enterprise E3</span><span class="sxs-lookup"><span data-stu-id="02a01-120">Office 365 Enterprise E3</span></span>
    - <span data-ttu-id="02a01-121">Office 365 Enterprise E5</span><span class="sxs-lookup"><span data-stu-id="02a01-121">Office 365 Enterprise E5</span></span>
    - <span data-ttu-id="02a01-122">Office 365 A3</span><span class="sxs-lookup"><span data-stu-id="02a01-122">Office 365 A3</span></span>
    - <span data-ttu-id="02a01-123">Office 365 A5</span><span class="sxs-lookup"><span data-stu-id="02a01-123">Office 365 A5</span></span>

## <a name="when-to-use-office-scripts"></a><span data-ttu-id="02a01-124">Office スクリプトの使用に適した状況</span><span class="sxs-lookup"><span data-stu-id="02a01-124">When to use Office Scripts</span></span>

<span data-ttu-id="02a01-125">スクリプトを使用すると、自分が行った Excel の操作を記録して、さまざまなブックやワークシートに対してその操作を再現できます。</span><span class="sxs-lookup"><span data-stu-id="02a01-125">Scripts allow you to record and replay your Excel actions on different workbooks and worksheets.</span></span> <span data-ttu-id="02a01-126">同じ操作を何度も繰り返し行っている場合は、そのすべての作業を簡単に実行できる Office スクリプトに変換することができます。</span><span class="sxs-lookup"><span data-stu-id="02a01-126">If you find yourself doing the same things over and over again, you can turn all that work into an easy-to-run Office Script.</span></span> <span data-ttu-id="02a01-127">Excel でボタンを押してスクリプトを実行するか、Power Automate と組み合わせてワークフロー全体を効率化します。</span><span class="sxs-lookup"><span data-stu-id="02a01-127">Run your script with a button-press in Excel or combine it with Power Automate to streamline your entire workflow.</span></span>

<span data-ttu-id="02a01-128">たとえば、毎日仕事の始めに Excel で会計サイトから .csv ファイルを開いているとします。</span><span class="sxs-lookup"><span data-stu-id="02a01-128">As an example, say you start your work day by opening a .csv file from an accounting site in Excel.</span></span> <span data-ttu-id="02a01-129">それから数分かけて、不要な列を削除し、テーブルの書式を設定し、数式を追加し、新しいワークシートにピボットテーブルを作成します。</span><span class="sxs-lookup"><span data-stu-id="02a01-129">You then spend several minutes deleting unnecessary columns, formatting a table, adding formulas, and creating a PivotTable in a new worksheet.</span></span> <span data-ttu-id="02a01-130">毎日繰り返しているこのような操作を、操作レコーダーで 1 回記録できます。</span><span class="sxs-lookup"><span data-stu-id="02a01-130">Those actions you repeat daily can be recorded once with the Action Recorder.</span></span> <span data-ttu-id="02a01-131">それ以降は、スクリプトを実行するだけで、.csv の変換処理すべてが自動的に実行されます。</span><span class="sxs-lookup"><span data-stu-id="02a01-131">From then on, running the script will take care of your entire .csv conversion.</span></span> <span data-ttu-id="02a01-132">手順を忘れる危険がなくなるだけでなく、特に操作を教えなくても他の人とプロセスを共有することもできます。</span><span class="sxs-lookup"><span data-stu-id="02a01-132">You'll not only remove the risk of forgetting steps, but be able to share your process with others without having to teach them anything.</span></span> <span data-ttu-id="02a01-133">Office スクリプトを使用すると一般的なタスクを自動化できるので、自分自身と職場の作業効率や生産性を向上できます。</span><span class="sxs-lookup"><span data-stu-id="02a01-133">Office Scripts automate your common tasks so you and your workplace can be more efficient and productive.</span></span>

## <a name="action-recorder"></a><span data-ttu-id="02a01-134">操作レコーダー</span><span class="sxs-lookup"><span data-stu-id="02a01-134">Action Recorder</span></span>

![いくつかの操作を記録した後の操作レコーダー。](../images/action-recorder-intro.png)

<span data-ttu-id="02a01-136">操作レコーダーは、ユーザーが Excel で実行した操作を記録して、スクリプトとして保存します。</span><span class="sxs-lookup"><span data-stu-id="02a01-136">The Action Recorder records actions you take in Excel and saves them as a script.</span></span> <span data-ttu-id="02a01-137">操作レコーダーを実行すると、セルの編集、書式の変更、テーブルの作成などの Excel の操作をキャプチャできます。</span><span class="sxs-lookup"><span data-stu-id="02a01-137">With the Action recorder running, you can capture the Excel actions as you edit cells, change formatting, and create tables.</span></span> <span data-ttu-id="02a01-138">作成されたスクリプトは、他のワークシートやブックで実行して、ユーザーが実行した元の操作を再現することもできます。</span><span class="sxs-lookup"><span data-stu-id="02a01-138">The resulting script can be run on other worksheets and workbooks to recreate your original actions.</span></span>

## <a name="code-editor"></a><span data-ttu-id="02a01-139">コード エディター</span><span class="sxs-lookup"><span data-stu-id="02a01-139">Code Editor</span></span>

![上記のスクリプトのスクリプト コードを表示しているコード エディター。](../images/code-editor-intro.png)

<span data-ttu-id="02a01-141">操作レコーダーで記録したすべてのスクリプトは、コード エディターで編集できます。</span><span class="sxs-lookup"><span data-stu-id="02a01-141">All scripts recorded with the Action Recorder can be edited through the Code Editor.</span></span> <span data-ttu-id="02a01-142">これにより、ニーズにぴったり合うようにスクリプトを微調整したり、カスタマイズしたりできます。</span><span class="sxs-lookup"><span data-stu-id="02a01-142">This lets you tweak and customize the script to better suit your exact needs.</span></span> <span data-ttu-id="02a01-143">また、条件付きステートメント (if/else) やループなど、Excel の UI からでは直接アクセスできないロジックや機能を追加することもできます。</span><span class="sxs-lookup"><span data-stu-id="02a01-143">You can also add logic and functionality that is not directly accessible through the Excel UI, such as conditional statements (if/else) and loops.</span></span>

<span data-ttu-id="02a01-144">Office スクリプトの機能を学習する簡単な方法の 1 つは、Excel on the web でスクリプトを記録し、作成されたコードを表示することです。</span><span class="sxs-lookup"><span data-stu-id="02a01-144">One easy way to start learning the capabilities of Office Scripts is to record scripts in Excel on the web and view the resulting code.</span></span> <span data-ttu-id="02a01-145">別の方法としては、用意されている[チュートリアル](../tutorials/excel-tutorial.md)に従うと、詳しいガイド付きで、より体系的に学習できます。</span><span class="sxs-lookup"><span data-stu-id="02a01-145">Another option is to follow our [tutorials](../tutorials/excel-tutorial.md) to learn in a more guided and structured way.</span></span>

## <a name="sharing-scripts"></a><span data-ttu-id="02a01-146">スクリプトの共有</span><span class="sxs-lookup"><span data-stu-id="02a01-146">Sharing scripts</span></span>

![[このブックで他のユーザーと共有する] オプションを表示するスクリプトの詳細ページ。](../images/script-sharing.png)

<span data-ttu-id="02a01-148">Office スクリプトは、Excel ブックの他のユーザーと共有できます。</span><span class="sxs-lookup"><span data-stu-id="02a01-148">Office Scripts can be shared with other users of an Excel workbook.</span></span> <span data-ttu-id="02a01-149">スクリプトをブック内の他のユーザーと共有すると、スクリプトはブックに添付されます。</span><span class="sxs-lookup"><span data-stu-id="02a01-149">When you share a script with others in a workbook, the script is attached to the workbook.</span></span> <span data-ttu-id="02a01-150">スクリプトは、OneDrive に保存され、共有すると、開いているブックにリンクが作成されます。</span><span class="sxs-lookup"><span data-stu-id="02a01-150">Your scripts are stored in your OneDrive, and when you share one, you create a link to it in the workbook you have open.</span></span>

<span data-ttu-id="02a01-151">共有および共有解除スクリプトの詳細については、「[Excel for the Web で Office スクリプトを共有する](https://support.microsoft.com/office/sharing-office-scripts-in-excel-for-the-web-226eddbc-3a44-4540-acfe-fccda3d1122b?storagetype=live&ui=en-US&rs=en-US&ad=US)」の記事を参照してください。</span><span class="sxs-lookup"><span data-stu-id="02a01-151">More details about sharing and unsharing scripts can be in the article [Sharing Office Scripts in Excel for the Web](https://support.microsoft.com/office/sharing-office-scripts-in-excel-for-the-web-226eddbc-3a44-4540-acfe-fccda3d1122b?storagetype=live&ui=en-US&rs=en-US&ad=US).</span></span>

## <a name="connecting-office-scripts-to-power-automate"></a><span data-ttu-id="02a01-152">Office スクリプトを Power Automate に接続する</span><span class="sxs-lookup"><span data-stu-id="02a01-152">Connecting Office Scripts to Power Automate</span></span>

<span data-ttu-id="02a01-153">[Power Automate](https://flow.microsoft.com/) は、複数のアプリとサービスの間のワークフローを自動化するためのサービスです。</span><span class="sxs-lookup"><span data-stu-id="02a01-153">[Power Automate](https://flow.microsoft.com/) is a service that helps you create automated workflows between multiple apps and services.</span></span> <span data-ttu-id="02a01-154">これらのワークフローでは、Office スクリプトを使用して、ブック外のスクリプトを制御できます。</span><span class="sxs-lookup"><span data-stu-id="02a01-154">Office Scripts can be used in these workflows, giving you control of your scripts outside of the workbook.</span></span> <span data-ttu-id="02a01-155">スケジュールに基づいてスクリプトを実行したり、メールに応じてスクリプトをトリガーしたりできます。</span><span class="sxs-lookup"><span data-stu-id="02a01-155">You can run your scripts on a schedule, trigger them in response to emails, and much more.</span></span> <span data-ttu-id="02a01-156">この自動化サービスに接続するための基本的な方法については、「[Power Automate を使用して Excel on the web で Office スクリプトを実行する](../tutorials/excel-power-automate-manual.md)」チュートリアルにアクセスします。</span><span class="sxs-lookup"><span data-stu-id="02a01-156">Visit the [Run Office Scripts in Excel on the web with Power Automate](../tutorials/excel-power-automate-manual.md) tutorial to learn the basics of connecting these automation services.</span></span>

## <a name="next-steps"></a><span data-ttu-id="02a01-157">次の手順</span><span class="sxs-lookup"><span data-stu-id="02a01-157">Next steps</span></span>

<span data-ttu-id="02a01-158">[Excel on the web の Office スクリプトに関するチュートリアル](../tutorials/excel-tutorial.md)を完了すると、Office スクリプトを初めて作成する方法を理解できます。</span><span class="sxs-lookup"><span data-stu-id="02a01-158">Complete the [Office Scripts in Excel on the web tutorial](../tutorials/excel-tutorial.md) to learn how to create your first Office Scripts.</span></span>

## <a name="see-also"></a><span data-ttu-id="02a01-159">関連項目</span><span class="sxs-lookup"><span data-stu-id="02a01-159">See also</span></span>

- [<span data-ttu-id="02a01-160">Excel on the web での Office スクリプトのスクリプトの基本事項</span><span class="sxs-lookup"><span data-stu-id="02a01-160">Scripting fundamentals for Office Scripts in Excel on the web</span></span>](../develop/scripting-fundamentals.md)
- [<span data-ttu-id="02a01-161">Office スクリプト API リファレンス</span><span class="sxs-lookup"><span data-stu-id="02a01-161">Office Scripts API reference</span></span>](/javascript/api/office-scripts/overview)
- [<span data-ttu-id="02a01-162">Office スクリプトのトラブルシューティング</span><span class="sxs-lookup"><span data-stu-id="02a01-162">Troubleshooting Office Scripts</span></span>](../testing/troubleshooting.md)
- [<span data-ttu-id="02a01-163">M365 での Office スクリプトの設定</span><span class="sxs-lookup"><span data-stu-id="02a01-163">Office Scripts settings in M365</span></span>](https://support.office.com/article/office-scripts-settings-in-m365-19d3c51a-6ca2-40ab-978d-60fa49554dcf)
- [<span data-ttu-id="02a01-164">Excel の Office スクリプトの概要 (support.office.com)</span><span class="sxs-lookup"><span data-stu-id="02a01-164">Introduction to Office Scripts in Excel (on support.office.com)</span></span>](https://support.office.com/article/introduction-to-office-scripts-in-excel-9fbe283d-adb8-4f13-a75b-a81c6baf163a)
- [<span data-ttu-id="02a01-165">Excel on the web での Office スクリプトの共有</span><span class="sxs-lookup"><span data-stu-id="02a01-165">Sharing Office Scripts in Excel for the Web</span></span>](https://support.microsoft.com/office/sharing-office-scripts-in-excel-for-the-web-226eddbc-3a44-4540-acfe-fccda3d1122b?storagetype=live&ui=en-US&rs=en-US&ad=US)
