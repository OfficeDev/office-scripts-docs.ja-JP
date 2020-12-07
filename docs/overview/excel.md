---
title: Excel on the web の Office スクリプト
description: Office スクリプト用の操作レコーダーとコード エディターの概要をご紹介します。
ms.date: 11/13/2020
localization_priority: Priority
ms.openlocfilehash: a065c8eb5fc52c7525383927b7e1490e703eb179
ms.sourcegitcommit: 82d3c0ef1e187bcdeceb2b5fc3411186674fe150
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 11/18/2020
ms.locfileid: "49571463"
---
# <a name="office-scripts-in-excel-on-the-web-preview"></a><span data-ttu-id="f929c-103">Excel on the web の Office スクリプト (プレビュー)</span><span class="sxs-lookup"><span data-stu-id="f929c-103">Office Scripts in Excel on the web (preview)</span></span>

<span data-ttu-id="f929c-104">Excel on the web の Office スクリプトを使用すると、日常のタスクを自動化できます。</span><span class="sxs-lookup"><span data-stu-id="f929c-104">Office Scripts in Excel on the web let you automate your day-to-day tasks.</span></span> <span data-ttu-id="f929c-105">Excel で行う操作を操作レコーダーで記録すると、スクリプトが作成されます。</span><span class="sxs-lookup"><span data-stu-id="f929c-105">You can record your Excel actions with the Action Recorder, which creates a script.</span></span> <span data-ttu-id="f929c-106">さらに、コード エディターでスクリプトの作成や編集をすることもできます。</span><span class="sxs-lookup"><span data-stu-id="f929c-106">You can also create and edit scripts with the Code Editor.</span></span> <span data-ttu-id="f929c-107">スクリプトは組織全体で共有できるため、同僚もワークフローを自動化できます。</span><span class="sxs-lookup"><span data-stu-id="f929c-107">Your scripts can then be shared across your organization so your coworkers can also automate their workflows.</span></span>

<span data-ttu-id="f929c-108">この一連のドキュメントで、これらのツールの使用方法について説明します。</span><span class="sxs-lookup"><span data-stu-id="f929c-108">This series of documents teaches you how to use these tools.</span></span> <span data-ttu-id="f929c-109">操作レコーダーの紹介では、頻繁に実行する Excel 操作の記録方法を説明します。</span><span class="sxs-lookup"><span data-stu-id="f929c-109">You'll be introduced to the Action Recorder and see how to record your frequent Excel actions.</span></span> <span data-ttu-id="f929c-110">また、コード エディターを使用して、独自のスクリプトを作成したり更新したりする方法についても説明します。</span><span class="sxs-lookup"><span data-stu-id="f929c-110">You'll also learn how to make or update your own scripts with the Code Editor.</span></span>

<br>

> [!VIDEO https://www.microsoft.com/videoplayer/embed/RE4qdFF]

## <a name="requirements"></a><span data-ttu-id="f929c-111">要件</span><span class="sxs-lookup"><span data-stu-id="f929c-111">Requirements</span></span>

[!INCLUDE [Preview note](../includes/preview-note.md)]

<span data-ttu-id="f929c-112">Office スクリプトを使用するには、以下が必要です。</span><span class="sxs-lookup"><span data-stu-id="f929c-112">To use Office Scripts, you'll need the following.</span></span>

1. <span data-ttu-id="f929c-113">[Excel on the web](https://www.office.com/launch/excel) (デスクトップなどのその他のプラットフォームは、サポートされていません)。</span><span class="sxs-lookup"><span data-stu-id="f929c-113">[Excel on the web](https://www.office.com/launch/excel) (other platforms, such as desktop, are not supported).</span></span>
1. <span data-ttu-id="f929c-114">OneDrive for Business。</span><span class="sxs-lookup"><span data-stu-id="f929c-114">OneDrive for Business.</span></span>
1. <span data-ttu-id="f929c-115">Microsoft 365 Office デスクトップ アプリにアクセスできる、次のような商用または教育機関向けの Microsoft 365 ライセンス。</span><span class="sxs-lookup"><span data-stu-id="f929c-115">Any commercial or educational Microsoft 365 license with access to the Microsoft 365 Office desktop apps, such as:</span></span>

    - <span data-ttu-id="f929c-116">Office 365 Business</span><span class="sxs-lookup"><span data-stu-id="f929c-116">Office 365 Business</span></span>
    - <span data-ttu-id="f929c-117">Office 365 Business Premium</span><span class="sxs-lookup"><span data-stu-id="f929c-117">Office 365 Business Premium</span></span>
    - <span data-ttu-id="f929c-118">Office 365 ProPlus</span><span class="sxs-lookup"><span data-stu-id="f929c-118">Office 365 ProPlus</span></span>
    - <span data-ttu-id="f929c-119">Office 365 ProPlus デバイス用</span><span class="sxs-lookup"><span data-stu-id="f929c-119">Office 365 ProPlus for Devices</span></span>
    - <span data-ttu-id="f929c-120">Office 365 Enterprise E3</span><span class="sxs-lookup"><span data-stu-id="f929c-120">Office 365 Enterprise E3</span></span>
    - <span data-ttu-id="f929c-121">Office 365 Enterprise E5</span><span class="sxs-lookup"><span data-stu-id="f929c-121">Office 365 Enterprise E5</span></span>
    - <span data-ttu-id="f929c-122">Office 365 A3</span><span class="sxs-lookup"><span data-stu-id="f929c-122">Office 365 A3</span></span>
    - <span data-ttu-id="f929c-123">Office 365 A5</span><span class="sxs-lookup"><span data-stu-id="f929c-123">Office 365 A5</span></span>

> [!NOTE]
> <span data-ttu-id="f929c-124">これらの条件を満たしているにもかかわらず **[自動化]** タブが表示されない場合は、管理者が機能を無効にしているか、ご利用の環境に何らかの問題がある可能性があります。</span><span class="sxs-lookup"><span data-stu-id="f929c-124">If you meet these requirements and are still not seeing the **Automate** tab, it's possible that your admin has disabled the feature or there's some other problem with your environment.</span></span> <span data-ttu-id="f929c-125">「[Automate tab not appearing or Office Scripts unavailable (自動化タブが表示されない、または Office スクリプトを使用できない)](../testing/troubleshooting.md#automate-tab-not-appearing-or-office-scripts-unavailable)」の手順に従い、Office スクリプトの使用を開始してください。</span><span class="sxs-lookup"><span data-stu-id="f929c-125">Please follow the steps under [Automate tab not appearing or Office Scripts unavailable](../testing/troubleshooting.md#automate-tab-not-appearing-or-office-scripts-unavailable) to start using Office Scripts.</span></span>

## <a name="when-to-use-office-scripts"></a><span data-ttu-id="f929c-126">Office スクリプトの使用に適した状況</span><span class="sxs-lookup"><span data-stu-id="f929c-126">When to use Office Scripts</span></span>

<span data-ttu-id="f929c-127">スクリプトを使用すると、自分が行った Excel の操作を記録して、さまざまなブックやワークシートに対してその操作を再現できます。</span><span class="sxs-lookup"><span data-stu-id="f929c-127">Scripts allow you to record and replay your Excel actions on different workbooks and worksheets.</span></span> <span data-ttu-id="f929c-128">同じ操作を何度も繰り返し行っている場合は、そのすべての作業を簡単に実行できる Office スクリプトに変換することができます。</span><span class="sxs-lookup"><span data-stu-id="f929c-128">If you find yourself doing the same things over and over again, you can turn all that work into an easy-to-run Office Script.</span></span> <span data-ttu-id="f929c-129">Excel でボタンを押してスクリプトを実行するか、Power Automate と組み合わせてワークフロー全体を効率化します。</span><span class="sxs-lookup"><span data-stu-id="f929c-129">Run your script with a button-press in Excel or combine it with Power Automate to streamline your entire workflow.</span></span>

<span data-ttu-id="f929c-130">たとえば、毎日仕事の始めに Excel で会計サイトから .csv ファイルを開いているとします。</span><span class="sxs-lookup"><span data-stu-id="f929c-130">As an example, say you start your work day by opening a .csv file from an accounting site in Excel.</span></span> <span data-ttu-id="f929c-131">それから数分かけて、不要な列を削除し、テーブルの書式を設定し、数式を追加し、新しいワークシートにピボットテーブルを作成します。</span><span class="sxs-lookup"><span data-stu-id="f929c-131">You then spend several minutes deleting unnecessary columns, formatting a table, adding formulas, and creating a PivotTable in a new worksheet.</span></span> <span data-ttu-id="f929c-132">毎日繰り返しているこのような操作を、操作レコーダーで 1 回記録できます。</span><span class="sxs-lookup"><span data-stu-id="f929c-132">Those actions you repeat daily can be recorded once with the Action Recorder.</span></span> <span data-ttu-id="f929c-133">それ以降は、スクリプトを実行するだけで、.csv の変換処理すべてが自動的に実行されます。</span><span class="sxs-lookup"><span data-stu-id="f929c-133">From then on, running the script will take care of your entire .csv conversion.</span></span> <span data-ttu-id="f929c-134">手順を忘れる危険がなくなるだけでなく、特に操作を教えなくても他の人とプロセスを共有することもできます。</span><span class="sxs-lookup"><span data-stu-id="f929c-134">You'll not only remove the risk of forgetting steps, but be able to share your process with others without having to teach them anything.</span></span> <span data-ttu-id="f929c-135">Office スクリプトを使用すると一般的なタスクを自動化できるので、自分自身と職場の作業効率や生産性を向上できます。</span><span class="sxs-lookup"><span data-stu-id="f929c-135">Office Scripts automate your common tasks so you and your workplace can be more efficient and productive.</span></span>

## <a name="action-recorder"></a><span data-ttu-id="f929c-136">操作レコーダー</span><span class="sxs-lookup"><span data-stu-id="f929c-136">Action Recorder</span></span>

![いくつかの操作を記録した後の操作レコーダー。](../images/action-recorder-intro.png)

<span data-ttu-id="f929c-138">操作レコーダーは、ユーザーが Excel で実行した操作を記録して、スクリプトとして保存します。</span><span class="sxs-lookup"><span data-stu-id="f929c-138">The Action Recorder records actions you take in Excel and saves them as a script.</span></span> <span data-ttu-id="f929c-139">操作レコーダーを実行すると、セルの編集、書式の変更、テーブルの作成などの Excel の操作をキャプチャできます。</span><span class="sxs-lookup"><span data-stu-id="f929c-139">With the Action recorder running, you can capture the Excel actions as you edit cells, change formatting, and create tables.</span></span> <span data-ttu-id="f929c-140">作成されたスクリプトは、他のワークシートやブックで実行して、ユーザーが実行した元の操作を再現することもできます。</span><span class="sxs-lookup"><span data-stu-id="f929c-140">The resulting script can be run on other worksheets and workbooks to recreate your original actions.</span></span>

## <a name="code-editor"></a><span data-ttu-id="f929c-141">コード エディター</span><span class="sxs-lookup"><span data-stu-id="f929c-141">Code Editor</span></span>

![上記のスクリプトのスクリプト コードを表示しているコード エディター。](../images/code-editor-intro.png)

<span data-ttu-id="f929c-143">操作レコーダーで記録したすべてのスクリプトは、コード エディターで編集できます。</span><span class="sxs-lookup"><span data-stu-id="f929c-143">All scripts recorded with the Action Recorder can be edited through the Code Editor.</span></span> <span data-ttu-id="f929c-144">これにより、ニーズにぴったり合うようにスクリプトを微調整したり、カスタマイズしたりできます。</span><span class="sxs-lookup"><span data-stu-id="f929c-144">This lets you tweak and customize the script to better suit your exact needs.</span></span> <span data-ttu-id="f929c-145">また、条件付きステートメント (if/else) やループなど、Excel の UI からでは直接アクセスできないロジックや機能を追加することもできます。</span><span class="sxs-lookup"><span data-stu-id="f929c-145">You can also add logic and functionality that is not directly accessible through the Excel UI, such as conditional statements (if/else) and loops.</span></span>

<span data-ttu-id="f929c-146">Office スクリプトの機能を学習する簡単な方法の 1 つは、Excel on the web でスクリプトを記録し、作成されたコードを表示することです。</span><span class="sxs-lookup"><span data-stu-id="f929c-146">One easy way to start learning the capabilities of Office Scripts is to record scripts in Excel on the web and view the resulting code.</span></span> <span data-ttu-id="f929c-147">別の方法としては、用意されている[チュートリアル](../tutorials/excel-tutorial.md)に従うと、詳しいガイド付きで、より体系的に学習できます。</span><span class="sxs-lookup"><span data-stu-id="f929c-147">Another option is to follow our [tutorials](../tutorials/excel-tutorial.md) to learn in a more guided and structured way.</span></span>

## <a name="sharing-scripts"></a><span data-ttu-id="f929c-148">スクリプトの共有</span><span class="sxs-lookup"><span data-stu-id="f929c-148">Sharing scripts</span></span>

![[このブックで他のユーザーと共有する] オプションを表示するスクリプトの詳細ページ。](../images/script-sharing.png)

<span data-ttu-id="f929c-150">Office スクリプトは、Excel ブックの他のユーザーと共有できます。</span><span class="sxs-lookup"><span data-stu-id="f929c-150">Office Scripts can be shared with other users of an Excel workbook.</span></span> <span data-ttu-id="f929c-151">スクリプトをブック内の他のユーザーと共有すると、スクリプトはブックに添付されます。</span><span class="sxs-lookup"><span data-stu-id="f929c-151">When you share a script with others in a workbook, the script is attached to the workbook.</span></span> <span data-ttu-id="f929c-152">スクリプトは、OneDrive に保存され、共有すると、開いているブックにリンクが作成されます。</span><span class="sxs-lookup"><span data-stu-id="f929c-152">Your scripts are stored in your OneDrive, and when you share one, you create a link to it in the workbook you have open.</span></span>

<span data-ttu-id="f929c-153">共有および共有解除スクリプトの詳細については、「[Excel for the Web で Office スクリプトを共有する](https://support.microsoft.com/office/sharing-office-scripts-in-excel-for-the-web-226eddbc-3a44-4540-acfe-fccda3d1122b)」の記事を参照してください。</span><span class="sxs-lookup"><span data-stu-id="f929c-153">More details about sharing and unsharing scripts can be in the article [Sharing Office Scripts in Excel for the Web](https://support.microsoft.com/office/sharing-office-scripts-in-excel-for-the-web-226eddbc-3a44-4540-acfe-fccda3d1122b).</span></span>

> [!NOTE]
> <span data-ttu-id="f929c-154">「[Office Scripts file storage and ownership (Office スクリプトのファイル ストレージと所有権)](script-storage.md)」では、OneDrive にスクリプトを保存する方法について詳しく説明しています。</span><span class="sxs-lookup"><span data-stu-id="f929c-154">Learn more about how scripts are stored in your OneDrive in [Office Scripts file storage and ownership](script-storage.md).</span></span>

## <a name="connecting-office-scripts-to-power-automate"></a><span data-ttu-id="f929c-155">Office スクリプトを Power Automate に接続する</span><span class="sxs-lookup"><span data-stu-id="f929c-155">Connecting Office Scripts to Power Automate</span></span>

<span data-ttu-id="f929c-156">[Power Automate](https://flow.microsoft.com/) は、複数のアプリとサービスの間のワークフローを自動化するためのサービスです。</span><span class="sxs-lookup"><span data-stu-id="f929c-156">[Power Automate](https://flow.microsoft.com/) is a service that helps you create automated workflows between multiple apps and services.</span></span> <span data-ttu-id="f929c-157">これらのワークフローでは、Office スクリプトを使用して、ブック外のスクリプトを制御できます。</span><span class="sxs-lookup"><span data-stu-id="f929c-157">Office Scripts can be used in these workflows, giving you control of your scripts outside of the workbook.</span></span> <span data-ttu-id="f929c-158">スケジュールに基づいてスクリプトを実行したり、メールに応じてスクリプトをトリガーしたりできます。</span><span class="sxs-lookup"><span data-stu-id="f929c-158">You can run your scripts on a schedule, trigger them in response to emails, and much more.</span></span> <span data-ttu-id="f929c-159">この自動化サービスに接続するための基本的な方法については、「[Power Automate を使用して Excel on the web で Office スクリプトを実行する](../tutorials/excel-power-automate-manual.md)」チュートリアルにアクセスします。</span><span class="sxs-lookup"><span data-stu-id="f929c-159">Visit the [Run Office Scripts in Excel on the web with Power Automate](../tutorials/excel-power-automate-manual.md) tutorial to learn the basics of connecting these automation services.</span></span>

## <a name="next-steps"></a><span data-ttu-id="f929c-160">次の手順</span><span class="sxs-lookup"><span data-stu-id="f929c-160">Next steps</span></span>

<span data-ttu-id="f929c-161">[Excel on the web の Office スクリプトに関するチュートリアル](../tutorials/excel-tutorial.md)を完了すると、Office スクリプトを初めて作成する方法を理解できます。</span><span class="sxs-lookup"><span data-stu-id="f929c-161">Complete the [Office Scripts in Excel on the web tutorial](../tutorials/excel-tutorial.md) to learn how to create your first Office Scripts.</span></span>

## <a name="see-also"></a><span data-ttu-id="f929c-162">関連項目</span><span class="sxs-lookup"><span data-stu-id="f929c-162">See also</span></span>

- [<span data-ttu-id="f929c-163">Excel on the web での Office スクリプトのスクリプトの基本事項</span><span class="sxs-lookup"><span data-stu-id="f929c-163">Scripting fundamentals for Office Scripts in Excel on the web</span></span>](../develop/scripting-fundamentals.md)
- [<span data-ttu-id="f929c-164">Office スクリプト API リファレンス</span><span class="sxs-lookup"><span data-stu-id="f929c-164">Office Scripts API reference</span></span>](/javascript/api/office-scripts/overview)
- [<span data-ttu-id="f929c-165">Office スクリプトのトラブルシューティング</span><span class="sxs-lookup"><span data-stu-id="f929c-165">Troubleshooting Office Scripts</span></span>](../testing/troubleshooting.md)
- [<span data-ttu-id="f929c-166">M365 での Office スクリプトの設定</span><span class="sxs-lookup"><span data-stu-id="f929c-166">Office Scripts settings in M365</span></span>](https://support.office.com/article/office-scripts-settings-in-m365-19d3c51a-6ca2-40ab-978d-60fa49554dcf)
- [<span data-ttu-id="f929c-167">Excel の Office スクリプトの概要 (support.office.com)</span><span class="sxs-lookup"><span data-stu-id="f929c-167">Introduction to Office Scripts in Excel (on support.office.com)</span></span>](https://support.office.com/article/introduction-to-office-scripts-in-excel-9fbe283d-adb8-4f13-a75b-a81c6baf163a)
- [<span data-ttu-id="f929c-168">Excel on the web での Office スクリプトの共有</span><span class="sxs-lookup"><span data-stu-id="f929c-168">Sharing Office Scripts in Excel for the Web</span></span>](https://support.microsoft.com/office/sharing-office-scripts-in-excel-for-the-web-226eddbc-3a44-4540-acfe-fccda3d1122b)
