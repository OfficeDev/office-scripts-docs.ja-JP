---
title: Excel on the web の Office スクリプト
description: Office スクリプト用の操作レコーダーとコード エディターの概要をご紹介します。
ms.date: 02/24/2020
localization_priority: Priority
ms.openlocfilehash: dd48467bc8105a3d31d9fa21e547c703e9e37cce
ms.sourcegitcommit: b075eed5a6f275274fbbf6d62633219eac416f26
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 03/10/2020
ms.locfileid: "42700431"
---
# <a name="office-scripts-in-excel-on-the-web-preview"></a><span data-ttu-id="62d3b-103">Excel on the web の Office スクリプト (プレビュー)</span><span class="sxs-lookup"><span data-stu-id="62d3b-103">Office Scripts in Excel on the web (preview)</span></span>

<span data-ttu-id="62d3b-104">Excel on the web の Office スクリプトを使用すると、日常のタスクを自動化できます。</span><span class="sxs-lookup"><span data-stu-id="62d3b-104">Office Scripts in Excel on the web let you automate your day-to-day tasks.</span></span> <span data-ttu-id="62d3b-105">Excel で行う操作を操作レコーダーで記録すると、スクリプトが作成されます。</span><span class="sxs-lookup"><span data-stu-id="62d3b-105">You can record your Excel actions with the Action Recorder, which creates a script.</span></span> <span data-ttu-id="62d3b-106">さらに、コード エディターでスクリプトの作成や編集をすることもできます。</span><span class="sxs-lookup"><span data-stu-id="62d3b-106">You can also create and edit scripts with the Code Editor.</span></span> <span data-ttu-id="62d3b-107">この一連のドキュメントで、これらのツールの使用方法について説明します。</span><span class="sxs-lookup"><span data-stu-id="62d3b-107">This series of documents teaches you how to use these tools.</span></span> <span data-ttu-id="62d3b-108">操作レコーダーの紹介では、頻繁に実行する Excel 操作の記録方法を説明します。</span><span class="sxs-lookup"><span data-stu-id="62d3b-108">You'll be introduced to the Action Recorder and see how to record your frequent Excel actions.</span></span> <span data-ttu-id="62d3b-109">また、コード エディターを使用して、独自のスクリプトを作成したり更新したりする方法についても説明します。</span><span class="sxs-lookup"><span data-stu-id="62d3b-109">You'll also learn how to make or update your own scripts with the Code Editor.</span></span>

<br>

> [!VIDEO https://www.microsoft.com/videoplayer/embed/RE4qdFF]

[!INCLUDE [Preview note](../includes/preview-note.md)]

## <a name="when-to-use-office-scripts"></a><span data-ttu-id="62d3b-110">Office スクリプトの使用に適した状況</span><span class="sxs-lookup"><span data-stu-id="62d3b-110">When to use Office Scripts</span></span>

<span data-ttu-id="62d3b-111">スクリプトを使用すると、自分が行った Excel の操作を記録して、さまざまなブックやワークシートに対してその操作を再現できます。</span><span class="sxs-lookup"><span data-stu-id="62d3b-111">Scripts allow you to record and replay your Excel actions on different workbooks and worksheets.</span></span> <span data-ttu-id="62d3b-112">同じ操作を何度も繰り返し行う必要がある場合は、Office スクリプトを使用すると、ワークフロー全体を 1 度ボタンを押すだけの操作に短縮できます。</span><span class="sxs-lookup"><span data-stu-id="62d3b-112">If you find yourself doing the same things over and over again, an Office Script can help you by reducing your whole workflow to a single button press.</span></span>

<span data-ttu-id="62d3b-113">たとえば、毎日仕事の始めに Excel で会計サイトから .csv ファイルを開いているとします。</span><span class="sxs-lookup"><span data-stu-id="62d3b-113">As an example, say you start your work day by opening a .csv file from an accounting site in Excel.</span></span> <span data-ttu-id="62d3b-114">それから数分かけて、不要な列を削除し、テーブルの書式を設定し、数式を追加し、新しいワークシートにピボットテーブルを作成します。</span><span class="sxs-lookup"><span data-stu-id="62d3b-114">You then spend several minutes deleting unnecessary columns, formatting a table, adding formulas, and creating a PivotTable in a new worksheet.</span></span> <span data-ttu-id="62d3b-115">毎日繰り返しているこのような操作を、操作レコーダーで 1 回記録できます。</span><span class="sxs-lookup"><span data-stu-id="62d3b-115">Those actions you repeat daily can be recorded once with the Action Recorder.</span></span> <span data-ttu-id="62d3b-116">それ以降は、スクリプトを実行するだけで、.csv の変換処理すべてが自動的に実行されます。</span><span class="sxs-lookup"><span data-stu-id="62d3b-116">For then on, running the script will take care of your entire .csv conversion.</span></span> <span data-ttu-id="62d3b-117">手順を忘れる危険がなくなるだけでなく、特に操作を教えなくても他の人とプロセスを共有することもできます。</span><span class="sxs-lookup"><span data-stu-id="62d3b-117">You'll not only remove the risk of forgetting steps, but be able to share your process with others without having to teach them anything.</span></span> <span data-ttu-id="62d3b-118">Office スクリプトを使用すると一般的なタスクを自動化できるので、自分自身と職場の作業効率や生産性を向上できます。</span><span class="sxs-lookup"><span data-stu-id="62d3b-118">Office Scripts automate your common tasks so you and your workplace can be more efficient and productive.</span></span>

## <a name="action-recorder"></a><span data-ttu-id="62d3b-119">操作レコーダー</span><span class="sxs-lookup"><span data-stu-id="62d3b-119">Action Recorder</span></span>

![いくつかの操作を記録した後の操作レコーダー。](../images/action-recorder-intro.png)

<span data-ttu-id="62d3b-121">操作レコーダーは、ユーザーが Excel で実行した操作を記録し、その操作をスクリプトに変換します。</span><span class="sxs-lookup"><span data-stu-id="62d3b-121">The Action Recorder records actions you take in Excel and translates them into a script.</span></span> <span data-ttu-id="62d3b-122">操作レコーダーを実行すると、セルの編集、書式の変更、テーブルの作成などの Excel の操作をキャプチャできます。</span><span class="sxs-lookup"><span data-stu-id="62d3b-122">With the Action recorder running, you can capture the Excel actions as you edit cells, change formatting, and create tables.</span></span> <span data-ttu-id="62d3b-123">作成されたスクリプトは、他のワークシートやブックで実行して、ユーザーが実行した元の操作を再現することもできます。</span><span class="sxs-lookup"><span data-stu-id="62d3b-123">The resulting script can be run on other worksheets and workbooks to recreate your original actions.</span></span>

## <a name="code-editor"></a><span data-ttu-id="62d3b-124">コード エディター</span><span class="sxs-lookup"><span data-stu-id="62d3b-124">Code Editor</span></span>

![上記のスクリプトのスクリプト コードを表示しているコード エディター。](../images/code-editor-intro.png)

<span data-ttu-id="62d3b-126">操作レコーダーで記録したすべてのスクリプトは、コード エディターで編集できます。</span><span class="sxs-lookup"><span data-stu-id="62d3b-126">All scripts recorded with the Action Recorder can be edited through the Code Editor.</span></span> <span data-ttu-id="62d3b-127">これにより、ニーズにぴったり合うようにスクリプトを微調整したり、カスタマイズしたりできます。</span><span class="sxs-lookup"><span data-stu-id="62d3b-127">This lets you tweak and customize the script to better suit your exact needs.</span></span> <span data-ttu-id="62d3b-128">また、条件付きステートメント (if/else) やループなど、Excel の UI からでは直接アクセスできないロジックや機能を追加することもできます。</span><span class="sxs-lookup"><span data-stu-id="62d3b-128">You can also add logic and functionality that is not directly accessible through the Excel UI, such as conditional statements (if/else) and loops.</span></span>

<span data-ttu-id="62d3b-129">Office スクリプトの機能を学習する簡単な方法の 1 つは、Excel on the web でスクリプトを記録し、作成されたコードを表示することです。</span><span class="sxs-lookup"><span data-stu-id="62d3b-129">One easy way to start learning the capabilities of Office Scripts is to record scripts in Excel on the web and view the resulting code.</span></span> <span data-ttu-id="62d3b-130">別の方法としては、用意されている[チュートリアル](../tutorials/excel-tutorial.md)に従うと、詳しいガイド付きで、より体系的に学習できます。</span><span class="sxs-lookup"><span data-stu-id="62d3b-130">Another option is to follow our [tutorials](../tutorials/excel-tutorial.md) to learn in a more guided and structured way.</span></span>

## <a name="next-steps"></a><span data-ttu-id="62d3b-131">次の手順</span><span class="sxs-lookup"><span data-stu-id="62d3b-131">Next steps</span></span>

<span data-ttu-id="62d3b-132">[Excel on the web の Office スクリプトに関するチュートリアル](../tutorials/excel-tutorial.md)を完了すると、Office スクリプトを初めて作成する方法を理解できます。</span><span class="sxs-lookup"><span data-stu-id="62d3b-132">Complete the [Office Scripts in Excel on the web tutorial](../tutorials/excel-tutorial.md) to learn how to create your first Office Scripts.</span></span>

## <a name="see-also"></a><span data-ttu-id="62d3b-133">関連項目</span><span class="sxs-lookup"><span data-stu-id="62d3b-133">See also</span></span>

- [<span data-ttu-id="62d3b-134">Excel on the web での Office スクリプトのスクリプトの基本事項</span><span class="sxs-lookup"><span data-stu-id="62d3b-134">Scripting fundamentals for Office Script in Excel on the web</span></span>](../develop/scripting-fundamentals.md)
- [<span data-ttu-id="62d3b-135">Office スクリプト API リファレンス</span><span class="sxs-lookup"><span data-stu-id="62d3b-135">Office Scripts API reference</span></span>](/javascript/api/office-scripts/overview)
- [<span data-ttu-id="62d3b-136">Office スクリプトのトラブルシューティング</span><span class="sxs-lookup"><span data-stu-id="62d3b-136">Troubleshooting Office Scripts</span></span>](../testing/troubleshooting.md)
- [<span data-ttu-id="62d3b-137">M365 での Office スクリプトの設定</span><span class="sxs-lookup"><span data-stu-id="62d3b-137">Office Scripts settings in M365</span></span>](https://support.office.com/article/office-scripts-settings-in-m365-19d3c51a-6ca2-40ab-978d-60fa49554dcf)
- [<span data-ttu-id="62d3b-138">Excel の Office スクリプトの概要 (support.office.com)</span><span class="sxs-lookup"><span data-stu-id="62d3b-138">Introduction to Office Scripts in Excel (on support.office.com)</span></span>](https://support.office.com/article/introduction-to-office-scripts-in-excel-9fbe283d-adb8-4f13-a75b-a81c6baf163a)
