---
title: 手動 Power Automation フローからスクリプトを呼び出す
description: Power Automate の Office スクリプトで、手動のトリガーを使う方法を説明します。
ms.date: 07/14/2020
localization_priority: Priority
ms.openlocfilehash: 70fca2620973ecefe9eda40f02e28f064b713677
ms.sourcegitcommit: ebd1079c7e2695ac0e7e4c616f2439975e196875
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 07/17/2020
ms.locfileid: "45160442"
---
# <a name="call-scripts-from-a-manual-power-automate-flow-preview"></a><span data-ttu-id="05459-103">手動 Power Automation フローからスクリプトを呼び出す (プレビュー)</span><span class="sxs-lookup"><span data-stu-id="05459-103">Call scripts from a manual Power Automate flow (preview)</span></span>

<span data-ttu-id="05459-104">このチュートリアルでは、[Power Automate](https://flow.microsoft.com)を使用して、Excel on the web 用の Office スクリプトを実行する方法について説明します。</span><span class="sxs-lookup"><span data-stu-id="05459-104">This tutorial teaches you how to run an Office Script for Excel on the web through [Power Automate](https://flow.microsoft.com).</span></span>

## <a name="prerequisites"></a><span data-ttu-id="05459-105">前提条件</span><span class="sxs-lookup"><span data-stu-id="05459-105">Prerequisites</span></span>

[!INCLUDE [Tutorial prerequisites](../includes/power-automate-tutorial-prerequisites.md)]

> [!IMPORTANT]
> <span data-ttu-id="05459-106">このチュートリアルは、お客様が[「Excel on the web で Office スクリプトを記録、編集、作成する」](excel-tutorial.md)のチュートリアルを既に完了していることを前提にしています。</span><span class="sxs-lookup"><span data-stu-id="05459-106">This tutorial assumes you have completed the [Record, edit, and create Office Scripts in Excel on the web](excel-tutorial.md) tutorial.</span></span>

## <a name="prepare-the-workbook"></a><span data-ttu-id="05459-107">ブックを準備する</span><span class="sxs-lookup"><span data-stu-id="05459-107">Prepare the workbook</span></span>

<span data-ttu-id="05459-108">Power Automate は、`Workbook.getActiveWorksheet` のような相対参照を使用して、ブックのワークブック コンポーネントにアクセスすることはできません。</span><span class="sxs-lookup"><span data-stu-id="05459-108">Power Automate can't use relative references like `Workbook.getActiveWorksheet` to access workbook components.</span></span> <span data-ttu-id="05459-109">したがって、Power Automate が参照できる、名前が統一されたワークブックとワークシートが必要です。</span><span class="sxs-lookup"><span data-stu-id="05459-109">So, we need a workbook and worksheet with consistent names that Power Automate can reference.</span></span>

1. <span data-ttu-id="05459-110">**MyWorkbook** という名前の新しいブックを作成します。</span><span class="sxs-lookup"><span data-stu-id="05459-110">Create a new workbook named **MyWorkbook**.</span></span>

2. <span data-ttu-id="05459-111">**MyWorkbook** というワークブック内に、**TutorialWorksheet** という名前のワークシートを作成します。</span><span class="sxs-lookup"><span data-stu-id="05459-111">In the **MyWorkbook** workbook, create a worksheet called **TutorialWorksheet**.</span></span>

## <a name="create-an-office-script"></a><span data-ttu-id="05459-112">オフィス スクリプトを作成する</span><span class="sxs-lookup"><span data-stu-id="05459-112">Create an Office Script</span></span>

1. <span data-ttu-id="05459-113">**[オートメーション]** タブに移動して **[コード エディター]** を選択します。</span><span class="sxs-lookup"><span data-stu-id="05459-113">Go to the **Automate** tab and select **Code Editor**.</span></span>

2. <span data-ttu-id="05459-114">**[新しいスクリプト]** を選択します。</span><span class="sxs-lookup"><span data-stu-id="05459-114">Select **New Script**.</span></span>

3. <span data-ttu-id="05459-115">既定のスクリプトを次のスクリプトに置き換えます。</span><span class="sxs-lookup"><span data-stu-id="05459-115">Replace the default script with the following script.</span></span> <span data-ttu-id="05459-116">このスクリプトは、**TutorialWorksheet** というワークシートの最初の 2 つのセルに現在の日付と時刻を追加します。</span><span class="sxs-lookup"><span data-stu-id="05459-116">This script adds the current date and time to the first two cells of the **TutorialWorksheet** worksheet.</span></span>

    ```TypeScript
    function main(workbook: ExcelScript.Workbook) {
      // Get the "TutorialWorksheet" worksheet from the workbook.
      let worksheet = workbook.getWorksheet("TutorialWorksheet");

      // Get the cells at A1 and B1.
      let dateRange = worksheet.getRange("A1");
      let timeRange = worksheet.getRange("B1");

      // Get the current date and time using the JavaScript Date object.
      let date = new Date(Date.now());

      // Add the date string to A1.
      dateRange.setValue(date.toLocaleDateString());

      // Add the time string to B1.
      timeRange.setValue(date.toLocaleTimeString());
    }
    ```

4. <span data-ttu-id="05459-117">スクリプトの名前を **[日付と時刻の設定]** に変更します。</span><span class="sxs-lookup"><span data-stu-id="05459-117">Rename the script to **Set date and time**.</span></span> <span data-ttu-id="05459-118">スクリプト名を押して変更します。</span><span class="sxs-lookup"><span data-stu-id="05459-118">Press the script name to change it.</span></span>

5. <span data-ttu-id="05459-119">スクリプトを保存するには **[スクリプトの保存]** を押します。</span><span class="sxs-lookup"><span data-stu-id="05459-119">Save the script by pressing **Save Script**.</span></span>

## <a name="create-an-automated-workflow-with-power-automate"></a><span data-ttu-id="05459-120">Power Automate を使用して自動化されたワークフローを作成する</span><span class="sxs-lookup"><span data-stu-id="05459-120">Create an automated workflow with Power Automate</span></span>

1. <span data-ttu-id="05459-121">[「Power Automate のサイト」](https://flow.microsoft.com)にサインインします。</span><span class="sxs-lookup"><span data-stu-id="05459-121">Sign in to the [Power Automate site](https://flow.microsoft.com).</span></span>

2. <span data-ttu-id="05459-122">画面の左側に表示されるメニューで、**[作成]** を押します。</span><span class="sxs-lookup"><span data-stu-id="05459-122">In the menu that's displayed on the left side of the screen, press **Create**.</span></span> <span data-ttu-id="05459-123">これにより、新しいワークフローを作成する方法の一覧を表示できます。</span><span class="sxs-lookup"><span data-stu-id="05459-123">This brings you to list of ways to create new workflows.</span></span>

    ![Power Automate の [作成] ボタン。](../images/power-automate-tutorial-1.png)

3. <span data-ttu-id="05459-125">**[白紙から初める]** セクションで、**[インスタント フロー]** を選択します。</span><span class="sxs-lookup"><span data-stu-id="05459-125">In the **Start from blank** section, select **Instant flow**.</span></span> <span data-ttu-id="05459-126">これで、手動でアクティベートされたワークフローが作成されます。</span><span class="sxs-lookup"><span data-stu-id="05459-126">This creates a manually activated workflow.</span></span>

    ![新しいワークフローを作成するための [クイックフロー] オプション。](../images/power-automate-tutorial-2.png)

4. <span data-ttu-id="05459-128">表示されたダイアログ ウィンドウで、フローの名前を **[フロー名]** テキスト ボックスに入力し、**[フローをトリガーする方法の選択]** 内のオプションの一覧から **[手動でフローをトリガーする]** を選択し、**[作成]** を押します。</span><span class="sxs-lookup"><span data-stu-id="05459-128">In the dialog window that appears, enter a name for your flow in the **Flow name** text box, select **Manually trigger a flow** from the list of options under **Choose how to trigger the flow**, and press **Create**.</span></span>

    ![新しいインスタント フローを作成するための手動トリガー オプション。](../images/power-automate-tutorial-3.png)

    <span data-ttu-id="05459-130">手動でトリガーするフローは、いくつかあるフローの種類のうちの 1 つです。</span><span class="sxs-lookup"><span data-stu-id="05459-130">Note that a manually triggered flow is just one of many types of flows.</span></span> <span data-ttu-id="05459-131">次のチュートリアルでは、メールを受信したときに自動的に実行されるフローを作成します。</span><span class="sxs-lookup"><span data-stu-id="05459-131">In the next tutorial, you'll make a flow that automatically runs when you receive an email.</span></span>

5. <span data-ttu-id="05459-132">**[新しいステップ]** を押します。</span><span class="sxs-lookup"><span data-stu-id="05459-132">Press **New step**.</span></span>

6. <span data-ttu-id="05459-133">**[標準]** タブを選択し、**Excel Online (ビジネス)** を選択します。</span><span class="sxs-lookup"><span data-stu-id="05459-133">Select the **Standard** tab, then select **Excel Online (Business)**.</span></span>

    ![Excel Online (ビジネス) 用の Power Automate オプション。](../images/power-automate-tutorial-4.png)

7. <span data-ttu-id="05459-135">**[アクション]** の下から、**[スクリプトの実行 (プレビュー)]** を選択します。</span><span class="sxs-lookup"><span data-stu-id="05459-135">Under **Actions**, select **Run script (preview)**.</span></span>

    ![スクリプトの実行 (プレビュー)用の Power Automate アクションのオプション。](../images/power-automate-tutorial-5.png)

8. <span data-ttu-id="05459-137">**スクリプトの実行**コネクタには、次の設定を指定します。</span><span class="sxs-lookup"><span data-stu-id="05459-137">Specify the following settings for the **Run script** connector:</span></span>

    - <span data-ttu-id="05459-138">**場所**: OneDrive for Business</span><span class="sxs-lookup"><span data-stu-id="05459-138">**Location**: OneDrive for Business</span></span>
    - <span data-ttu-id="05459-139">**ドキュメント ライブラリ**: OneDrive</span><span class="sxs-lookup"><span data-stu-id="05459-139">**Document Library**: OneDrive</span></span>
    - <span data-ttu-id="05459-140">**ファイル**: MyWorkbook.xlsx</span><span class="sxs-lookup"><span data-stu-id="05459-140">**File**: MyWorkbook.xlsx</span></span>
    - <span data-ttu-id="05459-141">**スクリプト**: 日時を設定</span><span class="sxs-lookup"><span data-stu-id="05459-141">**Script**: Set date and time</span></span>

    ![Power Automate でスクリプトを実行するためのコネクタの設定。](../images/power-automate-tutorial-6.png)

9. <span data-ttu-id="05459-143">**[保存]** を押します。</span><span class="sxs-lookup"><span data-stu-id="05459-143">Press **Save**.</span></span>

<span data-ttu-id="05459-144">これで、フローは Power Automate で実行できるようになりました。</span><span class="sxs-lookup"><span data-stu-id="05459-144">Your flow is now ready to be run through Power Automate.</span></span> <span data-ttu-id="05459-145">フロー エディターの **[テスト]** ボタンを使用してテストするか、チュートリアルの残りの手順に従って、フロー コレクションからフローを実行できます。</span><span class="sxs-lookup"><span data-stu-id="05459-145">You can test it using the **Test** button in the flow editor or follow the remaining tutorial steps to run the flow from your flow collection.</span></span>

## <a name="run-the-script-through-power-automate"></a><span data-ttu-id="05459-146">Power Automate でスクリプトを実行する</span><span class="sxs-lookup"><span data-stu-id="05459-146">Run the script through Power Automate</span></span>

1. <span data-ttu-id="05459-147">Power Automate のメイン ページで、**[自分のフロー]** を選択します。</span><span class="sxs-lookup"><span data-stu-id="05459-147">From the main Power Automate page, select **My flows**.</span></span>

    ![Power Automate の [自分のフロー] ボタン。](../images/power-automate-tutorial-7.png)

2. <span data-ttu-id="05459-149">**[自分のフロー]** タブに表示されているフローの一覧から、**[自分のチュートリアル フロー]** を選択すると、以前に作成したフローの詳細が表示されます。</span><span class="sxs-lookup"><span data-stu-id="05459-149">Select **My tutorial flow** from the list of flows displayed in the **My flows** tab. This shows the details of the flow we previously created.</span></span>

3. <span data-ttu-id="05459-150">**[実行]** を押します。</span><span class="sxs-lookup"><span data-stu-id="05459-150">Press **Run**.</span></span>

    ![Power Automate の [実行] ボタン。](../images/power-automate-tutorial-8.png)

4. <span data-ttu-id="05459-152">フローを実行するための作業ウィンドウが表示されます。</span><span class="sxs-lookup"><span data-stu-id="05459-152">A task pane will appear for running the flow.</span></span> <span data-ttu-id="05459-153">Excel Online への**サインイン** を要求された場合は、**[続ける]** を押します。</span><span class="sxs-lookup"><span data-stu-id="05459-153">If you are asked to **Sign in** to Excel Online, do so by pressing **Continue**.</span></span>

5. <span data-ttu-id="05459-154">**[フローの実行]** を押します。</span><span class="sxs-lookup"><span data-stu-id="05459-154">Press **Run flow**.</span></span> <span data-ttu-id="05459-155">これにより、関連する Office スクリプトを実行するフローが実行されます。</span><span class="sxs-lookup"><span data-stu-id="05459-155">This runs the flow, which runs the related Office Script.</span></span>

6. <span data-ttu-id="05459-156">**[完了]** を押します。</span><span class="sxs-lookup"><span data-stu-id="05459-156">Press **Done**.</span></span> <span data-ttu-id="05459-157">それに応じて **[実行]** セクションが更新されます。</span><span class="sxs-lookup"><span data-stu-id="05459-157">You should see the **Runs** section update accordingly.</span></span>

7. <span data-ttu-id="05459-158">ページを更新して、Power Automate の結果を表示します。</span><span class="sxs-lookup"><span data-stu-id="05459-158">Refresh the page to see the results of the Power Automate.</span></span> <span data-ttu-id="05459-159">成功した場合は、ワークブックに移動して、更新されたセルを確認します。</span><span class="sxs-lookup"><span data-stu-id="05459-159">If it succeeded, go to the workbook to see the updated cells.</span></span> <span data-ttu-id="05459-160">エラーが発生した場合は、フローの設定を確認し、もう一度実行します。</span><span class="sxs-lookup"><span data-stu-id="05459-160">If it failed, verify the flow's settings and run it a second time.</span></span>

    ![正常にフローが発生したことを示す Power Automate 出力。](../images/power-automate-tutorial-9.png)

## <a name="next-steps"></a><span data-ttu-id="05459-162">次の手順</span><span class="sxs-lookup"><span data-stu-id="05459-162">Next steps</span></span>

<span data-ttu-id="05459-163">[「自動で実行される Power Automate フロー内で、データをスクリプトに渡す」](excel-power-automate-trigger.md)のチュートリアルを完了します。</span><span class="sxs-lookup"><span data-stu-id="05459-163">Complete the [Pass data to scripts in an automatically-run Power Automate flow](excel-power-automate-trigger.md) tutorial.</span></span> <span data-ttu-id="05459-164">このコースでは、ワークフロー サービスから Office スクリプトにデータを渡す方法と、特定のイベントが発生したときに Power Automate フローを実行する方法について説明します。</span><span class="sxs-lookup"><span data-stu-id="05459-164">It teaches you how to pass data from a workflow service to your Office Script and run the Power Automate flow when certain events occur.</span></span>
