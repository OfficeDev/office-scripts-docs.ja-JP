---
title: 手動 Power Automation フローからスクリプトを呼び出す
description: Power Automate の Office スクリプトで、手動のトリガーを使う方法を説明します。
ms.date: 12/28/2020
localization_priority: Priority
ms.openlocfilehash: 0a5fc93dbad1ee9804840fa11a06b689b7e7abda
ms.sourcegitcommit: f7a7aebfb687f2a35dbed07ed62ff352a114525a
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 05/06/2021
ms.locfileid: "52232873"
---
# <a name="call-scripts-from-a-manual-power-automate-flow-preview"></a><span data-ttu-id="ee535-103">手動 Power Automation フローからスクリプトを呼び出す (プレビュー)</span><span class="sxs-lookup"><span data-stu-id="ee535-103">Call scripts from a manual Power Automate flow (preview)</span></span>

<span data-ttu-id="ee535-104">このチュートリアルでは、[Power Automate](https://flow.microsoft.com)を使用して、Excel on the web 用の Office スクリプトを実行する方法について説明します。</span><span class="sxs-lookup"><span data-stu-id="ee535-104">This tutorial teaches you how to run an Office Script for Excel on the web through [Power Automate](https://flow.microsoft.com).</span></span> <span data-ttu-id="ee535-105">現在の時刻で 2 つのセルの値を更新するスクリプトを作成します。</span><span class="sxs-lookup"><span data-stu-id="ee535-105">You'll make a script that updates the values of two cells with the current time.</span></span> <span data-ttu-id="ee535-106">次に、このスクリプトを手動でトリガーした Power Automate フローに接続し、Power Automate のボタンを押したときにいつでもこのスクリプトが実行されるようにします。</span><span class="sxs-lookup"><span data-stu-id="ee535-106">You'll then connect that script to a manually triggered Power Automate flow, so that the script is run whenever a button in Power Automate is pressed.</span></span> <span data-ttu-id="ee535-107">基本的なパターンを理解したら、フローを拡大して他のアプリケーションを含めることができ、毎日のワークフローの自動化を進めることが可能です。</span><span class="sxs-lookup"><span data-stu-id="ee535-107">Once you understand the basic pattern, you can expand the flow to include other applications and automate more of your daily workflow.</span></span>

> [!TIP]
> <span data-ttu-id="ee535-108">Office スクリプトを初めて使用する場合は、チュートリアルの「[Excel on the web で Office スクリプトを記録、編集、作成する](excel-tutorial.md)」から始めることをお勧めします。</span><span class="sxs-lookup"><span data-stu-id="ee535-108">If you are new to Office Scripts, we recommend starting with the [Record, edit, and create Office Scripts in Excel on the web](excel-tutorial.md) tutorial.</span></span> <span data-ttu-id="ee535-109">[Office スクリプトは TypeScript を使用](../overview/code-editor-environment.md)します。このチュートリアルは、JavaScript や TypeScript について初級から中級レベルの知識を持つユーザーを対象としています。</span><span class="sxs-lookup"><span data-stu-id="ee535-109">[Office Scripts use TypeScript](../overview/code-editor-environment.md) and this tutorial is intended for people with beginner to intermediate-level knowledge of JavaScript or TypeScript.</span></span> <span data-ttu-id="ee535-110">JavaScript を使い慣れていない場合は、「[Mozilla の JavaScript チュートリアル](https://developer.mozilla.org/docs/Web/JavaScript/Guide/Introduction)」から始めることをお勧めします。</span><span class="sxs-lookup"><span data-stu-id="ee535-110">If you're new to JavaScript, we recommend starting with the [Mozilla JavaScript tutorial](https://developer.mozilla.org/docs/Web/JavaScript/Guide/Introduction).</span></span>

## <a name="prerequisites"></a><span data-ttu-id="ee535-111">前提条件</span><span class="sxs-lookup"><span data-stu-id="ee535-111">Prerequisites</span></span>

[!INCLUDE [Tutorial prerequisites](../includes/power-automate-tutorial-prerequisites.md)]

## <a name="prepare-the-workbook"></a><span data-ttu-id="ee535-112">ブックを準備する</span><span class="sxs-lookup"><span data-stu-id="ee535-112">Prepare the workbook</span></span>

<span data-ttu-id="ee535-113">Power Automate では、ブック コンポーネントにアクセスするために `Workbook.getActiveWorksheet` などの[相対参照](../testing/power-automate-troubleshooting.md#avoid-using-relative-references)を使わないようにする必要があります。</span><span class="sxs-lookup"><span data-stu-id="ee535-113">Power Automate shouldn't use [relative references](../testing/power-automate-troubleshooting.md#avoid-using-relative-references) like `Workbook.getActiveWorksheet` to access workbook components.</span></span> <span data-ttu-id="ee535-114">したがって、Power Automate が参照できる、名前が統一されたワークブックとワークシートが必要です。</span><span class="sxs-lookup"><span data-stu-id="ee535-114">So, we need a workbook and worksheet with consistent names that Power Automate can reference.</span></span>

1. <span data-ttu-id="ee535-115">**MyWorkbook** という名前の新しいブックを作成します。</span><span class="sxs-lookup"><span data-stu-id="ee535-115">Create a new workbook named **MyWorkbook**.</span></span>

2. <span data-ttu-id="ee535-116">**MyWorkbook** というワークブック内に、**TutorialWorksheet** という名前のワークシートを作成します。</span><span class="sxs-lookup"><span data-stu-id="ee535-116">In the **MyWorkbook** workbook, create a worksheet called **TutorialWorksheet**.</span></span>

## <a name="create-an-office-script"></a><span data-ttu-id="ee535-117">オフィス スクリプトを作成する</span><span class="sxs-lookup"><span data-stu-id="ee535-117">Create an Office Script</span></span>

1. <span data-ttu-id="ee535-118">**[オートメーション]** タブに移動して **[すべてのスクリプト]** を選択します。</span><span class="sxs-lookup"><span data-stu-id="ee535-118">Go to the **Automate** tab and select **All Scripts**.</span></span>

2. <span data-ttu-id="ee535-119">**[新しいスクリプト]** を選択します。</span><span class="sxs-lookup"><span data-stu-id="ee535-119">Select **New Script**.</span></span>

3. <span data-ttu-id="ee535-120">既定のスクリプトを次のスクリプトに置き換えます。</span><span class="sxs-lookup"><span data-stu-id="ee535-120">Replace the default script with the following script.</span></span> <span data-ttu-id="ee535-121">このスクリプトは、**TutorialWorksheet** というワークシートの最初の 2 つのセルに現在の日付と時刻を追加します。</span><span class="sxs-lookup"><span data-stu-id="ee535-121">This script adds the current date and time to the first two cells of the **TutorialWorksheet** worksheet.</span></span>

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

4. <span data-ttu-id="ee535-122">スクリプトの名前を **[日付と時刻の設定]** に変更します。</span><span class="sxs-lookup"><span data-stu-id="ee535-122">Rename the script to **Set date and time**.</span></span> <span data-ttu-id="ee535-123">スクリプト名を押して変更します。</span><span class="sxs-lookup"><span data-stu-id="ee535-123">Press the script name to change it.</span></span>

5. <span data-ttu-id="ee535-124">スクリプトを保存するには **[スクリプトの保存]** を押します。</span><span class="sxs-lookup"><span data-stu-id="ee535-124">Save the script by pressing **Save Script**.</span></span>

## <a name="create-an-automated-workflow-with-power-automate"></a><span data-ttu-id="ee535-125">Power Automate を使用して自動化されたワークフローを作成する</span><span class="sxs-lookup"><span data-stu-id="ee535-125">Create an automated workflow with Power Automate</span></span>

1. <span data-ttu-id="ee535-126">[「Power Automate のサイト」](https://flow.microsoft.com)にサインインします。</span><span class="sxs-lookup"><span data-stu-id="ee535-126">Sign in to the [Power Automate site](https://flow.microsoft.com).</span></span>

2. <span data-ttu-id="ee535-127">画面の左側に表示されるメニューで、**[作成]** を押します。</span><span class="sxs-lookup"><span data-stu-id="ee535-127">In the menu that's displayed on the left side of the screen, press **Create**.</span></span> <span data-ttu-id="ee535-128">これにより、新しいワークフローを作成する方法の一覧を表示できます。</span><span class="sxs-lookup"><span data-stu-id="ee535-128">This brings you to list of ways to create new workflows.</span></span>

    :::image type="content" source="../images/power-automate-tutorial-1.png" alt-text="Power Automate の [作成] ボタン":::

3. <span data-ttu-id="ee535-130">**[白紙から初める]** セクションで、**[インスタント フロー]** を選択します。</span><span class="sxs-lookup"><span data-stu-id="ee535-130">In the **Start from blank** section, select **Instant flow**.</span></span> <span data-ttu-id="ee535-131">これで、手動でアクティベートされたワークフローが作成されます。</span><span class="sxs-lookup"><span data-stu-id="ee535-131">This creates a manually activated workflow.</span></span>

    :::image type="content" source="../images/power-automate-tutorial-2.png" alt-text="新しいワークフローを作成するための Power Automate インスタント フロー オプション":::

4. <span data-ttu-id="ee535-133">表示されたダイアログ ウィンドウで、フローの名前を **[フロー名]** テキスト ボックスに入力し、**[フローをトリガーする方法の選択]** 内のオプションの一覧から **[手動でフローをトリガーする]** を選択し、**[作成]** を押します。</span><span class="sxs-lookup"><span data-stu-id="ee535-133">In the dialog window that appears, enter a name for your flow in the **Flow name** text box, select **Manually trigger a flow** from the list of options under **Choose how to trigger the flow**, and press **Create**.</span></span>

    :::image type="content" source="../images/power-automate-tutorial-3.png" alt-text="Power Automate の [手動でフローをトリガーする] オプション":::

    <span data-ttu-id="ee535-135">手動でトリガーするフローは、いくつかあるフローの種類のうちの 1 つです。</span><span class="sxs-lookup"><span data-stu-id="ee535-135">Note that a manually triggered flow is just one of many types of flows.</span></span> <span data-ttu-id="ee535-136">次のチュートリアルでは、メールを受信したときに自動的に実行されるフローを作成します。</span><span class="sxs-lookup"><span data-stu-id="ee535-136">In the next tutorial, you'll make a flow that automatically runs when you receive an email.</span></span>

5. <span data-ttu-id="ee535-137">**[新しいステップ]** を押します。</span><span class="sxs-lookup"><span data-stu-id="ee535-137">Press **New step**.</span></span>

6. <span data-ttu-id="ee535-138">**[標準]** タブを選択し、**Excel Online (ビジネス)** を選択します。</span><span class="sxs-lookup"><span data-stu-id="ee535-138">Select the **Standard** tab, then select **Excel Online (Business)**.</span></span>

    :::image type="content" source="../images/power-automate-tutorial-4.png" alt-text="Power Automate の [Excel Online (Business)] オプション":::

7. <span data-ttu-id="ee535-140">**[アクション]** の下の **[スクリプトの実行 (プレビュー)]** を選択します。</span><span class="sxs-lookup"><span data-stu-id="ee535-140">Under **Actions**, select **Run script (preview)**.</span></span>

    :::image type="content" source="../images/power-automate-tutorial-5.png" alt-text="Power Automate の [スクリプトの実行 (プレビュー)] アクションのオプション":::

8. <span data-ttu-id="ee535-142">次に、フロー ステップで使用するブックおよびスクリプトを選択します。</span><span class="sxs-lookup"><span data-stu-id="ee535-142">Next, you'll select the workbook and script to use in the flow step.</span></span> <span data-ttu-id="ee535-143">このチュートリアルでは、OneDrive に作成したブックを使用しますが、OneDrive サイトまたは SharePoint サイトでは任意のブックを使用できます。</span><span class="sxs-lookup"><span data-stu-id="ee535-143">For the tutorial, you'll use the workbook you created in your OneDrive, but you could use any workbook in a OneDrive or SharePoint site.</span></span> <span data-ttu-id="ee535-144">**スクリプトの実行** コネクタには、次の設定を指定します。</span><span class="sxs-lookup"><span data-stu-id="ee535-144">Specify the following settings for the **Run script** connector:</span></span>

    - <span data-ttu-id="ee535-145">**場所**: OneDrive for Business</span><span class="sxs-lookup"><span data-stu-id="ee535-145">**Location**: OneDrive for Business</span></span>
    - <span data-ttu-id="ee535-146">**ドキュメント ライブラリ**: OneDrive</span><span class="sxs-lookup"><span data-stu-id="ee535-146">**Document Library**: OneDrive</span></span>
    - <span data-ttu-id="ee535-147">**ファイル**: MyWorkbook.xlsx *(ファイル ブラウザーを使用して選択されています)*</span><span class="sxs-lookup"><span data-stu-id="ee535-147">**File**: MyWorkbook.xlsx *(Chosen through the file browser)*</span></span>
    - <span data-ttu-id="ee535-148">**スクリプト**: 日時を設定</span><span class="sxs-lookup"><span data-stu-id="ee535-148">**Script**: Set date and time</span></span>

    :::image type="content" source="../images/power-automate-tutorial-6.png" alt-text="スクリプトを実行するための Power Automate コネクタの設定":::

9. <span data-ttu-id="ee535-150">**[保存]** を押します。</span><span class="sxs-lookup"><span data-stu-id="ee535-150">Press **Save**.</span></span>

<span data-ttu-id="ee535-151">これで、フローは Power Automate で実行できるようになりました。</span><span class="sxs-lookup"><span data-stu-id="ee535-151">Your flow is now ready to be run through Power Automate.</span></span> <span data-ttu-id="ee535-152">フロー エディターの **[テスト]** ボタンを使用してテストするか、チュートリアルの残りの手順に従って、フロー コレクションからフローを実行できます。</span><span class="sxs-lookup"><span data-stu-id="ee535-152">You can test it using the **Test** button in the flow editor or follow the remaining tutorial steps to run the flow from your flow collection.</span></span>

## <a name="run-the-script-through-power-automate"></a><span data-ttu-id="ee535-153">Power Automate でスクリプトを実行する</span><span class="sxs-lookup"><span data-stu-id="ee535-153">Run the script through Power Automate</span></span>

1. <span data-ttu-id="ee535-154">Power Automate のメイン ページで、**[自分のフロー]** を選択します。</span><span class="sxs-lookup"><span data-stu-id="ee535-154">From the main Power Automate page, select **My flows**.</span></span>

    :::image type="content" source="../images/power-automate-tutorial-7.png" alt-text="Power Automate の [自分のフロー] ボタン":::

2. <span data-ttu-id="ee535-156">**[自分のフロー]** タブに表示されているフローの一覧から、**[自分のチュートリアル フロー]** を選択すると、以前に作成したフローの詳細が表示されます。</span><span class="sxs-lookup"><span data-stu-id="ee535-156">Select **My tutorial flow** from the list of flows displayed in the **My flows** tab. This shows the details of the flow we previously created.</span></span>

3. <span data-ttu-id="ee535-157">**[実行]** を押します。</span><span class="sxs-lookup"><span data-stu-id="ee535-157">Press **Run**.</span></span>

    :::image type="content" source="../images/power-automate-tutorial-8.png" alt-text="Power Automate の [実行] ボタン":::

4. <span data-ttu-id="ee535-159">フローを実行するための作業ウィンドウが表示されます。</span><span class="sxs-lookup"><span data-stu-id="ee535-159">A task pane will appear for running the flow.</span></span> <span data-ttu-id="ee535-160">Excel Online への **サインイン** を要求された場合は、**[続ける]** を押します。</span><span class="sxs-lookup"><span data-stu-id="ee535-160">If you are asked to **Sign in** to Excel Online, do so by pressing **Continue**.</span></span>

5. <span data-ttu-id="ee535-161">**[フローの実行]** を押します。</span><span class="sxs-lookup"><span data-stu-id="ee535-161">Press **Run flow**.</span></span> <span data-ttu-id="ee535-162">これにより、関連する Office スクリプトを実行するフローが実行されます。</span><span class="sxs-lookup"><span data-stu-id="ee535-162">This runs the flow, which runs the related Office Script.</span></span>

6. <span data-ttu-id="ee535-163">**[完了]** を押します。</span><span class="sxs-lookup"><span data-stu-id="ee535-163">Press **Done**.</span></span> <span data-ttu-id="ee535-164">それに応じて **[実行]** セクションが更新されます。</span><span class="sxs-lookup"><span data-stu-id="ee535-164">You should see the **Runs** section update accordingly.</span></span>

7. <span data-ttu-id="ee535-165">ページを更新して、Power Automate の結果を表示します。</span><span class="sxs-lookup"><span data-stu-id="ee535-165">Refresh the page to see the results of the Power Automate.</span></span> <span data-ttu-id="ee535-166">成功した場合は、ワークブックに移動して、更新されたセルを確認します。</span><span class="sxs-lookup"><span data-stu-id="ee535-166">If it succeeded, go to the workbook to see the updated cells.</span></span> <span data-ttu-id="ee535-167">エラーが発生した場合は、フローの設定を確認し、もう一度実行します。</span><span class="sxs-lookup"><span data-stu-id="ee535-167">If it failed, verify the flow's settings and run it a second time.</span></span>

    :::image type="content" source="../images/power-automate-tutorial-9.png" alt-text="正常にフローが発生したことを示す Power Automate 出力":::

## <a name="next-steps"></a><span data-ttu-id="ee535-169">次の手順</span><span class="sxs-lookup"><span data-stu-id="ee535-169">Next steps</span></span>

<span data-ttu-id="ee535-170">[「自動で実行される Power Automate フロー内で、データをスクリプトに渡す」](excel-power-automate-trigger.md)のチュートリアルを完了します。</span><span class="sxs-lookup"><span data-stu-id="ee535-170">Complete the [Pass data to scripts in an automatically-run Power Automate flow](excel-power-automate-trigger.md) tutorial.</span></span> <span data-ttu-id="ee535-171">このコースでは、ワークフロー サービスから Office スクリプトにデータを渡す方法と、特定のイベントが発生したときに Power Automate フローを実行する方法について説明します。</span><span class="sxs-lookup"><span data-stu-id="ee535-171">It teaches you how to pass data from a workflow service to your Office Script and run the Power Automate flow when certain events occur.</span></span>
