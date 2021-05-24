---
title: 自動で実行される Power Automate フローにスクリプトからデータを返す
description: Power Automate を使用して Excel on the web 用の Office スクリプトを実行してリマインダー メールを送信する方法を示すチュートリアル。
ms.date: 12/15/2020
localization_priority: Priority
ms.openlocfilehash: e7f1051076bf84cfbbec0fcdd72777766dbcf152
ms.sourcegitcommit: 4687693f02fc90a57ba30c461f35046e02e6f5fb
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 05/19/2021
ms.locfileid: "52545007"
---
# <a name="return-data-from-a-script-to-an-automatically-run-power-automate-flow"></a><span data-ttu-id="c7bac-103">自動で実行される Power Automate フローにスクリプトからデータを返す</span><span class="sxs-lookup"><span data-stu-id="c7bac-103">Return data from a script to an automatically-run Power Automate flow</span></span>

<span data-ttu-id="c7bac-104">このチュートリアルでは、自動化された [Power Automate](https://flow.microsoft.com) ワークフローの一部として、Excel on the web 用の Office スクリプトから情報を返す方法について説明します。</span><span class="sxs-lookup"><span data-stu-id="c7bac-104">This tutorial teaches you how to return information from an Office Script for Excel on the web as part of an automated [Power Automate](https://flow.microsoft.com) workflow.</span></span> <span data-ttu-id="c7bac-105">スケジュールを確認し、フローに従ってリマインダー メールを送信するスクリプトを作成します。</span><span class="sxs-lookup"><span data-stu-id="c7bac-105">You'll make a script that looks through a schedule and works with a flow to send reminder emails.</span></span> <span data-ttu-id="c7bac-106">このフローは定期的に実行され、ユーザーに代わってこれらのリマインダーを提供します。</span><span class="sxs-lookup"><span data-stu-id="c7bac-106">This flow will run on a regular schedule, providing these reminders on your behalf.</span></span>

> [!TIP]
> <span data-ttu-id="c7bac-107">Office スクリプトを初めて使用する場合は、チュートリアルの「[Excel on the web で Office スクリプトを記録、編集、作成する](excel-tutorial.md)」から始めることをお勧めします。</span><span class="sxs-lookup"><span data-stu-id="c7bac-107">If you are new to Office Scripts, we recommend starting with the [Record, edit, and create Office Scripts in Excel on the web](excel-tutorial.md) tutorial.</span></span>
>
> <span data-ttu-id="c7bac-108">Power Automate を初めて使用する場合は、チュートリアルの「[手動 Power Automate フローからスクリプトを呼び出す](excel-power-automate-manual.md)」と「[自動で実行される Power Automate フロー内で、データをスクリプトに渡す](excel-power-automate-trigger.md)」から始めることを勧めします。</span><span class="sxs-lookup"><span data-stu-id="c7bac-108">If you are new to Power Automate, we recommend starting with the [Call scripts from a manual Power Automate flow](excel-power-automate-manual.md) and [Pass data to scripts in an automatically-run Power Automate flow](excel-power-automate-trigger.md) tutorials.</span></span>
>
> <span data-ttu-id="c7bac-109">[Office スクリプトは TypeScript を使用](../overview/code-editor-environment.md)します。このチュートリアルは、JavaScript や TypeScript について初級から中級レベルの知識を持つユーザーを対象としています。</span><span class="sxs-lookup"><span data-stu-id="c7bac-109">[Office Scripts use TypeScript](../overview/code-editor-environment.md) and this tutorial is intended for people with beginner to intermediate-level knowledge of JavaScript or TypeScript.</span></span> <span data-ttu-id="c7bac-110">JavaScript を使い慣れていない場合は、「[Mozilla の JavaScript チュートリアル](https://developer.mozilla.org/docs/Web/JavaScript/Guide/Introduction)」から始めることをお勧めします。</span><span class="sxs-lookup"><span data-stu-id="c7bac-110">If you're new to JavaScript, we recommend starting with the [Mozilla JavaScript tutorial](https://developer.mozilla.org/docs/Web/JavaScript/Guide/Introduction).</span></span>

## <a name="prerequisites"></a><span data-ttu-id="c7bac-111">前提条件</span><span class="sxs-lookup"><span data-stu-id="c7bac-111">Prerequisites</span></span>

[!INCLUDE [Tutorial prerequisites](../includes/power-automate-tutorial-prerequisites.md)]

## <a name="prepare-the-workbook"></a><span data-ttu-id="c7bac-112">ブックを準備する</span><span class="sxs-lookup"><span data-stu-id="c7bac-112">Prepare the workbook</span></span>

1. <span data-ttu-id="c7bac-113">ブック <a href="on-call-rotation.xlsx">on-call-rotation.xlsx</a> を 自分の OneDrive にダウンロードします。</span><span class="sxs-lookup"><span data-stu-id="c7bac-113">Download the workbook <a href="on-call-rotation.xlsx">on-call-rotation.xlsx</a> to your OneDrive.</span></span>

1. <span data-ttu-id="c7bac-114">Excel on the web で **on-call-rotation.xlsx** を開きます。</span><span class="sxs-lookup"><span data-stu-id="c7bac-114">Open **on-call-rotation.xlsx** in Excel on the web.</span></span>

1. <span data-ttu-id="c7bac-115">テーブルに行を追加して、自分の名前、メール アドレス、および現在の日付と重なるように開始日と終了日を入力します。</span><span class="sxs-lookup"><span data-stu-id="c7bac-115">Add a row to the table with your name, email address, and start and end dates that overlap with the current date.</span></span>

    > [!IMPORTANT]
    > <span data-ttu-id="c7bac-116">これから作成するスクリプトは、テーブル内の最初に一致するエントリを使用するため、自分の名前が現在の週のどの行よりも上にあることを確認してください。</span><span class="sxs-lookup"><span data-stu-id="c7bac-116">The script you'll write uses the first matching entry in the table, so make sure your name is above any row with the current week.</span></span>

    :::image type="content" source="../images/power-automate-return-tutorial-1.png" alt-text="呼び出し時の回転テーブルのデータを含むワークシート":::

## <a name="create-an-office-script"></a><span data-ttu-id="c7bac-118">Office スクリプトを作成する</span><span class="sxs-lookup"><span data-stu-id="c7bac-118">Create an Office Script</span></span>

1. <span data-ttu-id="c7bac-119">**[オートメーション]** タブに移動して **[すべてのスクリプト]** を選択します。</span><span class="sxs-lookup"><span data-stu-id="c7bac-119">Go to the **Automate** tab and select **All Scripts**.</span></span>

1. <span data-ttu-id="c7bac-120">**[新しいスクリプト]** を選択します。</span><span class="sxs-lookup"><span data-stu-id="c7bac-120">Select **New Script**.</span></span>

1. <span data-ttu-id="c7bac-121">スクリプトに **Get On-Call Person** という名前を付けます。</span><span class="sxs-lookup"><span data-stu-id="c7bac-121">Name the script **Get On-Call Person**.</span></span>

1. <span data-ttu-id="c7bac-122">これで空のスクリプトができました。</span><span class="sxs-lookup"><span data-stu-id="c7bac-122">You should now have an empty script.</span></span> <span data-ttu-id="c7bac-123">スクリプトを使用して、スプレッドシートからメール アドレスを取得します。</span><span class="sxs-lookup"><span data-stu-id="c7bac-123">We want to use the script to get an email address from the spreadsheet.</span></span> <span data-ttu-id="c7bac-124">文字列が返されるように、`main` を次のように変更します。</span><span class="sxs-lookup"><span data-stu-id="c7bac-124">Change `main` to return a string, like this:</span></span>

    ```TypeScript
    function main(workbook: ExcelScript.Workbook) : string {
    }
    ```

1. <span data-ttu-id="c7bac-125">続いて、テーブルからすべてのデータを取得する必要があります。</span><span class="sxs-lookup"><span data-stu-id="c7bac-125">Next, we need to get all the data from the table.</span></span> <span data-ttu-id="c7bac-126">それにより、スクリプトを使用して各行を確認できます。</span><span class="sxs-lookup"><span data-stu-id="c7bac-126">That lets us look through each row with the script.</span></span> <span data-ttu-id="c7bac-127">`main` 関数に次のコードを追加します。</span><span class="sxs-lookup"><span data-stu-id="c7bac-127">Add the following code inside the `main` function.</span></span>

    ```TypeScript
    // Get the H1 worksheet.
    let worksheet = workbook.getWorksheet("H1");

    // Get the first (and only) table in the worksheet.
    let table = worksheet.getTables()[0];

    // Get the data from the table.
    let tableValues = table.getRangeBetweenHeaderAndTotal().getValues();
    ```

1. <span data-ttu-id="c7bac-128">テーブル内の日付は、[Excel の日付システム](https://support.microsoft.com/office/date-systems-in-excel-e7fe7167-48a9-4b96-bb53-5612a800b487)を使用して保存されます。</span><span class="sxs-lookup"><span data-stu-id="c7bac-128">The dates in the table are stored using [Excel's date serial number](https://support.microsoft.com/office/date-systems-in-excel-e7fe7167-48a9-4b96-bb53-5612a800b487).</span></span> <span data-ttu-id="c7bac-129">これらの日付は、比較できるように JavaScript の日付に変換する必要があります。</span><span class="sxs-lookup"><span data-stu-id="c7bac-129">We need to convert those dates to JavaScript dates in order to compare them.</span></span> <span data-ttu-id="c7bac-130">ヘルパー関数をスクリプトに追加します。</span><span class="sxs-lookup"><span data-stu-id="c7bac-130">We'll add a helper function to our script.</span></span> <span data-ttu-id="c7bac-131">`main` 関数の外に次のコードを追加します。</span><span class="sxs-lookup"><span data-stu-id="c7bac-131">Add the following code outside of the `main` function:</span></span>

    ```TypeScript
    // Convert the Excel date to a JavaScript Date object.
    function convertDate(excelDateValue: number) {
        let javaScriptDate = new Date(Math.round((excelDateValue - 25569) * 86400 * 1000));
        return javaScriptDate;
    }
    ```

1. <span data-ttu-id="c7bac-132">次に、現在誰が呼び出し期間中かを把握する必要があります。</span><span class="sxs-lookup"><span data-stu-id="c7bac-132">Now, we need to figure out which person is on call right now.</span></span> <span data-ttu-id="c7bac-133">それらの行では、開始日と終了日の間に現在の日付が含まれています。</span><span class="sxs-lookup"><span data-stu-id="c7bac-133">Their row will have a start and end date surrounding the current date.</span></span> <span data-ttu-id="c7bac-134">ここでは、一度に 1 人だけが呼び出し期間であると想定してスクリプトを作成します。</span><span class="sxs-lookup"><span data-stu-id="c7bac-134">We'll write the script to assume only one person is on call at a time.</span></span> <span data-ttu-id="c7bac-135">スクリプトで配列を返して複数の値を処理することもできますが、現時点では、最初に一致するメール アドレスを返すようにします。</span><span class="sxs-lookup"><span data-stu-id="c7bac-135">Scripts can return arrays to handle multiple values, but for now we'll return the first matching email address.</span></span> <span data-ttu-id="c7bac-136">次の関数を `main` 関数の最後に追加します。</span><span class="sxs-lookup"><span data-stu-id="c7bac-136">Add the following code to the end of the `main` function.</span></span>

    ```TypeScript
    // Look for the first row where today's date is between the row's start and end dates.
    let currentDate = new Date();
    for (let row = 0; row < tableValues.length; row++) {
        let startDate = convertDate(tableValues[row][2] as number);
        let endDate = convertDate(tableValues[row][3] as number);
        if (startDate <= currentDate && endDate >= currentDate) {
            // Return the first matching email address.
            return tableValues[row][1].toString();
        }
    }
    ```

1. <span data-ttu-id="c7bac-137">最終的なスクリプトは、次のようになります。</span><span class="sxs-lookup"><span data-stu-id="c7bac-137">The final script should look like this:</span></span>

    ```TypeScript
    function main(workbook: ExcelScript.Workbook) : string {
        // Get the H1 worksheet.
        let worksheet = workbook.getWorksheet("H1");

        // Get the first (and only) table in the worksheet.
        let table = worksheet.getTables()[0];
    
        // Get the data from the table.
        let tableValues = table.getRangeBetweenHeaderAndTotal().getValues();
    
        // Look for the first row where today's date is between the row's start and end dates.
        let currentDate = new Date();
        for (let row = 0; row < tableValues.length; row++) {
            let startDate = convertDate(tableValues[row][2] as number);
            let endDate = convertDate(tableValues[row][3] as number);
            if (startDate <= currentDate && endDate >= currentDate) {
                // Return the first matching email address.
                return tableValues[row][1].toString();
            }
        }
    }

    // Convert the Excel date to a JavaScript Date object.
    function convertDate(excelDateValue: number) {
        let javaScriptDate = new Date(Math.round((excelDateValue - 25569) * 86400 * 1000));
        return javaScriptDate;
    }
    ```

## <a name="create-an-automated-workflow-with-power-automate"></a><span data-ttu-id="c7bac-138">Power Automate を使用して自動化されたワークフローを作成する</span><span class="sxs-lookup"><span data-stu-id="c7bac-138">Create an automated workflow with Power Automate</span></span>

1. <span data-ttu-id="c7bac-139">[「Power Automate のサイト」](https://flow.microsoft.com)にサインインします。</span><span class="sxs-lookup"><span data-stu-id="c7bac-139">Sign in to the [Power Automate site](https://flow.microsoft.com).</span></span>

1. <span data-ttu-id="c7bac-140">画面の左側に表示されるメニューで、**[作成]** を押します。</span><span class="sxs-lookup"><span data-stu-id="c7bac-140">In the menu that's displayed on the left side of the screen, press **Create**.</span></span> <span data-ttu-id="c7bac-141">これにより、新しいワークフローを作成する方法の一覧を表示できます。</span><span class="sxs-lookup"><span data-stu-id="c7bac-141">This brings you to list of ways to create new workflows.</span></span>

    :::image type="content" source="../images/power-automate-tutorial-1.png" alt-text="Power Automate の [作成] ボタン":::

1. <span data-ttu-id="c7bac-143">**[空白から開始]** セクションで **[スケジュール済みクラウド フロー]** を選択します。</span><span class="sxs-lookup"><span data-stu-id="c7bac-143">Under the **Start from blank** section, select **Scheduled cloud flow**.</span></span>

    :::image type="content" source="../images/power-automate-return-tutorial-2.png" alt-text="Power Automate の [スケジュール済みクラウド フロー] ボタン":::

1. <span data-ttu-id="c7bac-145">続いて、このフローのスケジュールを設定します。</span><span class="sxs-lookup"><span data-stu-id="c7bac-145">Now we need to set the schedule for this flow.</span></span> <span data-ttu-id="c7bac-146">使用しているスプレッドシートには、2021 年前半の毎週月曜日から始まる新しい呼び出し期間の割り当てが含まれています。</span><span class="sxs-lookup"><span data-stu-id="c7bac-146">Our spreadsheet has a new on-call assignment starting every Monday in the first half of 2021.</span></span> <span data-ttu-id="c7bac-147">月曜日の朝一番に実行するようにフローを設定します。</span><span class="sxs-lookup"><span data-stu-id="c7bac-147">Let's set the flow to run first thing Monday mornings.</span></span> <span data-ttu-id="c7bac-148">次のオプションを使用して、毎週月曜日に実行するようにフローを構成します。</span><span class="sxs-lookup"><span data-stu-id="c7bac-148">Use the following options to configure the flow to run on Monday each week.</span></span>

    - <span data-ttu-id="c7bac-149">**フロー名**: Notify On-Call Person</span><span class="sxs-lookup"><span data-stu-id="c7bac-149">**Flow name**: Notify On-Call Person</span></span>
    - <span data-ttu-id="c7bac-150">**開始**: 21/1/4 時間 1:00 AM</span><span class="sxs-lookup"><span data-stu-id="c7bac-150">**Starting**: 1/4/21 at 1:00am</span></span>
    - <span data-ttu-id="c7bac-151">**繰り返し間隔**: 1 週</span><span class="sxs-lookup"><span data-stu-id="c7bac-151">**Repeat every**: 1 Week</span></span>
    - <span data-ttu-id="c7bac-152">**設定曜日**: 月</span><span class="sxs-lookup"><span data-stu-id="c7bac-152">**On these days**: M</span></span>

    :::image type="content" source="../images/power-automate-return-tutorial-3.png" alt-text="オプションが表示された Power Automate の [スケジュールされたクラウド フローを作成する] ダイアログ。オプションには、フロー名、開始時刻、繰り返しの頻度、フローを実行する曜日が含まれます":::

1. <span data-ttu-id="c7bac-154">**[作成]** を押します。</span><span class="sxs-lookup"><span data-stu-id="c7bac-154">Press **Create**.</span></span>

1. <span data-ttu-id="c7bac-155">**[新しいステップ]** を押します。</span><span class="sxs-lookup"><span data-stu-id="c7bac-155">Press **New step**.</span></span>

1. <span data-ttu-id="c7bac-156">**[標準]** タブを選択し、**Excel Online (ビジネス)** を選択します。</span><span class="sxs-lookup"><span data-stu-id="c7bac-156">Select the **Standard** tab, then select **Excel Online (Business)**.</span></span>

    :::image type="content" source="../images/power-automate-tutorial-4.png" alt-text="Power Automate の [Excel Online (Business)] オプション":::

1. <span data-ttu-id="c7bac-158">**[アクション]** で、**[スクリプトの実行]** を選択します。</span><span class="sxs-lookup"><span data-stu-id="c7bac-158">Under **Actions**, select **Run script**.</span></span>

    :::image type="content" source="../images/power-automate-tutorial-5.png" alt-text="Power Automate の [スクリプトの実行] アクションのオプション":::

1. <span data-ttu-id="c7bac-160">次に、フロー ステップで使用するブックとスクリプトを選択します。</span><span class="sxs-lookup"><span data-stu-id="c7bac-160">Next, you'll select the workbook and script to use in the flow step.</span></span> <span data-ttu-id="c7bac-161">自分の OneDrive で作成したブック **on-call-rotation.xlsx** を使用します。</span><span class="sxs-lookup"><span data-stu-id="c7bac-161">Use the **on-call-rotation.xlsx** workbook you created in your OneDrive.</span></span> <span data-ttu-id="c7bac-162">**スクリプトの実行** コネクタには、次の設定を指定します。</span><span class="sxs-lookup"><span data-stu-id="c7bac-162">Specify the following settings for the **Run script** connector:</span></span>

    - <span data-ttu-id="c7bac-163">**場所**: OneDrive for Business</span><span class="sxs-lookup"><span data-stu-id="c7bac-163">**Location**: OneDrive for Business</span></span>
    - <span data-ttu-id="c7bac-164">**ドキュメント ライブラリ**: OneDrive</span><span class="sxs-lookup"><span data-stu-id="c7bac-164">**Document Library**: OneDrive</span></span>
    - <span data-ttu-id="c7bac-165">**ファイル**: on-call-rotation.xlsx *(ファイル ブラウザーを使用して選択されています)*</span><span class="sxs-lookup"><span data-stu-id="c7bac-165">**File**: on-call-rotation.xlsx *(Chosen through the file browser)*</span></span>
    - <span data-ttu-id="c7bac-166">**スクリプト**: Get On-Call Person</span><span class="sxs-lookup"><span data-stu-id="c7bac-166">**Script**: Get On-Call Person</span></span>

    :::image type="content" source="../images/power-automate-return-tutorial-4.png" alt-text="スクリプトを実行するための Power Automate コネクタの設定":::

1. <span data-ttu-id="c7bac-168">**[新しいステップ]** を押します。</span><span class="sxs-lookup"><span data-stu-id="c7bac-168">Press **New step**.</span></span>

1. <span data-ttu-id="c7bac-169">リマインダー メールを送信してフローを終了します。</span><span class="sxs-lookup"><span data-stu-id="c7bac-169">We'll end the flow by sending the reminder email.</span></span> <span data-ttu-id="c7bac-170">コネクタの検索バーを使用して、**[メールの送信 (V2)]** を選択します。</span><span class="sxs-lookup"><span data-stu-id="c7bac-170">Select **Send an email (V2)** by using the connector's search bar.</span></span> <span data-ttu-id="c7bac-171">スクリプトによって返されるメール アドレスを追加するために、**動的なコンテンツの追加** コントロールを使用します。</span><span class="sxs-lookup"><span data-stu-id="c7bac-171">Use the **Add dynamic content** control to add the email address returned by the script.</span></span> <span data-ttu-id="c7bac-172">これは、**result** というラベル付きの Excel アイコンで示されます。</span><span class="sxs-lookup"><span data-stu-id="c7bac-172">This will be labelled **result** with the Excel icon next to it.</span></span> <span data-ttu-id="c7bac-173">件名、本文は自由に入力できます。</span><span class="sxs-lookup"><span data-stu-id="c7bac-173">You can provide whatever subject and body text you'd like.</span></span>

    :::image type="content" source="../images/power-automate-return-tutorial-5.png" alt-text="メールを送信するための Power Automate Outlook コネクタの設定。オプションには、送信するファイル、メールの件名、メールの本文、および詳細オプションが含まれます":::

    > [!NOTE]
    > <span data-ttu-id="c7bac-p111">このチュートリアルでは、Outlook を使用します。代わりに、お好きなメール サービスを自由に使用することもできますが、一部のオプションは異なる場合があります。</span><span class="sxs-lookup"><span data-stu-id="c7bac-p111">This tutorial uses Outlook. Feel free to use your preferred email service instead, though some options may be different.</span></span>

1. <span data-ttu-id="c7bac-177">**[保存]** を押します。</span><span class="sxs-lookup"><span data-stu-id="c7bac-177">Press **Save**.</span></span>

## <a name="test-the-script-in-power-automate"></a><span data-ttu-id="c7bac-178">Power Automate でスクリプトをテストする</span><span class="sxs-lookup"><span data-stu-id="c7bac-178">Test the script in Power Automate</span></span>

<span data-ttu-id="c7bac-179">作成したフローは毎週月曜日に実行されます。</span><span class="sxs-lookup"><span data-stu-id="c7bac-179">Your flow will run every Monday morning.</span></span> <span data-ttu-id="c7bac-180">画面の右上隅にある **[テスト]** ボタンを押すと、スクリプトをテストできます。</span><span class="sxs-lookup"><span data-stu-id="c7bac-180">You can test the script now by pressing the **Test** button in the upper-right corner of the screen.</span></span> <span data-ttu-id="c7bac-181">**[手動]** を選択し、**[テストの実行]** を押して直ちにフローを実行し、動作をテストします。</span><span class="sxs-lookup"><span data-stu-id="c7bac-181">Select **Manually** and press **Run Test** to run the flow now and test the behavior.</span></span> <span data-ttu-id="c7bac-182">続行するには、Excel と Outlook にアクセス許可を付与する必要がある場合があります。</span><span class="sxs-lookup"><span data-stu-id="c7bac-182">You may need to grant permissions to Excel and Outlook to continue.</span></span>

:::image type="content" source="../images/power-automate-return-tutorial-6.png" alt-text="Power Automate の [テスト] ボタン":::

> [!TIP]
> <span data-ttu-id="c7bac-184">フローでメールを送信できない場合は、スプレッドシートで、有効なメールが現在の日付範囲用としてテーブルの先頭にリストされていることを再確認してください。</span><span class="sxs-lookup"><span data-stu-id="c7bac-184">If your flow fails to send an email, double-check in the spreadsheet that a valid email is listed for the current date range at the top of the table.</span></span>

## <a name="next-steps"></a><span data-ttu-id="c7bac-185">次の手順</span><span class="sxs-lookup"><span data-stu-id="c7bac-185">Next steps</span></span>

<span data-ttu-id="c7bac-186">Office スクリプトを Power Automate に接続する方法に関する詳細については、 [「Power Automate で Office スクリプトを実行する」](../develop/power-automate-integration.md)を参照してください。</span><span class="sxs-lookup"><span data-stu-id="c7bac-186">Visit [Run Office Scripts with Power Automate](../develop/power-automate-integration.md) to learn more about connecting Office Scripts with Power Automate.</span></span>

<span data-ttu-id="c7bac-187">[「自動タスク リマインダーのサンプル シナリオ」](../resources/scenarios/task-reminders.md)では、Office スクリプトと Power Automate を Teams アダプティブ カードと組み合わせる方法についても説明します。</span><span class="sxs-lookup"><span data-stu-id="c7bac-187">You can also check out the [Automated task reminders sample scenario](../resources/scenarios/task-reminders.md) to learn how to combine Office Scripts and Power Automate with Teams Adaptive Cards.</span></span>
