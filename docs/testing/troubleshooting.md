---
title: Office スクリプトのトラブルシューティング
description: Office スクリプトのデバッグに関するヒントとテクニック、およびヘルプ リソース。
ms.date: 05/17/2021
localization_priority: Normal
ms.openlocfilehash: ff0ac1e63084c7c541d2a4925f1f011d16fa4992
ms.sourcegitcommit: 4687693f02fc90a57ba30c461f35046e02e6f5fb
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 05/19/2021
ms.locfileid: "52545558"
---
# <a name="troubleshoot-office-scripts"></a><span data-ttu-id="440d0-103">Office スクリプトのトラブルシューティング</span><span class="sxs-lookup"><span data-stu-id="440d0-103">Troubleshoot Office Scripts</span></span>

<span data-ttu-id="440d0-104">Officeスクリプトを開発する際に、間違いを犯す可能性があります。</span><span class="sxs-lookup"><span data-stu-id="440d0-104">As you develop Office Scripts, you may make mistakes.</span></span> <span data-ttu-id="440d0-105">大丈夫です。</span><span class="sxs-lookup"><span data-stu-id="440d0-105">It's okay.</span></span> <span data-ttu-id="440d0-106">問題を見つけ、スクリプトを完璧に動作させるために役立つツールがあります。</span><span class="sxs-lookup"><span data-stu-id="440d0-106">You have the tools to help find the problems and get your scripts working perfectly.</span></span>

## <a name="types-of-errors"></a><span data-ttu-id="440d0-107">エラーの種類</span><span class="sxs-lookup"><span data-stu-id="440d0-107">Types of errors</span></span>

<span data-ttu-id="440d0-108">Officeスクリプト エラーは、次の 2 つのカテゴリのいずれかに分類されます。</span><span class="sxs-lookup"><span data-stu-id="440d0-108">Office Scripts errors fall into one of two categories:</span></span>

* <span data-ttu-id="440d0-109">コンパイル時のエラーまたは警告</span><span class="sxs-lookup"><span data-stu-id="440d0-109">Compile-time errors or warnings</span></span>
* <span data-ttu-id="440d0-110">ランタイム エラー</span><span class="sxs-lookup"><span data-stu-id="440d0-110">Runtime errors</span></span>

### <a name="compile-time-errors"></a><span data-ttu-id="440d0-111">コンパイル時エラー</span><span class="sxs-lookup"><span data-stu-id="440d0-111">Compile-time errors</span></span>

<span data-ttu-id="440d0-112">コンパイル時のエラーと警告は、最初はコード エディターに表示されます。</span><span class="sxs-lookup"><span data-stu-id="440d0-112">Compile-time errors and warnings are initially shown in the Code Editor.</span></span> <span data-ttu-id="440d0-113">これらは、エディタで赤い波線で示されます。</span><span class="sxs-lookup"><span data-stu-id="440d0-113">These are shown by the wavy red underlines in the editor.</span></span> <span data-ttu-id="440d0-114">また、[コード エディタ] 作業ウィンドウの下部にある [ **問題** ] タブにも表示されます。</span><span class="sxs-lookup"><span data-stu-id="440d0-114">They are also displayed under the **Problems** tab at the bottom of the Code Editor task pane.</span></span> <span data-ttu-id="440d0-115">エラーを選択すると、問題の詳細が表示され、解決策が提案されます。</span><span class="sxs-lookup"><span data-stu-id="440d0-115">Selecting the error will give more details about the problem and suggest solutions.</span></span> <span data-ttu-id="440d0-116">コンパイル時エラーは、スクリプトを実行する前に対処する必要があります。</span><span class="sxs-lookup"><span data-stu-id="440d0-116">Compile-time errors should be addressed before running the script.</span></span>

:::image type="content" source="../images/explicit-any-editor-message.png" alt-text="コード エディターのホバー テキストに表示されるコンパイラ エラー":::

<span data-ttu-id="440d0-118">オレンジ色の警告の下線や灰色の情報メッセージが表示されることもあります。</span><span class="sxs-lookup"><span data-stu-id="440d0-118">You may also see orange warning underlines and grey informational messages.</span></span> <span data-ttu-id="440d0-119">これらは、スクリプトが意図しない影響を及ぼす可能性があるパフォーマンスの提案やその他の可能性を示します。</span><span class="sxs-lookup"><span data-stu-id="440d0-119">These indicate performance suggestions or other possibilities where the script may have unintentional effects.</span></span> <span data-ttu-id="440d0-120">このような警告は、それらを却下する前に綿密に検討する必要があります。</span><span class="sxs-lookup"><span data-stu-id="440d0-120">Such warnings should be examined closely before dismissing them.</span></span>

### <a name="runtime-errors"></a><span data-ttu-id="440d0-121">ランタイム エラー</span><span class="sxs-lookup"><span data-stu-id="440d0-121">Runtime errors</span></span>

<span data-ttu-id="440d0-122">ランタイム エラーは、スクリプトのロジックの問題が原因で発生します。</span><span class="sxs-lookup"><span data-stu-id="440d0-122">Runtime errors happen because of logic issues in the script.</span></span> <span data-ttu-id="440d0-123">スクリプトで使用されているオブジェクトがブックに含まれなかったり、テーブルの書式が予想と異なる場合、スクリプトの要件と現在のブックとの間に若干の不一致が生じ、その他の若干の相違が生じます。</span><span class="sxs-lookup"><span data-stu-id="440d0-123">This could be because an object used in the script isn't in the workbook, a table is formatted differently than anticipated, or some other slight discrepancy between the script's requirements and the current workbook.</span></span> <span data-ttu-id="440d0-124">次のスクリプトは、"TestSheet" という名前のワークシートが存在しない場合にエラーを生成します。</span><span class="sxs-lookup"><span data-stu-id="440d0-124">The following script generates an error when a worksheet named "TestSheet" is not present.</span></span>

```TypeScript
function main(workbook: ExcelScript.Workbook) {
  let mySheet = workbook.getWorksheet('TestSheet');

  // This will throw an error if there is no "TestSheet".
  mySheet.getRange("A1");
}
```

### <a name="console-messages"></a><span data-ttu-id="440d0-125">コンソール メッセージ</span><span class="sxs-lookup"><span data-stu-id="440d0-125">Console messages</span></span>

<span data-ttu-id="440d0-126">コンパイル時エラーと実行時エラーの両方が、スクリプトの実行時にコンソールにエラーメッセージを表示します。</span><span class="sxs-lookup"><span data-stu-id="440d0-126">Both compile-time and runtime errors display error messages in the console when a script runs.</span></span> <span data-ttu-id="440d0-127">問題が発生した行番号を指定します。</span><span class="sxs-lookup"><span data-stu-id="440d0-127">They give a line number where the problem was encountered.</span></span> <span data-ttu-id="440d0-128">問題の根本原因は、コンソールに示されているものとは異なるコード行である可能性があります。</span><span class="sxs-lookup"><span data-stu-id="440d0-128">Keep in mind that the root cause of any issue may be a different line of code than what is indicated in the console.</span></span>

<span data-ttu-id="440d0-129">次の図は、[明示的 `any` ](../develop/typescript-restrictions.md)なコンパイラ エラーのコンソール出力を示しています。</span><span class="sxs-lookup"><span data-stu-id="440d0-129">The following image shows the console output for the [explicit `any`](../develop/typescript-restrictions.md) compiler error.</span></span> <span data-ttu-id="440d0-130">`[5, 16]`エラー文字列の先頭にあるテキストを書き留めます。</span><span class="sxs-lookup"><span data-stu-id="440d0-130">Note the text `[5, 16]` at the beginning of the error string.</span></span> <span data-ttu-id="440d0-131">これは、エラーが 5 行目で、文字 16 から始まる場合を示します。</span><span class="sxs-lookup"><span data-stu-id="440d0-131">This indicates the error is on line 5, starting at character 16.</span></span>
:::image type="content" source="../images/explicit-any-error-message.png" alt-text="コード エディターコンソールで明示的な 'any' エラー メッセージが表示される":::

<span data-ttu-id="440d0-133">次の図は、ランタイム エラーのコンソール出力を示しています。</span><span class="sxs-lookup"><span data-stu-id="440d0-133">The follow image shows the console output for a runtime error.</span></span> <span data-ttu-id="440d0-134">ここでは、既存のワークシートの名前を持つワークシートを追加しようとします。</span><span class="sxs-lookup"><span data-stu-id="440d0-134">Here, the script tries to add a worksheet with a the name of an existing worksheet.</span></span> <span data-ttu-id="440d0-135">ここでも、エラーの前にある 「行 2」に注意して、調査する行を示します。</span><span class="sxs-lookup"><span data-stu-id="440d0-135">Again, note the "Line 2" preceding the error to show which line to investigate.</span></span>
:::image type="content" source="../images/runtime-error-console.png" alt-text="'addWorksheet' 呼び出しからのエラーを表示するコード エディター コンソール":::

## <a name="console-logs"></a><span data-ttu-id="440d0-137">コンソールログ</span><span class="sxs-lookup"><span data-stu-id="440d0-137">Console logs</span></span>

<span data-ttu-id="440d0-138">ステートメントを使用してメッセージを画面に出力 `console.log` します。</span><span class="sxs-lookup"><span data-stu-id="440d0-138">Print messages to the screen with the `console.log` statement.</span></span> <span data-ttu-id="440d0-139">これらのログには、変数の現在の値や、トリガーされるコードパスを示すことができます。</span><span class="sxs-lookup"><span data-stu-id="440d0-139">These logs can show you the current value of variables or which code paths are being triggered.</span></span> <span data-ttu-id="440d0-140">これを行うには、 `console.log` 任意のオブジェクトをパラメーターとして呼び出します。</span><span class="sxs-lookup"><span data-stu-id="440d0-140">To do this, call `console.log` with any object as a parameter.</span></span> <span data-ttu-id="440d0-141">通常、 `string` コンソールで読み取るのが最も簡単なタイプは a です。</span><span class="sxs-lookup"><span data-stu-id="440d0-141">Usually, a `string` is the easiest type to read in the console.</span></span>

```TypeScript
console.log("Logging myRange's address.");
console.log(myRange.getAddress());
```

<span data-ttu-id="440d0-142">渡された文字列 `console.log` は、作業ウィンドウの下部にあるコード エディタのログ コンソールに表示されます。</span><span class="sxs-lookup"><span data-stu-id="440d0-142">Strings passed to `console.log` are displayed in the Code Editor's logging console, at the bottom of the task pane.</span></span> <span data-ttu-id="440d0-143">ログは **[出力** ]タブにありますが、ログが書き込まれると自動的にフォーカスが移動します。</span><span class="sxs-lookup"><span data-stu-id="440d0-143">Logs are found on the **Output** tab, though the tab automatically gains focus when a log is written.</span></span>

<span data-ttu-id="440d0-144">ログはブックには影響しません。</span><span class="sxs-lookup"><span data-stu-id="440d0-144">Logs do not affect the workbook.</span></span>

## <a name="automate-tab-not-appearing-or-office-scripts-unavailable"></a><span data-ttu-id="440d0-145">[自動化] タブが表示されないか、スクリプトが使用できないOffice</span><span class="sxs-lookup"><span data-stu-id="440d0-145">Automate tab not appearing or Office Scripts unavailable</span></span>

<span data-ttu-id="440d0-146">次の手順は、[**自動化**] タブがExcel on the webに表示されない問題のトラブルシューティングに役立ちます。</span><span class="sxs-lookup"><span data-stu-id="440d0-146">The following steps should help troubleshoot any problems related to the **Automate** tab not appearing in Excel on the web.</span></span>

1. <span data-ttu-id="440d0-147">[Microsoft 365 ライセンスに [Office スクリプト] が含まれていることを確認](../overview/excel.md#requirements)します。</span><span class="sxs-lookup"><span data-stu-id="440d0-147">[Make sure your Microsoft 365 license includes Office Scripts](../overview/excel.md#requirements).</span></span>
1. <span data-ttu-id="440d0-148">[お使いのブラウザがサポートされていることを確認](platform-limits.md#browser-support)します。</span><span class="sxs-lookup"><span data-stu-id="440d0-148">[Check that your browser is supported](platform-limits.md#browser-support).</span></span>
1. <span data-ttu-id="440d0-149">[サードパーティの Cookie が有効になっていることを確認](platform-limits.md#third-party-cookies)します。</span><span class="sxs-lookup"><span data-stu-id="440d0-149">[Ensure third-party cookies are enabled](platform-limits.md#third-party-cookies).</span></span>
1. <span data-ttu-id="440d0-150">[管理者がMicrosoft 365管理センターのOfficeスクリプトを無効にしていないことを確認](/microsoft-365/admin/manage/manage-office-scripts-settings)します。</span><span class="sxs-lookup"><span data-stu-id="440d0-150">[Ensure that your admin has not disabled Office Scripts in the Microsoft 365 admin center](/microsoft-365/admin/manage/manage-office-scripts-settings).</span></span>

[!INCLUDE [Teams support note](../includes/teams-support-note.md)]

## <a name="troubleshoot-scripts-in-power-automate"></a><span data-ttu-id="440d0-151">Power Automateのスクリプトのトラブルシューティング</span><span class="sxs-lookup"><span data-stu-id="440d0-151">Troubleshoot scripts in Power Automate</span></span>

<span data-ttu-id="440d0-152">Power Automateを通じてスクリプトを実行する方法については、「 [Power Automate で実行されているOfficeスクリプトのトラブルシューティング](power-automate-troubleshooting.md)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="440d0-152">For information specific to running scripts through Power Automate, see [Troubleshoot Office Scripts running in Power Automate](power-automate-troubleshooting.md).</span></span>

## <a name="help-resources"></a><span data-ttu-id="440d0-153">ヘルプ リソース</span><span class="sxs-lookup"><span data-stu-id="440d0-153">Help resources</span></span>

<span data-ttu-id="440d0-154">[スタックオーバーフロー](https://stackoverflow.com/questions/tagged/office-scripts) は、コーディングの問題を支援する開発者のコミュニティです。</span><span class="sxs-lookup"><span data-stu-id="440d0-154">[Stack Overflow](https://stackoverflow.com/questions/tagged/office-scripts) is a community of developers willing to help with coding problems.</span></span> <span data-ttu-id="440d0-155">スタックオーバーフロー検索を迅速に行うと、問題の解決策を見つけることができます。</span><span class="sxs-lookup"><span data-stu-id="440d0-155">Often, you'll be able to find the solution to your problem through a quick Stack Overflow search.</span></span> <span data-ttu-id="440d0-156">それ以下の場合は、質問をして「オフィススクリプト」タグでタグ付けします。</span><span class="sxs-lookup"><span data-stu-id="440d0-156">If not, ask your question and tag it with the "office-scripts" tag.</span></span> <span data-ttu-id="440d0-157">Office アドインではなく、Office *スクリプト* を作成していることを必ずお *知りください*。</span><span class="sxs-lookup"><span data-stu-id="440d0-157">Be sure to mention you're creating an Office *Script*, not an Office *Add-in*.</span></span>

<span data-ttu-id="440d0-158">Office JavaScript API で問題が発生した場合は[、OfficeDev/office-js](https://github.com/OfficeDev/office-js) GitHubリポジトリで問題を作成します。</span><span class="sxs-lookup"><span data-stu-id="440d0-158">If you encounter a problem with the Office JavaScript API, create an issue in the [OfficeDev/office-js](https://github.com/OfficeDev/office-js) GitHub repository.</span></span> <span data-ttu-id="440d0-159">製品チームのメンバーは、問題に対応し、さらなる支援を提供します。</span><span class="sxs-lookup"><span data-stu-id="440d0-159">Members of the product team will respond to issues and provide further assistance.</span></span> <span data-ttu-id="440d0-160">**OfficeDev/office-js** リポジトリで問題を作成すると、製品チームが対処する必要がある Office JavaScript API ライブラリに問題が見つかったことが示されます。</span><span class="sxs-lookup"><span data-stu-id="440d0-160">Creating an issue in the **OfficeDev/office-js** repository indicates you have found a flaw in the Office JavaScript API library that the product team should address.</span></span>

<span data-ttu-id="440d0-161">アクション レコーダまたはエディタに問題がある場合は、Excelの **ヘルプ>フィードバック** ボタンからフィードバックを送信します。</span><span class="sxs-lookup"><span data-stu-id="440d0-161">If there is a problem with the Action Recorder or Editor, send feedback through the **Help > Feedback** button in Excel.</span></span>

## <a name="see-also"></a><span data-ttu-id="440d0-162">関連項目</span><span class="sxs-lookup"><span data-stu-id="440d0-162">See also</span></span>

- [<span data-ttu-id="440d0-163">Office スクリプトのベスト プラクティス</span><span class="sxs-lookup"><span data-stu-id="440d0-163">Best practices in Office Scripts</span></span>](../develop/best-practices.md)
- [<span data-ttu-id="440d0-164">Officeスクリプトを使用したプラットフォームの制限</span><span class="sxs-lookup"><span data-stu-id="440d0-164">Platform limits with Office Scripts</span></span>](platform-limits.md)
- [<span data-ttu-id="440d0-165">Officeスクリプトのパフォーマンスを向上させる</span><span class="sxs-lookup"><span data-stu-id="440d0-165">Improve the performance of your Office Scripts</span></span>](../develop/web-client-performance.md)
- [<span data-ttu-id="440d0-166">PowerAutomate で実行されているOffice スクリプトのトラブルシューティング</span><span class="sxs-lookup"><span data-stu-id="440d0-166">Troubleshoot Office Scripts running in PowerAutomate</span></span>](power-automate-troubleshooting.md)
- [<span data-ttu-id="440d0-167">Office スクリプトの効果を元に戻す</span><span class="sxs-lookup"><span data-stu-id="440d0-167">Undo the effects of Office Scripts</span></span>](undo.md)
