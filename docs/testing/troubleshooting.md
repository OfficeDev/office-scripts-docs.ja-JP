---
title: スクリプトOfficeトラブルシューティング
description: スクリプトのデバッグのヒントとOfficeヘルプ リソース。
ms.date: 05/17/2021
localization_priority: Normal
ms.openlocfilehash: 251ad72588422a86c52c81666164c2c4bd79bdb5
ms.sourcegitcommit: 4693c8f79428ec74695328275703af0ba1bfea8f
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 06/23/2021
ms.locfileid: "53074649"
---
# <a name="troubleshoot-office-scripts"></a><span data-ttu-id="a1662-103">スクリプトOfficeトラブルシューティング</span><span class="sxs-lookup"><span data-stu-id="a1662-103">Troubleshoot Office Scripts</span></span>

<span data-ttu-id="a1662-104">スクリプトを開発Office、間違いを犯す可能性があります。</span><span class="sxs-lookup"><span data-stu-id="a1662-104">As you develop Office Scripts, you may make mistakes.</span></span> <span data-ttu-id="a1662-105">大丈夫です。</span><span class="sxs-lookup"><span data-stu-id="a1662-105">It's okay.</span></span> <span data-ttu-id="a1662-106">問題を見つけてスクリプトを完全に機能するためのツールがあります。</span><span class="sxs-lookup"><span data-stu-id="a1662-106">You have the tools to help find the problems and get your scripts working perfectly.</span></span>

## <a name="types-of-errors"></a><span data-ttu-id="a1662-107">エラーの種類</span><span class="sxs-lookup"><span data-stu-id="a1662-107">Types of errors</span></span>

<span data-ttu-id="a1662-108">Officeスクリプトエラーは、次の 2 つのカテゴリに分類されます。</span><span class="sxs-lookup"><span data-stu-id="a1662-108">Office Scripts errors fall into one of two categories:</span></span>

* <span data-ttu-id="a1662-109">コンパイル時のエラーまたは警告</span><span class="sxs-lookup"><span data-stu-id="a1662-109">Compile-time errors or warnings</span></span>
* <span data-ttu-id="a1662-110">ランタイム エラー</span><span class="sxs-lookup"><span data-stu-id="a1662-110">Runtime errors</span></span>

### <a name="compile-time-errors"></a><span data-ttu-id="a1662-111">コンパイル時エラー</span><span class="sxs-lookup"><span data-stu-id="a1662-111">Compile-time errors</span></span>

<span data-ttu-id="a1662-112">コンパイル時のエラーと警告は、最初はコード エディターに表示されます。</span><span class="sxs-lookup"><span data-stu-id="a1662-112">Compile-time errors and warnings are initially shown in the Code Editor.</span></span> <span data-ttu-id="a1662-113">これらは、エディターの波状の赤い下線で表示されます。</span><span class="sxs-lookup"><span data-stu-id="a1662-113">These are shown by the wavy red underlines in the editor.</span></span> <span data-ttu-id="a1662-114">また、[コード エディター] 作業ウィンドウ **の** 下部にある [問題] タブにも表示されます。</span><span class="sxs-lookup"><span data-stu-id="a1662-114">They are also displayed under the **Problems** tab at the bottom of the Code Editor task pane.</span></span> <span data-ttu-id="a1662-115">エラーを選択すると、問題の詳細と解決策の提案が表示されます。</span><span class="sxs-lookup"><span data-stu-id="a1662-115">Selecting the error will give more details about the problem and suggest solutions.</span></span> <span data-ttu-id="a1662-116">コンパイル時のエラーは、スクリプトを実行する前に対処する必要があります。</span><span class="sxs-lookup"><span data-stu-id="a1662-116">Compile-time errors should be addressed before running the script.</span></span>

:::image type="content" source="../images/explicit-any-editor-message.png" alt-text="コード エディターのホバー テキストに表示されるコンパイラ エラー。":::

<span data-ttu-id="a1662-118">オレンジ色の警告の下線と灰色の情報メッセージが表示される場合があります。</span><span class="sxs-lookup"><span data-stu-id="a1662-118">You may also see orange warning underlines and grey informational messages.</span></span> <span data-ttu-id="a1662-119">これらは、スクリプトが意図しない効果を持つ可能性があるパフォーマンスの提案や他の可能性を示します。</span><span class="sxs-lookup"><span data-stu-id="a1662-119">These indicate performance suggestions or other possibilities where the script may have unintentional effects.</span></span> <span data-ttu-id="a1662-120">このような警告は、却下する前に注意して調べる必要があります。</span><span class="sxs-lookup"><span data-stu-id="a1662-120">Such warnings should be examined closely before dismissing them.</span></span>

### <a name="runtime-errors"></a><span data-ttu-id="a1662-121">ランタイム エラー</span><span class="sxs-lookup"><span data-stu-id="a1662-121">Runtime errors</span></span>

<span data-ttu-id="a1662-122">ランタイム エラーは、スクリプトのロジックの問題が原因で発生します。</span><span class="sxs-lookup"><span data-stu-id="a1662-122">Runtime errors happen because of logic issues in the script.</span></span> <span data-ttu-id="a1662-123">これは、スクリプトで使用されるオブジェクトがブック内に含めなかったり、テーブルの形式が予想と異なっている、またはスクリプトの要件と現在のブックの間に若干の不一致が生じていった場合に発生する可能性があります。</span><span class="sxs-lookup"><span data-stu-id="a1662-123">This could be because an object used in the script isn't in the workbook, a table is formatted differently than anticipated, or some other slight discrepancy between the script's requirements and the current workbook.</span></span> <span data-ttu-id="a1662-124">次のスクリプトは、"TestSheet" という名前のワークシートが存在しない場合にエラーを生成します。</span><span class="sxs-lookup"><span data-stu-id="a1662-124">The following script generates an error when a worksheet named "TestSheet" is not present.</span></span>

```TypeScript
function main(workbook: ExcelScript.Workbook) {
  let mySheet = workbook.getWorksheet('TestSheet');

  // This will throw an error if there is no "TestSheet".
  mySheet.getRange("A1");
}
```

### <a name="console-messages"></a><span data-ttu-id="a1662-125">コンソール メッセージ</span><span class="sxs-lookup"><span data-stu-id="a1662-125">Console messages</span></span>

<span data-ttu-id="a1662-126">コンパイル時と実行時の両方のエラーは、スクリプトの実行時にコンソールにエラー メッセージを表示します。</span><span class="sxs-lookup"><span data-stu-id="a1662-126">Both compile-time and runtime errors display error messages in the console when a script runs.</span></span> <span data-ttu-id="a1662-127">問題が発生した行番号を指定します。</span><span class="sxs-lookup"><span data-stu-id="a1662-127">They give a line number where the problem was encountered.</span></span> <span data-ttu-id="a1662-128">問題の根本原因は、コンソールで示されているコードとは異なるコード行である可能性があります。</span><span class="sxs-lookup"><span data-stu-id="a1662-128">Keep in mind that the root cause of any issue may be a different line of code than what is indicated in the console.</span></span>

<span data-ttu-id="a1662-129">次の図は、明示的なコンパイラ エラーのコンソール[出力を `any` ](../develop/typescript-restrictions.md)示しています。</span><span class="sxs-lookup"><span data-stu-id="a1662-129">The following image shows the console output for the [explicit `any`](../develop/typescript-restrictions.md) compiler error.</span></span> <span data-ttu-id="a1662-130">エラー文字列の `[5, 16]` 先頭にあるテキストに注意してください。</span><span class="sxs-lookup"><span data-stu-id="a1662-130">Note the text `[5, 16]` at the beginning of the error string.</span></span> <span data-ttu-id="a1662-131">これは、エラーが 5 行目で、文字 16 から始まるかどうかを示します。</span><span class="sxs-lookup"><span data-stu-id="a1662-131">This indicates the error is on line 5, starting at character 16.</span></span>
:::image type="content" source="../images/explicit-any-error-message.png" alt-text="明示的な 'any' エラー メッセージを表示するコード エディター コンソール。":::

<span data-ttu-id="a1662-133">次の図は、実行時エラーのコンソール出力を示しています。</span><span class="sxs-lookup"><span data-stu-id="a1662-133">The follow image shows the console output for a runtime error.</span></span> <span data-ttu-id="a1662-134">ここでは、既存のワークシートの名前を持つワークシートを追加します。</span><span class="sxs-lookup"><span data-stu-id="a1662-134">Here, the script tries to add a worksheet with a the name of an existing worksheet.</span></span> <span data-ttu-id="a1662-135">ここでも、エラーの前の "2 行目" に注意して、調査する行を表示します。</span><span class="sxs-lookup"><span data-stu-id="a1662-135">Again, note the "Line 2" preceding the error to show which line to investigate.</span></span>
:::image type="content" source="../images/runtime-error-console.png" alt-text="'addWorksheet' 呼び出しからのエラーを表示するコード エディター コンソール。":::

## <a name="console-logs"></a><span data-ttu-id="a1662-137">コンソール ログ</span><span class="sxs-lookup"><span data-stu-id="a1662-137">Console logs</span></span>

<span data-ttu-id="a1662-138">ステートメントを使用してメッセージを画面に印刷 `console.log` します。</span><span class="sxs-lookup"><span data-stu-id="a1662-138">Print messages to the screen with the `console.log` statement.</span></span> <span data-ttu-id="a1662-139">これらのログには、変数の現在の値、またはトリガーされるコード パスが表示されます。</span><span class="sxs-lookup"><span data-stu-id="a1662-139">These logs can show you the current value of variables or which code paths are being triggered.</span></span> <span data-ttu-id="a1662-140">これを行うには、任意 `console.log` のオブジェクトをパラメーターとして呼び出します。</span><span class="sxs-lookup"><span data-stu-id="a1662-140">To do this, call `console.log` with any object as a parameter.</span></span> <span data-ttu-id="a1662-141">通常、コンソール `string` で読み取りが最も簡単な型は a です。</span><span class="sxs-lookup"><span data-stu-id="a1662-141">Usually, a `string` is the easiest type to read in the console.</span></span>

```TypeScript
console.log("Logging myRange's address.");
console.log(myRange.getAddress());
```

<span data-ttu-id="a1662-142">渡された文字列は、作業ウィンドウの下部にあるコード エディターのログ コンソール `console.log` に表示されます。</span><span class="sxs-lookup"><span data-stu-id="a1662-142">Strings passed to `console.log` are displayed in the Code Editor's logging console, at the bottom of the task pane.</span></span> <span data-ttu-id="a1662-143">ログは [出力] タブ **にあります** が、ログの書き込み時にタブが自動的にフォーカスを取得します。</span><span class="sxs-lookup"><span data-stu-id="a1662-143">Logs are found on the **Output** tab, though the tab automatically gains focus when a log is written.</span></span>

<span data-ttu-id="a1662-144">ログはブックには影響を与えかねない。</span><span class="sxs-lookup"><span data-stu-id="a1662-144">Logs do not affect the workbook.</span></span>

## <a name="automate-tab-not-appearing-or-office-scripts-unavailable"></a><span data-ttu-id="a1662-145">[自動化] タブが表示されないか、Officeスクリプトが使用できない</span><span class="sxs-lookup"><span data-stu-id="a1662-145">Automate tab not appearing or Office Scripts unavailable</span></span>

<span data-ttu-id="a1662-146">次の手順は、[自動化] タブに関連する問題のトラブルシューティングに役立つExcel on the web。</span><span class="sxs-lookup"><span data-stu-id="a1662-146">The following steps should help troubleshoot any problems related to the **Automate** tab not appearing in Excel on the web.</span></span>

1. <span data-ttu-id="a1662-147">[ライセンスにスクリプトMicrosoft 365含Officeしてください](../overview/excel.md#requirements)。</span><span class="sxs-lookup"><span data-stu-id="a1662-147">[Make sure your Microsoft 365 license includes Office Scripts](../overview/excel.md#requirements).</span></span>
1. <span data-ttu-id="a1662-148">[ブラウザーがサポートされていないことを確認します](platform-limits.md#browser-support)。</span><span class="sxs-lookup"><span data-stu-id="a1662-148">[Check that your browser is supported](platform-limits.md#browser-support).</span></span>
1. <span data-ttu-id="a1662-149">[サードパーティの Cookie が有効になっているか確認します](platform-limits.md#third-party-cookies)。</span><span class="sxs-lookup"><span data-stu-id="a1662-149">[Ensure third-party cookies are enabled](platform-limits.md#third-party-cookies).</span></span>
1. <span data-ttu-id="a1662-150">[管理者が[スクリプト] の [スクリプトOffice無効にMicrosoft 365 管理センター。](/microsoft-365/admin/manage/manage-office-scripts-settings)</span><span class="sxs-lookup"><span data-stu-id="a1662-150">[Ensure that your admin has not disabled Office Scripts in the Microsoft 365 admin center](/microsoft-365/admin/manage/manage-office-scripts-settings).</span></span>

[!INCLUDE [Teams support note](../includes/teams-support-note.md)]

## <a name="troubleshoot-scripts-in-power-automate"></a><span data-ttu-id="a1662-151">スクリプトのトラブルシューティングを行Power Automate</span><span class="sxs-lookup"><span data-stu-id="a1662-151">Troubleshoot scripts in Power Automate</span></span>

<span data-ttu-id="a1662-152">スクリプトの実行に関する詳細については、「Power Automateで実行されているスクリプトOffice[トラブルシューティング」を参照Power Automate。](power-automate-troubleshooting.md)</span><span class="sxs-lookup"><span data-stu-id="a1662-152">For information specific to running scripts through Power Automate, see [Troubleshoot Office Scripts running in Power Automate](power-automate-troubleshooting.md).</span></span>

## <a name="help-resources"></a><span data-ttu-id="a1662-153">ヘルプ リソース</span><span class="sxs-lookup"><span data-stu-id="a1662-153">Help resources</span></span>

<span data-ttu-id="a1662-154">[スタック オーバーフロー](https://stackoverflow.com/questions/tagged/office-scripts) は、コーディングの問題を支援する開発者のコミュニティです。</span><span class="sxs-lookup"><span data-stu-id="a1662-154">[Stack Overflow](https://stackoverflow.com/questions/tagged/office-scripts) is a community of developers willing to help with coding problems.</span></span> <span data-ttu-id="a1662-155">多くの場合、スタック オーバーフローのクイック検索を使用して、問題の解決策を見つける可能性があります。</span><span class="sxs-lookup"><span data-stu-id="a1662-155">Often, you'll be able to find the solution to your problem through a quick Stack Overflow search.</span></span> <span data-ttu-id="a1662-156">そうでない場合は、質問をして"office-scripts" タグでタグ付けします。</span><span class="sxs-lookup"><span data-stu-id="a1662-156">If not, ask your question and tag it with the "office-scripts" tag.</span></span> <span data-ttu-id="a1662-157">アドインではなく、Office *スクリプト* を作成Office *してください*。</span><span class="sxs-lookup"><span data-stu-id="a1662-157">Be sure to mention you're creating an Office *Script*, not an Office *Add-in*.</span></span>

<span data-ttu-id="a1662-158">Office スクリプトの機能要求を送信するには、アイデアを [ユーザー音声[](https://excel.uservoice.com/forums/274580-excel-for-the-web?category_id=143439)] ページに投稿するか、機能要求が既に存在する場合は、そのユーザーに投票を追加します。</span><span class="sxs-lookup"><span data-stu-id="a1662-158">To submit a feature request for Office Scripts, post your idea to our [User Voice page](https://excel.uservoice.com/forums/274580-excel-for-the-web?category_id=143439), or if the feature request already exists there, add your vote for it.</span></span> <span data-ttu-id="a1662-159">[マクロ、スクリプト、アドイン] Excel for the webで要求をファイルしてください。</span><span class="sxs-lookup"><span data-stu-id="a1662-159">Be sure to file the request under Excel for the web in the "Macros, Scripts and Add-ins" category.</span></span>

<span data-ttu-id="a1662-160">アクション レコーダーまたはエディターに問題がある場合は、お知らせください。</span><span class="sxs-lookup"><span data-stu-id="a1662-160">If there is a problem with the Action Recorder or Editor, please let us know.</span></span> <span data-ttu-id="a1662-161">[コード エディター] 作業ウィンドウの **...** メニューで、[フィードバックの送信] ボタンを選択して問題を共有します。</span><span class="sxs-lookup"><span data-stu-id="a1662-161">In the Code Editor task pane's **...** menu, select the **Send feedback** button to share any issues.</span></span>

:::image type="content" source="../images/code-editor-feedback.png" alt-text="[フィードバックの送信] ボタンを含むコード エディターのオーバーフロー メニュー。":::

## <a name="see-also"></a><span data-ttu-id="a1662-163">関連項目</span><span class="sxs-lookup"><span data-stu-id="a1662-163">See also</span></span>

- [<span data-ttu-id="a1662-164">Office スクリプトでのベスト プラクティス</span><span class="sxs-lookup"><span data-stu-id="a1662-164">Best practices in Office Scripts</span></span>](../develop/best-practices.md)
- [<span data-ttu-id="a1662-165">スクリプトを使用したプラットフォームOffice制限</span><span class="sxs-lookup"><span data-stu-id="a1662-165">Platform limits with Office Scripts</span></span>](platform-limits.md)
- [<span data-ttu-id="a1662-166">スクリプトのパフォーマンスをOfficeする</span><span class="sxs-lookup"><span data-stu-id="a1662-166">Improve the performance of your Office Scripts</span></span>](../develop/web-client-performance.md)
- [<span data-ttu-id="a1662-167">PowerAutomate Office実行されているスクリプトのトラブルシューティング</span><span class="sxs-lookup"><span data-stu-id="a1662-167">Troubleshoot Office Scripts running in PowerAutomate</span></span>](power-automate-troubleshooting.md)
- [<span data-ttu-id="a1662-168">Office スクリプトの効果を元に戻す</span><span class="sxs-lookup"><span data-stu-id="a1662-168">Undo the effects of Office Scripts</span></span>](undo.md)
