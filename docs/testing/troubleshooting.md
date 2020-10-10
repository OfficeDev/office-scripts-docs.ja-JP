---
title: Office スクリプトのトラブルシューティング
description: Office スクリプトのヒントとテクニック、およびヘルプリソースをデバッグします。
ms.date: 10/08/2020
localization_priority: Normal
ms.openlocfilehash: 9b3f4be778f3cdb4711d1e41d4d68f87ebca8152
ms.sourcegitcommit: 42fa3b629c93930b4e73e9c4c01d0c8bdf6d7487
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 10/09/2020
ms.locfileid: "48411548"
---
# <a name="troubleshooting-office-scripts"></a><span data-ttu-id="01f45-103">Office スクリプトのトラブルシューティング</span><span class="sxs-lookup"><span data-stu-id="01f45-103">Troubleshooting Office Scripts</span></span>

<span data-ttu-id="01f45-104">Office スクリプトを開発する際には、誤りが発生することがあります。</span><span class="sxs-lookup"><span data-stu-id="01f45-104">As you develop Office Scripts, you may make mistakes.</span></span> <span data-ttu-id="01f45-105">大丈夫です。</span><span class="sxs-lookup"><span data-stu-id="01f45-105">It's okay.</span></span> <span data-ttu-id="01f45-106">問題を見つけてスクリプトを完全に動作させるためのツールが用意されています。</span><span class="sxs-lookup"><span data-stu-id="01f45-106">We have tools that help find the problems and get your scripts working perfectly.</span></span>

## <a name="console-logs"></a><span data-ttu-id="01f45-107">コンソールログ</span><span class="sxs-lookup"><span data-stu-id="01f45-107">Console logs</span></span>

<span data-ttu-id="01f45-108">トラブルシューティング中に、画面にメッセージを出力することもできます。</span><span class="sxs-lookup"><span data-stu-id="01f45-108">Sometimes while troubleshooting, you'll want to print messages to the screen.</span></span> <span data-ttu-id="01f45-109">これにより、変数の現在の値や、どのコードパスがトリガーされているかを確認できます。</span><span class="sxs-lookup"><span data-stu-id="01f45-109">These can show you the current value of variables or which code paths are being triggered.</span></span> <span data-ttu-id="01f45-110">これを行うには、テキストをコンソールに記録します。</span><span class="sxs-lookup"><span data-stu-id="01f45-110">To do this, log text to the console.</span></span>

```TypeScript
console.log("Logging myRange's address.");
console.log(myRange.getAddress());
```

<span data-ttu-id="01f45-111">に渡される文字列 `console.log` は、コードエディターのログコンソールに表示されます。</span><span class="sxs-lookup"><span data-stu-id="01f45-111">Strings passed to `console.log` will be displayed in the Code Editor's logging console.</span></span> <span data-ttu-id="01f45-112">コンソールをオンにするには、**省略記号**ボタンを押して [**ログ...** ] を選択します。</span><span class="sxs-lookup"><span data-stu-id="01f45-112">To turn on the console, press the **Ellipses** button and select **Logs...**</span></span>

<span data-ttu-id="01f45-113">ログがブックに影響を与えることはありません。</span><span class="sxs-lookup"><span data-stu-id="01f45-113">Logs do not affect the workbook.</span></span>

## <a name="error-messages"></a><span data-ttu-id="01f45-114">エラー メッセージ</span><span class="sxs-lookup"><span data-stu-id="01f45-114">Error messages</span></span>

<span data-ttu-id="01f45-115">Excel スクリプトで問題が発生すると、エラーが生成されます。</span><span class="sxs-lookup"><span data-stu-id="01f45-115">When your Excel Script encounters a problem running, it produces an error.</span></span> <span data-ttu-id="01f45-116">**ログを表示**するかどうかの確認を求めるポップアップが表示されます。</span><span class="sxs-lookup"><span data-stu-id="01f45-116">You'll see a prompt pop-up asking if you want to **View Logs**.</span></span> <span data-ttu-id="01f45-117">そのボタンを押してコンソールを開き、エラーを表示します。</span><span class="sxs-lookup"><span data-stu-id="01f45-117">Press that button to open the console and display any errors.</span></span>

## <a name="automate-tab-not-appearing-or-office-scripts-unavailable"></a><span data-ttu-id="01f45-118">タブが表示されない、または Office スクリプトが利用できないことを自動化する</span><span class="sxs-lookup"><span data-stu-id="01f45-118">Automate tab not appearing or Office Scripts unavailable</span></span>

<span data-ttu-id="01f45-119">次の手順は、web 上の Excel に表示されない [ **自動化** ] タブに関連する問題のトラブルシューティングに役立ちます。</span><span class="sxs-lookup"><span data-stu-id="01f45-119">The following steps should help troubleshoot any problems related to the **Automate** tab not appearing in Excel on the web.</span></span>

1. <span data-ttu-id="01f45-120">[Microsoft 365 ライセンスに Office スクリプトが含まれていることを確認して](../overview/excel.md#requirements)ください。</span><span class="sxs-lookup"><span data-stu-id="01f45-120">[Make sure your Microsoft 365 license includes Office Scripts](../overview/excel.md#requirements).</span></span>
1. <span data-ttu-id="01f45-121">[管理者に機能を有効にしてもらい](/microsoft-365/admin/manage/manage-office-scripts-settings)ます。</span><span class="sxs-lookup"><span data-stu-id="01f45-121">[Have your admin enable the feature](/microsoft-365/admin/manage/manage-office-scripts-settings).</span></span>
1. <span data-ttu-id="01f45-122">[ブラウザーがサポートされていることを確認して](platform-limits.md#browser-support)ください。</span><span class="sxs-lookup"><span data-stu-id="01f45-122">[Check that your browser is supported](platform-limits.md#browser-support).</span></span>
1. <span data-ttu-id="01f45-123">[サードパーティの cookie が有効になっていることを確認](platform-limits.md#third-party-cookies)します。</span><span class="sxs-lookup"><span data-stu-id="01f45-123">[Ensure third-party cookies are enabled](platform-limits.md#third-party-cookies).</span></span>

[!INCLUDE [Teams support note](../includes/teams-support-note.md)]

## <a name="help-resources"></a><span data-ttu-id="01f45-124">ヘルプリソース</span><span class="sxs-lookup"><span data-stu-id="01f45-124">Help resources</span></span>

<span data-ttu-id="01f45-125">[スタックオーバーフロー](https://stackoverflow.com/questions/tagged/office-scripts) は、コーディングに関する問題の解決に役立つ開発者のコミュニティです。</span><span class="sxs-lookup"><span data-stu-id="01f45-125">[Stack Overflow](https://stackoverflow.com/questions/tagged/office-scripts) is a community of developers willing to help with coding problems.</span></span> <span data-ttu-id="01f45-126">多くの場合、クイックスタックオーバーフロー検索を使用して、問題の解決策を見つけることができます。</span><span class="sxs-lookup"><span data-stu-id="01f45-126">Often, you'll be able to find the solution to your problem through a quick Stack Overflow search.</span></span> <span data-ttu-id="01f45-127">そうでない場合は、質問して、「office スクリプト」タグでタグを付けてください。</span><span class="sxs-lookup"><span data-stu-id="01f45-127">If not, ask your question and tag it with the "office-scripts" tag.</span></span> <span data-ttu-id="01f45-128">Office*アドイン*ではなく、office*スクリプト*を作成していることを必ずお伝えください。</span><span class="sxs-lookup"><span data-stu-id="01f45-128">Be sure to mention you're creating an Office *Script*, not an Office *Add-in*.</span></span>

<span data-ttu-id="01f45-129">Office JavaScript API で問題が発生した場合は、 [Officedev/Office/js](https://github.com/OfficeDev/office-js) GitHub リポジトリに問題を作成します。</span><span class="sxs-lookup"><span data-stu-id="01f45-129">If you encounter a problem with the Office JavaScript API, create an issue in the [OfficeDev/office-js](https://github.com/OfficeDev/office-js) GitHub repository.</span></span> <span data-ttu-id="01f45-130">製品チームのメンバーは問題に対応し、さらに支援を提供します。</span><span class="sxs-lookup"><span data-stu-id="01f45-130">Members of the product team will respond to issues and provide further assistance.</span></span> <span data-ttu-id="01f45-131">**Officedev/office-js**リポジトリで問題を発生させることは、製品チームが対処する必要のある OFFICE JavaScript API ライブラリに問題が見つかったことを示しています。</span><span class="sxs-lookup"><span data-stu-id="01f45-131">Creating an issue in the **OfficeDev/office-js** repository indicates you have found a flaw in the Office JavaScript API library that the product team should address.</span></span>

<span data-ttu-id="01f45-132">操作レコーダーまたは Editor に問題がある場合は、Excel の **ヘルプ > フィードバック** ボタンを使用してフィードバックを送信してください。</span><span class="sxs-lookup"><span data-stu-id="01f45-132">If there is a problem with the Action Recorder or Editor, send feedback through the **Help > Feedback** button in Excel.</span></span>

## <a name="see-also"></a><span data-ttu-id="01f45-133">関連項目</span><span class="sxs-lookup"><span data-stu-id="01f45-133">See also</span></span>

- [<span data-ttu-id="01f45-134">Excel on the web の Office スクリプト</span><span class="sxs-lookup"><span data-stu-id="01f45-134">Office Scripts in Excel on the web</span></span>](../overview/excel.md)
- [<span data-ttu-id="01f45-135">Web 上の Excel での Office スクリプトのスクリプトの基礎</span><span class="sxs-lookup"><span data-stu-id="01f45-135">Scripting Fundamentals for Office Scripts in Excel on the web</span></span>](../develop/scripting-fundamentals.md)
- [<span data-ttu-id="01f45-136">Office スクリプトでのプラットフォームの制限</span><span class="sxs-lookup"><span data-stu-id="01f45-136">Platform Limits with Office Scripts</span></span>](platform-limits.md)
- [<span data-ttu-id="01f45-137">Office スクリプトのパフォーマンスを向上させる</span><span class="sxs-lookup"><span data-stu-id="01f45-137">Improve the performance of your Office Scripts</span></span>](../develop/web-client-performance.md)
- [<span data-ttu-id="01f45-138">Office スクリプトの効果を元に戻す</span><span class="sxs-lookup"><span data-stu-id="01f45-138">Undo the effects of an Office Script</span></span>](undo.md)
