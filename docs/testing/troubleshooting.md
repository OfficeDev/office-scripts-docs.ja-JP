---
title: Office スクリプトのトラブルシューティング
description: Office スクリプトのヒントとテクニック、およびヘルプリソースをデバッグします。
ms.date: 05/14/2020
localization_priority: Normal
ms.openlocfilehash: 6448980eec45214a589444229db0fd781b9fea13
ms.sourcegitcommit: aec3c971c6640429f89b6bb99d2c95ea06725599
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 06/25/2020
ms.locfileid: "44878620"
---
# <a name="troubleshooting-office-scripts"></a><span data-ttu-id="f56f2-103">Office スクリプトのトラブルシューティング</span><span class="sxs-lookup"><span data-stu-id="f56f2-103">Troubleshooting Office Scripts</span></span>

<span data-ttu-id="f56f2-104">Office スクリプトを開発する際には、誤りが発生することがあります。</span><span class="sxs-lookup"><span data-stu-id="f56f2-104">As you develop Office Scripts, you may make mistakes.</span></span> <span data-ttu-id="f56f2-105">大丈夫です。</span><span class="sxs-lookup"><span data-stu-id="f56f2-105">It's okay.</span></span> <span data-ttu-id="f56f2-106">問題を見つけてスクリプトを完全に動作させるためのツールが用意されています。</span><span class="sxs-lookup"><span data-stu-id="f56f2-106">We have tools that help find the problems and get your scripts working perfectly.</span></span>

## <a name="console-logs"></a><span data-ttu-id="f56f2-107">コンソールログ</span><span class="sxs-lookup"><span data-stu-id="f56f2-107">Console logs</span></span>

<span data-ttu-id="f56f2-108">トラブルシューティング中に、画面にメッセージを出力することもできます。</span><span class="sxs-lookup"><span data-stu-id="f56f2-108">Sometimes while troubleshooting, you'll want to print messages to the screen.</span></span> <span data-ttu-id="f56f2-109">これにより、変数の現在の値や、どのコードパスがトリガーされているかを確認できます。</span><span class="sxs-lookup"><span data-stu-id="f56f2-109">These can show you the current value of variables or which code paths are being triggered.</span></span> <span data-ttu-id="f56f2-110">これを行うには、テキストをコンソールに記録します。</span><span class="sxs-lookup"><span data-stu-id="f56f2-110">To do this, log text to the console.</span></span>

```TypeScript
console.log("Logging myRange's address.");
console.log(myRange.getAddress());
```

<span data-ttu-id="f56f2-111">に渡される文字列 `console.log` は、コードエディターのログコンソールに表示されます。</span><span class="sxs-lookup"><span data-stu-id="f56f2-111">Strings passed to`console.log` will be displayed in the Code Editor's logging console.</span></span> <span data-ttu-id="f56f2-112">コンソールをオンにするには、**省略記号**ボタンを押して [**ログ...** ] を選択します。</span><span class="sxs-lookup"><span data-stu-id="f56f2-112">To turn on the console, press the **Ellipses** button and select **Logs...**</span></span>

<span data-ttu-id="f56f2-113">ログがブックに影響を与えることはありません。</span><span class="sxs-lookup"><span data-stu-id="f56f2-113">Logs do not affect the workbook.</span></span>

## <a name="error-messages"></a><span data-ttu-id="f56f2-114">エラー メッセージ</span><span class="sxs-lookup"><span data-stu-id="f56f2-114">Error messages</span></span>

<span data-ttu-id="f56f2-115">Excel スクリプトで問題が発生すると、エラーが生成されます。</span><span class="sxs-lookup"><span data-stu-id="f56f2-115">When your Excel Script encounters a problem running, it produces an error.</span></span> <span data-ttu-id="f56f2-116">**ログを表示**するかどうかの確認を求めるポップアップが表示されます。</span><span class="sxs-lookup"><span data-stu-id="f56f2-116">You'll see a prompt pop-up asking if you want to **View Logs**.</span></span> <span data-ttu-id="f56f2-117">そのボタンを押してコンソールを開き、エラーを表示します。</span><span class="sxs-lookup"><span data-stu-id="f56f2-117">Press that button to open the console and display any errors.</span></span>

## <a name="help-resources"></a><span data-ttu-id="f56f2-118">ヘルプリソース</span><span class="sxs-lookup"><span data-stu-id="f56f2-118">Help resources</span></span>

<span data-ttu-id="f56f2-119">[スタックオーバーフロー](https://stackoverflow.com/questions/tagged/office-scripts)は、コーディングに関する問題の解決に役立つ開発者のコミュニティです。</span><span class="sxs-lookup"><span data-stu-id="f56f2-119">[Stack Overflow](https://stackoverflow.com/questions/tagged/office-scripts) is a community of developers willing to help with coding problems.</span></span> <span data-ttu-id="f56f2-120">多くの場合、クイックスタックオーバーフロー検索を使用して、問題の解決策を見つけることができます。</span><span class="sxs-lookup"><span data-stu-id="f56f2-120">Often, you'll be able to find the solution to your problem through a quick Stack Overflow search.</span></span> <span data-ttu-id="f56f2-121">そうでない場合は、質問して、「office スクリプト」タグでタグを付けてください。</span><span class="sxs-lookup"><span data-stu-id="f56f2-121">If not, ask your question and tag it with the "office-scripts" tag.</span></span> <span data-ttu-id="f56f2-122">Office*アドイン*ではなく、office*スクリプト*を作成していることを必ずお伝えください。</span><span class="sxs-lookup"><span data-stu-id="f56f2-122">Be sure to mention you're creating an Office *Script*, not an Office *Add-in*.</span></span>

<span data-ttu-id="f56f2-123">Office JavaScript API で問題が発生した場合は、 [Officedev/Office/js](https://github.com/OfficeDev/office-js) GitHub リポジトリに問題を作成します。</span><span class="sxs-lookup"><span data-stu-id="f56f2-123">If you encounter a problem with the Office JavaScript API, create an issue in the [OfficeDev/office-js](https://github.com/OfficeDev/office-js) GitHub repository.</span></span> <span data-ttu-id="f56f2-124">製品チームのメンバーは問題に対応し、さらに支援を提供します。</span><span class="sxs-lookup"><span data-stu-id="f56f2-124">Members of the product team will respond to issues and provide further assistance.</span></span> <span data-ttu-id="f56f2-125">**Officedev/office-js**リポジトリで問題を発生させることは、製品チームが対処する必要のある OFFICE JavaScript API ライブラリに問題が見つかったことを示しています。</span><span class="sxs-lookup"><span data-stu-id="f56f2-125">Creating an issue in the **OfficeDev/office-js** repository indicates you have found a flaw in the Office JavaScript API library that the product team should address.</span></span>

<span data-ttu-id="f56f2-126">操作レコーダーまたは Editor に問題がある場合は、Excel の**ヘルプ > フィードバック**ボタンを使用してフィードバックを送信してください。</span><span class="sxs-lookup"><span data-stu-id="f56f2-126">If there is a problem with the Action Recorder or Editor, send feedback through the **Help > Feedback** button in Excel.</span></span>

## <a name="see-also"></a><span data-ttu-id="f56f2-127">関連項目</span><span class="sxs-lookup"><span data-stu-id="f56f2-127">See also</span></span>

- [<span data-ttu-id="f56f2-128">Excel on the web の Office スクリプト</span><span class="sxs-lookup"><span data-stu-id="f56f2-128">Office Scripts in Excel on the web</span></span>](../overview/excel.md)
- [<span data-ttu-id="f56f2-129">Web 上の Excel での Office スクリプトのスクリプトの基礎</span><span class="sxs-lookup"><span data-stu-id="f56f2-129">Scripting Fundamentals for Office Scripts in Excel on the web</span></span>](../develop/scripting-fundamentals.md)
- [<span data-ttu-id="f56f2-130">Office スクリプトの効果を元に戻す</span><span class="sxs-lookup"><span data-stu-id="f56f2-130">Undo the effects of an Office Script</span></span>](undo.md)
- [<span data-ttu-id="f56f2-131">Office スクリプトのパフォーマンスを向上させる</span><span class="sxs-lookup"><span data-stu-id="f56f2-131">Improve the performance of your Office Scripts</span></span>](../develop/web-client-performance.md)
