---
title: Office スクリプトのトラブルシューティング
description: Office スクリプトのヒントとテクニック、およびヘルプリソースをデバッグします。
ms.date: 12/13/2019
localization_priority: Normal
ms.openlocfilehash: 959faff875f342dc1b1ab158ad9ded24732b0894
ms.sourcegitcommit: b075eed5a6f275274fbbf6d62633219eac416f26
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 03/10/2020
ms.locfileid: "42700358"
---
# <a name="troubleshooting-office-scripts"></a><span data-ttu-id="459fa-103">Office スクリプトのトラブルシューティング</span><span class="sxs-lookup"><span data-stu-id="459fa-103">Troubleshooting Office Scripts</span></span>

<span data-ttu-id="459fa-104">Office スクリプトを開発する際には、誤りが発生することがあります。</span><span class="sxs-lookup"><span data-stu-id="459fa-104">As you develop Office Scripts, you may make mistakes.</span></span> <span data-ttu-id="459fa-105">大丈夫です。</span><span class="sxs-lookup"><span data-stu-id="459fa-105">It's okay.</span></span> <span data-ttu-id="459fa-106">問題を見つけてスクリプトを完全に動作させるためのツールが用意されています。</span><span class="sxs-lookup"><span data-stu-id="459fa-106">We have tools that help find the problems and get your scripts working perfectly.</span></span>

## <a name="console-logs"></a><span data-ttu-id="459fa-107">コンソールログ</span><span class="sxs-lookup"><span data-stu-id="459fa-107">Console logs</span></span>

<span data-ttu-id="459fa-108">トラブルシューティング中に、画面にメッセージを出力することもできます。</span><span class="sxs-lookup"><span data-stu-id="459fa-108">Sometimes while troubleshooting, you'll want to print messages to the screen.</span></span> <span data-ttu-id="459fa-109">これにより、変数の現在の値や、どのコードパスがトリガーされているかを確認できます。</span><span class="sxs-lookup"><span data-stu-id="459fa-109">These can show you the current value of variables or which code paths are being triggered.</span></span> <span data-ttu-id="459fa-110">これを行うには、テキストをコンソールに記録します。</span><span class="sxs-lookup"><span data-stu-id="459fa-110">To do this, log text to the console.</span></span>

```TypeScript
console.log("Logging my range's address.");
myRange.load("address");
await context.sync();
console.log(myRange.address);
```

> [!IMPORTANT]
> <span data-ttu-id="459fa-111">オブジェクトプロパティを`load`ログに記録`sync`する前に、ワークシートデータとブックを忘れずに使用してください。</span><span class="sxs-lookup"><span data-stu-id="459fa-111">Don't forget to `load` worksheet data and `sync` with the workbook before logging object properties.</span></span>

<span data-ttu-id="459fa-112">に`console.log`渡される文字列は、コードエディターのログコンソールに表示されます。</span><span class="sxs-lookup"><span data-stu-id="459fa-112">Strings passed to`console.log` will be displayed in the Code Editor's logging console.</span></span> <span data-ttu-id="459fa-113">コンソールをオンにするには、**省略記号**ボタンを押して [**ログ...** ] を選択します。</span><span class="sxs-lookup"><span data-stu-id="459fa-113">To turn on the console, press the **Ellipses** button and select **Logs...**</span></span>

<span data-ttu-id="459fa-114">ログがブックに影響を与えることはありません。</span><span class="sxs-lookup"><span data-stu-id="459fa-114">Logs do not affect the workbook.</span></span>

## <a name="error-messages"></a><span data-ttu-id="459fa-115">エラー メッセージ</span><span class="sxs-lookup"><span data-stu-id="459fa-115">Error messages</span></span>

<span data-ttu-id="459fa-116">Excel スクリプトで問題が発生すると、エラーが生成されます。</span><span class="sxs-lookup"><span data-stu-id="459fa-116">When your Excel Script encounters a problem running, it produces an error.</span></span> <span data-ttu-id="459fa-117">**ログを表示**するかどうかの確認を求めるポップアップが表示されます。</span><span class="sxs-lookup"><span data-stu-id="459fa-117">You'll see a prompt pop-up asking if you want to **View Logs**.</span></span> <span data-ttu-id="459fa-118">そのボタンを押してコンソールを開き、エラーを表示します。</span><span class="sxs-lookup"><span data-stu-id="459fa-118">Press that button to open the console and display any errors.</span></span>

## <a name="help-resources"></a><span data-ttu-id="459fa-119">ヘルプリソース</span><span class="sxs-lookup"><span data-stu-id="459fa-119">Help resources</span></span>

<span data-ttu-id="459fa-120">[スタックオーバーフロー](https://stackoverflow.com/questions/tagged/office-scripts)は、コーディングに関する問題の解決に役立つ開発者のコミュニティです。</span><span class="sxs-lookup"><span data-stu-id="459fa-120">[Stack Overflow](https://stackoverflow.com/questions/tagged/office-scripts) is a community of developers willing to help with coding problems.</span></span> <span data-ttu-id="459fa-121">多くの場合、クイックスタックオーバーフロー検索を使用して、問題の解決策を見つけることができます。</span><span class="sxs-lookup"><span data-stu-id="459fa-121">Often, you'll be able to find the solution to your problem through a quick Stack Overflow search.</span></span> <span data-ttu-id="459fa-122">そうでない場合は、質問して、「office スクリプト」タグでタグを付けてください。</span><span class="sxs-lookup"><span data-stu-id="459fa-122">If not, ask your question and tag it with the "office-scripts" tag.</span></span> <span data-ttu-id="459fa-123">Office*アドイン*ではなく、office*スクリプト*を作成していることを必ずお伝えください。</span><span class="sxs-lookup"><span data-stu-id="459fa-123">Be sure to mention you're creating an Office *Script*, not an Office *Add-in*.</span></span>

<span data-ttu-id="459fa-124">Office JavaScript API で問題が発生した場合は、 [Officedev/Office/js](https://github.com/OfficeDev/office-js) GitHub リポジトリに問題を作成します。</span><span class="sxs-lookup"><span data-stu-id="459fa-124">If you encounter a problem with the Office JavaScript API, create an issue in the [OfficeDev/office-js](https://github.com/OfficeDev/office-js) GitHub repository.</span></span> <span data-ttu-id="459fa-125">製品チームのメンバーは問題に対応し、さらに支援を提供します。</span><span class="sxs-lookup"><span data-stu-id="459fa-125">Members of the product team will respond to issues and provide further assistance.</span></span> <span data-ttu-id="459fa-126">**Officedev/office-js**リポジトリで問題を発生させることは、製品チームが対処する必要のある OFFICE JavaScript API ライブラリに問題が見つかったことを示しています。</span><span class="sxs-lookup"><span data-stu-id="459fa-126">Creating an issue in the **OfficeDev/office-js** repository indicates you have found a flaw in the Office JavaScript API library that the product team should address.</span></span>

<span data-ttu-id="459fa-127">操作レコーダーまたは Editor に問題がある場合は、Excel の**ヘルプ > フィードバック**ボタンを使用してフィードバックを送信してください。</span><span class="sxs-lookup"><span data-stu-id="459fa-127">If there is a problem with the Action Recorder or Editor, send feedback through the **Help > Feedback** button in Excel.</span></span>

## <a name="see-also"></a><span data-ttu-id="459fa-128">関連項目</span><span class="sxs-lookup"><span data-stu-id="459fa-128">See also</span></span>

- [<span data-ttu-id="459fa-129">Web 上の Excel での Office スクリプト</span><span class="sxs-lookup"><span data-stu-id="459fa-129">Office Scripts in Excel on the web</span></span>](../overview/excel.md)
- [<span data-ttu-id="459fa-130">Web 上の Excel での Office スクリプトのスクリプトの基礎</span><span class="sxs-lookup"><span data-stu-id="459fa-130">Scripting Fundamentals for Office Scripts in Excel on the web</span></span>](../develop/scripting-fundamentals.md)
- [<span data-ttu-id="459fa-131">Office スクリプトの効果を元に戻す</span><span class="sxs-lookup"><span data-stu-id="459fa-131">Undo the effects of an Office Script</span></span>](undo.md)
