---
title: Power Automate フローでマクロ ファイルを使用する
description: Power Automate フローでマクロ ファイルまたは xlsm ファイルを使用する方法について説明します。
ms.date: 03/18/2021
localization_priority: Normal
ms.openlocfilehash: a7929fc485ae2118d30a4f2783538d0e04deca2a
ms.sourcegitcommit: 45ffe3dbd2c834b78592ad35928cf8096f5e80bc
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 04/14/2021
ms.locfileid: "51755015"
---
# <a name="how-to-use-macro-files-in-power-automate-flows"></a><span data-ttu-id="06f53-103">Power Automate フローでマクロ ファイルを使用する方法</span><span class="sxs-lookup"><span data-stu-id="06f53-103">How to use macro files in Power Automate flows</span></span>

<span data-ttu-id="06f53-104">[Power Automate フローは](https://flow.microsoft.com/)[、Excel](https://flow.microsoft.com/connectors/shared_excelonlinebusiness/excel-online-business/)ファイルを他の組織データや Teams、Outlook、SharePoint などのアプリに接続するのに役立つ Excel コネクタを提供します。</span><span class="sxs-lookup"><span data-stu-id="06f53-104">[Power Automate flows](https://flow.microsoft.com/) provide [Excel connectors](https://flow.microsoft.com/connectors/shared_excelonlinebusiness/excel-online-business/) to help connect Excel files with the rest of your organizational data and apps such as Teams, Outlook, and SharePoint.</span></span>

<span data-ttu-id="06f53-105">ただし、ファイル ドロップダウンでマクロ ファイルを選択できない (次のスクリーンショットの例を参照)。</span><span class="sxs-lookup"><span data-stu-id="06f53-105">However, macro files can't be selected in the file dropdown (see an example in the following screenshot).</span></span>

:::image type="content" source="../images/no-xlsm.png" alt-text="マクロ ファイルが選択されない状態を示す Power Automate Run スクリプト アクション。表示されるエラーは、'File' が必要です。":::

<span data-ttu-id="06f53-107">この問題を回避する方法の 1 つは、次のスクリーンショットに示すように、"ファイル メタデータの取得" アクション (OneDrive または SharePoint) を含め、"スクリプトの実行" アクションで ID プロパティを使用することです。</span><span class="sxs-lookup"><span data-stu-id="06f53-107">One way to get around this issue is by including the "Get File Metadata" action (OneDrive or SharePoint) and use the ID property in the "Run Script" action as shown in the following screenshot.</span></span>

:::image type="content" source="../images/xlsm-in-pa.png" alt-text="選択したマクロ ファイルとスクリプトの実行エラーを示す Power Automate Run スクリプト アクション。":::

> [!NOTE]
> <span data-ttu-id="06f53-109">一部の XLSM (特に、ActiveX/フォーム コントロールを含む) は、Excel オンライン コネクタでは機能しない場合があります。</span><span class="sxs-lookup"><span data-stu-id="06f53-109">Some XLSM (especially the ones with ActiveX/Form controls) may not work in the Excel online connector.</span></span> <span data-ttu-id="06f53-110">ソリューションを展開する前に必ずテストしてください。</span><span class="sxs-lookup"><span data-stu-id="06f53-110">Be sure to test before deploying your solution.</span></span>

<span data-ttu-id="06f53-111">[![スクリプトの実行アクションでの XLSM の使用に関するビデオを見る](../images/xlsm-vid.png)](https://youtu.be/o-H9BbywJQQ "スクリプトの実行アクションでの XLSM の使用に関するビデオ")</span><span class="sxs-lookup"><span data-stu-id="06f53-111">[![Watch video about using XLSM in Run Script action](../images/xlsm-vid.png)](https://youtu.be/o-H9BbywJQQ "Video about using XLSM in Run Script action")</span></span>
