---
title: データ フローでマクロ ファイルをPower Automateする
description: これらのフローでマクロ ファイルまたは xlsm ファイルを使用するPower Automateします。
ms.date: 04/28/2021
localization_priority: Normal
ms.openlocfilehash: b232a1d31a7ff6e28016c5e28fd8a83c8d3f1859
ms.sourcegitcommit: f7a7aebfb687f2a35dbed07ed62ff352a114525a
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 05/06/2021
ms.locfileid: "52232656"
---
# <a name="how-to-use-macro-files-in-power-automate-flows"></a><span data-ttu-id="6bba1-103">データ フローでマクロ ファイルをPower Automateする方法</span><span class="sxs-lookup"><span data-stu-id="6bba1-103">How to use macro files in Power Automate flows</span></span>

<span data-ttu-id="6bba1-104">[Power Automateフロー](https://flow.microsoft.com/)は、Excel[](https://flow.microsoft.com/connectors/shared_excelonlinebusiness/excel-online-business/)ファイルを Teams、Outlook、および SharePoint などの他の組織データと接続するのに役立つ Excel コネクタを提供します。</span><span class="sxs-lookup"><span data-stu-id="6bba1-104">[Power Automate flows](https://flow.microsoft.com/) provide [Excel connectors](https://flow.microsoft.com/connectors/shared_excelonlinebusiness/excel-online-business/) to help connect Excel files with the rest of your organizational data and apps such as Teams, Outlook, and SharePoint.</span></span>

<span data-ttu-id="6bba1-105">ただし、ファイル ドロップダウンでマクロ ファイルを選択できない (次のスクリーンショットの例を参照)。</span><span class="sxs-lookup"><span data-stu-id="6bba1-105">However, macro files can't be selected in the file dropdown (see an example in the following screenshot).</span></span>

:::image type="content" source="../images/no-xlsm.png" alt-text="[Power Automateスクリプトの実行] アクションで、選択されているマクロ ファイルが表示されません。表示されるエラーは 'File' が必要です":::

<span data-ttu-id="6bba1-107">この問題を回避する 1 つの方法は、次のスクリーンショットに示すように、"ファイル メタデータの取得" アクション (OneDrive または SharePoint) を含め、"スクリプトの実行" アクションで ID プロパティを使用することです。</span><span class="sxs-lookup"><span data-stu-id="6bba1-107">One way to get around this issue is by including the "Get File Metadata" action (OneDrive or SharePoint) and use the ID property in the "Run Script" action as shown in the following screenshot.</span></span>

:::image type="content" source="../images/xlsm-in-pa.png" alt-text="[Power Automateスクリプトの実行] アクションで、マクロ ファイルが選択され、スクリプトの実行エラーが表示されません。":::

> [!NOTE]
> <span data-ttu-id="6bba1-109">一部の XLSM (特に、ActiveX/フォーム コントロールを持つもの) は、オンライン コネクタExcel場合があります。</span><span class="sxs-lookup"><span data-stu-id="6bba1-109">Some XLSM (especially the ones with ActiveX/Form controls) may not work in the Excel online connector.</span></span> <span data-ttu-id="6bba1-110">ソリューションを展開する前に必ずテストしてください。</span><span class="sxs-lookup"><span data-stu-id="6bba1-110">Be sure to test before deploying your solution.</span></span>

## <a name="other-resources"></a><span data-ttu-id="6bba1-111">その他のリソース</span><span class="sxs-lookup"><span data-stu-id="6bba1-111">Other resources</span></span>

<span data-ttu-id="6bba1-112">[スクリプトの実行アクションで .xlsm ファイルを使用する方法については、Sudhi Ramamurthy の YouTube ビデオをご覧ください](https://youtu.be/o-H9BbywJQQ)。</span><span class="sxs-lookup"><span data-stu-id="6bba1-112">[Watch Sudhi Ramamurthy's YouTube video on how use an .xlsm file in a Run Script action](https://youtu.be/o-H9BbywJQQ).</span></span>
