---
title: データ フローでマクロ ファイルをPower Automateする
description: これらのフローでマクロ ファイルまたは xlsm ファイルを使用するPower Automateします。
ms.date: 04/28/2021
localization_priority: Normal
ms.openlocfilehash: 91e11424e4220a3e1f80cdd2711d05f219016147
ms.sourcegitcommit: 4693c8f79428ec74695328275703af0ba1bfea8f
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 06/23/2021
ms.locfileid: "53074642"
---
# <a name="how-to-use-macro-files-in-power-automate-flows"></a>データ フローでマクロ ファイルをPower Automateする方法

[Power Automateフロー](https://flow.microsoft.com/)は、Excel[](https://flow.microsoft.com/connectors/shared_excelonlinebusiness/excel-online-business/)ファイルを Teams、Outlook、および SharePoint などの他の組織データと接続するのに役立つ Excel コネクタを提供します。

ただし、ファイル ドロップダウンでマクロ ファイルを選択できない (次のスクリーンショットの例を参照)。

:::image type="content" source="../images/no-xlsm.png" alt-text="[Power Automateスクリプトの実行] アクションで、選択されているマクロ ファイルが表示されません。表示されるエラーは、'File' が必要です。":::

この問題を回避する 1 つの方法は、次のスクリーンショットに示すように、"ファイル メタデータの取得" アクション (OneDrive または SharePoint) を含め、"スクリプトの実行" アクションで ID プロパティを使用することです。

:::image type="content" source="../images/xlsm-in-pa.png" alt-text="[Power Automateスクリプトの実行] アクションで、マクロ ファイルが選択され、スクリプトの実行エラーが表示されません。":::

> [!NOTE]
> 一部の XLSM (特に、ActiveX/フォーム コントロールを持つもの) は、オンライン コネクタExcel場合があります。 ソリューションを展開する前に必ずテストしてください。

## <a name="other-resources"></a>その他のリソース

[スクリプトの実行アクションで .xlsm ファイルを使用する方法については、Sudhi Ramamurthy の YouTube ビデオをご覧ください](https://youtu.be/o-H9BbywJQQ)。
