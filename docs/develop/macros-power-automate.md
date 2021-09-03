---
title: データ フローでマクロ ファイルをPower Automateする
description: これらのフローでマクロ ファイルまたは xlsm ファイルを使用するPower Automateします。
ms.date: 09/01/2021
localization_priority: Normal
ms.openlocfilehash: 204eb8315f90c0ab0e20a34ec64517fbfbf304b1
ms.sourcegitcommit: 9472e78eb186ceffdaaceb2718d5074ce55a0e74
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 09/02/2021
ms.locfileid: "58866540"
---
# <a name="how-to-use-macro-files-in-power-automate-flows"></a>データ フローでマクロ ファイルをPower Automateする方法

通常[Excelオンライン (Business)](https://flow.microsoft.com/connectors/shared_excelonlinebusiness/excel-online-business/)コネクタは、Power Automate Microsoft Excel Open XML スプレッドシート (.xlsx) 形式のファイルでのみ動作します。 [](https://flow.microsoft.com/) ファイル ブラウザーは、コネクタ内のファイル.xlsx選択を制限します。 ただし、ファイル メタデータが使用されている場合、マクロ ファイルはコネクタの **スクリプト** の実行アクションと互換性があります。

フローで、[ファイル メタデータ **の取得**] アクションを使用して、OneDrive for Business [または](https://flow.microsoft.com/connectors/shared_onedriveforbusiness/onedrive-for-business/)SharePoint [します。](https://flow.microsoft.com/connectors/shared_sharepointonline/sharepoint/) [ **スクリプトの実行]** アクションは、このメタデータを有効なファイルとして受け入れる。 スクリプトを *実行する場合* は、[ファイル メタデータの **取得]** アクションから返される ID 動的コンテンツを "File" 引数として使用します。 次のスクリーンショットは、スクリプトの実行アクションに "Test Macro File.xlsm" というファイルのメタデータを提供するフロー **を示** しています。

:::image type="content" source="../images/xlsm-in-power-automate.png" alt-text="マクロ ファイルのメタデータをスクリプトの実行アクションに渡すファイル メタデータの取得アクションを含むフロー。":::

> [!WARNING]
> 一部の .xlsm ファイル (特に、ActiveXまたはフォーム コントロールを持つファイル) は、オンライン コネクタExcel機能しない場合があります。 ソリューションを展開する前に必ずテストしてください。

## <a name="other-resources"></a>その他のリソース

[スクリプトの実行アクションで .xlsm ファイルを使用する方法については、Sudhi Ramamurthy の YouTube ビデオをご覧ください](https://youtu.be/o-H9BbywJQQ)。
