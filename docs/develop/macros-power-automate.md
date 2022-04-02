---
title: マクロが有効なファイルをデータ フロー Power Automateする
description: マクロが有効なファイル (.xlsm ファイル) を使用する方法については、Power Automateします。
ms.date: 03/24/2022
ms.localizationpriority: medium
ms.openlocfilehash: 9f2ecefe9fb97d1c5514ddb52c3cbcd0596df426
ms.sourcegitcommit: 7023b9e23499806901a5ecf8ebc460b76887cca6
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 03/31/2022
ms.locfileid: "64585745"
---
# <a name="how-to-use-macro-enabled-files-in-power-automate-flows"></a>データ フローでマクロが有効なファイルをPower Automateする方法

.xlsm ファイルを新しいフロー Power Automateできます。 これにより、既存のオートメーション ソリューションを Web ベースの形式に変換できます。 .xslm ファイルに含まれるマクロは、ファイル内で実行Power Automate。 スクリプトOffice有効になっているのは、このスクリプトのみです。

通常[Excelのオンライン (Business)](https://flow.microsoft.com/connectors/shared_excelonlinebusiness/excel-online-business/) コネクタは、Power Automate [](https://flow.microsoft.com/) Open XML スプレッドシート (Microsoft Excel) 形式のファイルに.xlsxされます。 そのファイル ブラウザーでは、ファイルを選択.xlsxできます。 ただし、マクロが有効なファイルは、ファイル メタデータが使用されている場合、コネクタの **スクリプト** の実行アクションと互換性があります。

フローで、[ファイル **メタデータの取得**] アクションを使用して、[OneDrive for BusinessまたはSharePoint](https://flow.microsoft.com/connectors/shared_onedriveforbusiness/onedrive-for-business/)[します。](https://flow.microsoft.com/connectors/shared_sharepointonline/sharepoint/) [ **スクリプトの実行]** アクションは、このメタデータを有効なファイルとして受け入れる。 スクリプトを *実行する場合* は、[ファイル メタデータの **取得] アクション** から返される ID 動的コンテンツを "File" 引数として使用します。 次のスクリーンショットは、"Test Macro File.xlsm" と呼ばれるファイルのメタデータをスクリプトの実行アクションに提供する **フローを示** しています。

:::image type="content" source="../images/xlsm-in-power-automate.png" alt-text="マクロ ファイルのメタデータをスクリプトの実行アクションに渡すファイル メタデータの取得アクションを含むフロー。":::

> [!WARNING]
> 一部の .xlsm ファイル(特に、ActiveXまたはフォーム コントロールを持つファイル)は、オンライン コネクタExcel場合があります。 ソリューションを展開する前に必ずテストしてください。

## <a name="other-resources"></a>その他のリソース

[スクリプトの実行アクションで .xlsm ファイルを使用する方法については、Sudhi Ramamurthy の YouTube ビデオをご覧ください](https://youtu.be/o-H9BbywJQQ)。
