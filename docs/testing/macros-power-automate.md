---
title: Power Automate フローでマクロ ファイルを使用する
description: Power Automate フローでマクロ ファイルまたは xlsm ファイルを使用する方法について説明します。
ms.date: 03/18/2021
localization_priority: Normal
ms.openlocfilehash: ec1fe00eb9ddc382ae4bc02187de7a36c97288b1
ms.sourcegitcommit: 5d24e77df70aa2c1c982275d53213c2a9323ff86
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 04/02/2021
ms.locfileid: "51571477"
---
# <a name="how-to-use-macro-files-in-power-automate-flows"></a>Power Automate フローでマクロ ファイルを使用する方法

[Power Automate フローは](https://flow.microsoft.com/)[、Excel](https://flow.microsoft.com/connectors/shared_excelonlinebusiness/excel-online-business/)ファイルを他の組織データや Teams、Outlook、SharePoint などのアプリに接続するのに役立つ Excel コネクタを提供します。

ただし、ファイル ドロップダウンでマクロ ファイルを選択できない (次のスクリーンショットの例を参照)。

![スクリプトの実行アクションに xlsm はありません](../images/no-xlsm.png)

この問題を回避する方法の 1 つは、次のスクリーンショットに示すように、"ファイル メタデータの取得" アクション (OneDrive または SharePoint) を含め、"スクリプトの実行" アクションで ID プロパティを使用することです。

![スクリプトの実行アクションの xlsm](../images/xlsm-in-pa.png)

> [!NOTE]
> 一部の XLSM (特に、ActiveX/フォーム コントロールを含む) は、Excel オンライン コネクタでは機能しない場合があります。 ソリューションを展開する前に必ずテストしてください。

[![スクリプトの実行アクションでの XLSM の使用に関するビデオを見る](../images/xlsm-vid.png)](https://youtu.be/o-H9BbywJQQ "スクリプトの実行アクションでの XLSM の使用に関するビデオ")
