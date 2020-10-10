---
title: Office スクリプトを使用したプラットフォームの制限と要件
description: Web 上の Excel で使用する場合の Office スクリプトのリソース制限とブラウザーサポート
ms.date: 10/09/2020
localization_priority: Normal
ms.openlocfilehash: df468192f443b912e26411e46c9f953e046e55ec
ms.sourcegitcommit: 42fa3b629c93930b4e73e9c4c01d0c8bdf6d7487
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 10/09/2020
ms.locfileid: "48411558"
---
# <a name="platform-limits-and-requirements-with-office-scripts"></a>Office スクリプトを使用したプラットフォームの制限と要件

Office スクリプトを開発する際には、いくつかのプラットフォームの制限事項に注意する必要があります。 この記事では、web 上の Excel 用 Office スクリプトのブラウザーのサポートとデータの制限について説明します。

## <a name="browser-support"></a>ブラウザのサポート

Office スクリプト [は、web 用の office をサポート](https://support.microsoft.com/office/ad1303e0-a318-47aa-b409-d3a5eb44e452)する任意のブラウザーで動作します。 ただし、一部の JavaScript 機能は Internet Explorer 11 (IE 11) ではサポートされていません。 ES6 以降で導入された機能は、IE 11 で [は](https://www.w3schools.com/Js/js_es6.asp) 動作しません。 組織内のユーザーが依然としてそのブラウザーを使用している場合は、その環境でスクリプトを共有するときに必ずテストしてください。

[!INCLUDE [Teams support note](../includes/teams-support-note.md)]

### <a name="third-party-cookies"></a>サードパーティの cookie

ブラウザーでは、web 上の Excel で [ **自動化** ] タブが表示されるように、サードパーティの cookie が有効になっている必要があります。 タブが表示されていない場合は、ブラウザーの設定を確認します。 プライベートブラウザーセッションを使用している場合は、この設定を毎回有効にしなければならない場合があります。

> [!NOTE]
> 一部のブラウザーは、"サードパーティの cookie" ではなく "すべての cookie" としてこの設定を参照します。

## <a name="data-limits"></a>データの上限

一度に転送できる Excel データの量と、実行できる個々の電力を自動化するトランザクションの数には制限があります。

### <a name="excel"></a>Excel

スクリプトを使用してブックを呼び出すときに、web 用の Excel には次の制限があります。

- 要求と応答は **5 mb**に制限されます。
- 範囲は **500万のセル**に制限されます。

大規模なデータセットを処理するときにエラーが発生した場合は、大きな範囲ではなく、複数の狭い範囲を使用してください。 範囲外の [セル](/javascript/api/office-scripts/excelscript/excelscript.range#getspecialcells-celltype--cellvaluetype-) のような api を使用して、大きな範囲ではなく特定のセルを対象にすることもできます。

### <a name="power-automate"></a>Power Automate

Office スクリプトを電源自動化と共に使用する場合、1 **日あたりの通話**は最大200に制限されています。 この制限は、12:00 AM UTC でリセットされます。

Power 自動プラットフォームにも使用上の制限があります。これは、「 [Power 自動検出の制限と構成](/power-automate/limits-and-config)」に記載されています。

## <a name="see-also"></a>関連項目

- [Office スクリプトのトラブルシューティング](troubleshooting.md)
- [Office スクリプトの効果を元に戻す](undo.md)
- [Office スクリプトのパフォーマンスを向上させる](../develop/web-client-performance.md)
- [Web 上の Excel での Office スクリプトのスクリプトの基礎](../develop/scripting-fundamentals.md)
