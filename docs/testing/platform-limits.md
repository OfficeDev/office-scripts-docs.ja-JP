---
title: Office スクリプトを使用したプラットフォームの制限と要件
description: Excel on the webで使用する場合のリソース制限とOfficeスクリプトのブラウザサポート
ms.date: 05/17/2021
localization_priority: Normal
ms.openlocfilehash: 7e81aaf2f96faeb67c815814fe3b7f1795651318
ms.sourcegitcommit: 4687693f02fc90a57ba30c461f35046e02e6f5fb
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 05/19/2021
ms.locfileid: "52545582"
---
# <a name="platform-limits-and-requirements-with-office-scripts"></a>Office スクリプトを使用したプラットフォームの制限と要件

Officeスクリプトを開発する際に注意する必要があるプラットフォームの制限事項がいくつかあります。 この記事では、ブラウザーのサポートとExcel on the web用のOffice スクリプトのデータ制限について詳しく説明します。

## <a name="browser-support"></a>ブラウザのサポート

Officeスクリプトは、web の[Officeをサポート](https://support.microsoft.com/office/ad1303e0-a318-47aa-b409-d3a5eb44e452)する任意のブラウザーで動作します。 ただし、一部の JavaScript 機能は、インターネット エクスプ ローラー 11 (IE 11) ではサポートされていません。 [ES6 以降](https://www.w3schools.com/Js/js_es6.asp)で導入された機能は、IE 11 では動作しません。 組織のユーザーがそのブラウザーを引き続き使用している場合は、スクリプトを共有するときに必ずその環境でスクリプトをテストしてください。

[!INCLUDE [Teams support note](../includes/teams-support-note.md)]

### <a name="third-party-cookies"></a>サードパーティのクッキー

お使いのブラウザでは、Excel on the webの[**自動化**]タブを表示するためにサードパーティのクッキーを有効にする必要があります。 タブが表示されていない場合は、ブラウザの設定を確認してください。 プライベートブラウザセッションを使用している場合は、毎回この設定を再度有効にする必要があります。

> [!NOTE]
> ブラウザによっては、この設定を「サードパーティのクッキー」ではなく「すべてのクッキー」と呼んでいます。

#### <a name="instructions-for-adjusting-cookie-settings-in-popular-browsers"></a>一般的なブラウザでクッキーの設定を調整する手順

- [Chrome](https://support.google.com/chrome/answer/95647)
- [Edge](https://support.microsoft.com/microsoft-edge/temporarily-allow-cookies-and-site-data-in-microsoft-edge-597f04f2-c0ce-f08c-7c2b-541086362bd2)
- [Firefox](https://support.mozilla.org/kb/disable-third-party-cookies)
- [Safari](https://support.apple.com/guide/safari/manage-cookies-and-website-data-sfri11471/mac)

## <a name="data-limits"></a>データの上限

一度に転送できるデータExcel量と、個々のPower Automateトランザクションの数には制限があります。

### <a name="excel"></a>Excel

スクリプトを使用してブックを呼び出す場合、web のExcelには次の制限があります。

- 要求と応答は 5 **MB** に制限されています。
- 範囲は **500 万個のセル** に制限されます。

大きなデータセットを扱うときにエラーが発生した場合は、より大きな範囲ではなく、複数の小さい範囲を使用してみてください。 例については、「 [大規模なデータセットの記述サンプル」](../resources/samples/write-large-dataset.md) を参照してください。 [Range.getSpecialCells](/javascript/api/office-scripts/excelscript/excelscript.range#getspecialcells-celltype--cellvaluetype-)などの API を使用して、大きな範囲ではなく特定のセルをターゲットにすることもできます。

### <a name="power-automate"></a>Power Automate

Power AutomateでOfficeスクリプトを使用する場合、各ユーザーは **1 日あたりスクリプトの実行アクションに対して 400 回の呼び出しを行う** 必要があります。 この制限は、UTC の午前 12 時にリセットされます。

Power Automate プラットフォームには、次の記事で説明する使用制限もあります。

- [Power Automateの制限と構成](/power-automate/limits-and-config)
- [Excel オンライン (ビジネス) コネクタの既知の問題と制限事項](/connectors/excelonlinebusiness/#known-issues-and-limitations)

## <a name="see-also"></a>関連項目

- [Office スクリプトのトラブルシューティング](troubleshooting.md)
- [Office スクリプトの効果を元に戻す](undo.md)
- [Officeスクリプトのパフォーマンスを向上させる](../develop/web-client-performance.md)
- [Excel on the webでのスクリプトのスクリプトOfficeの基礎](../develop/scripting-fundamentals.md)
