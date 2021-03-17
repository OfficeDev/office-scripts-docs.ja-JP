---
title: プラットフォームの制限と要件 (スクリプトOffice)
description: Web 上の Excel で使用する場合Officeスクリプトのリソース制限とブラウザーのサポート
ms.date: 03/12/2021
localization_priority: Normal
ms.openlocfilehash: 93307b6204f409f26c77b5ead33188205d5c4b4d
ms.sourcegitcommit: 5bde455b06ee2ed007f3e462d8ad485b257774ef
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 03/17/2021
ms.locfileid: "50837266"
---
# <a name="platform-limits-and-requirements-with-office-scripts"></a>プラットフォームの制限と要件 (スクリプトOffice)

スクリプトの開発時に注意する必要があるプラットフォームのOfficeがあります。 この記事では、Web 上の Excel 用スクリプトOfficeブラウザーのサポートとデータ制限について説明します。

## <a name="browser-support"></a>ブラウザのサポート

Officeスクリプトは、Web 用のOffice [をサポートする任意のブラウザーで動作します](https://support.microsoft.com/office/ad1303e0-a318-47aa-b409-d3a5eb44e452)。 ただし、JavaScript の一部の機能は、11 Internet Explorer (IE 11) ではサポートされていません。 [ES6](https://www.w3schools.com/Js/js_es6.asp)以降で導入された機能は、IE 11 では動作しません。 組織内のユーザーが引き続きそのブラウザーを使用している場合は、共有するときに、その環境でスクリプトをテストしてください。

[!INCLUDE [Teams support note](../includes/teams-support-note.md)]

### <a name="third-party-cookies"></a>サードパーティの Cookie

Web 上の Excel で [自動化] タブを表示するには、ブラウザーでサードパーティの Cookie が有効になっている必要があります。 タブが表示されていない場合は、ブラウザーの設定を確認します。 プライベート ブラウザー セッションを使用している場合は、その度にこの設定を再び有効にする必要があります。

> [!NOTE]
> 一部のブラウザーでは、この設定を "サードパーティ Cookie" ではなく"すべての Cookie" と呼ぶ場合があります。

#### <a name="instructions-for-adjusting-cookie-settings-in-popular-browsers"></a>一般的なブラウザーで Cookie 設定を調整する手順

- [Chrome](https://support.google.com/chrome/answer/95647)
- [Edge](https://support.microsoft.com/microsoft-edge/temporarily-allow-cookies-and-site-data-in-microsoft-edge-597f04f2-c0ce-f08c-7c2b-541086362bd2)
- [Firefox](https://support.mozilla.org/kb/disable-third-party-cookies)
- [Safari](https://support.apple.com/guide/safari/manage-cookies-and-website-data-sfri11471/mac)

## <a name="data-limits"></a>データの上限

一度に転送できる Excel データの量と、個々の Power Automate トランザクションを実行できる数には制限があります。

### <a name="excel"></a>Excel

スクリプトを使用してブックを呼び出す場合、Web 用の Excel には次の制限があります。

- 要求と応答は **5 MB に制限されています**。
- 範囲は 500 万 **セルに制限されます**。

大規模なデータセットを扱う際にエラーが発生する場合は、より大きな範囲ではなく、複数の小さい範囲を使用してみてください。 [Range.getSpecialCells](/javascript/api/office-scripts/excelscript/excelscript.range#getspecialcells-celltype--cellvaluetype-)のような API を使用して、大きな範囲ではなく特定のセルをターゲットにすることもできます。

### <a name="power-automate"></a>Power Automate

Power Automate で Office スクリプトを使用する場合、各ユーザーは **1 日あたり 200 回の呼び出しに制限されます**。 この制限は、UTC の午前 12:00 にリセットされます。

Power Automate プラットフォームには使用上の制限があります。これは次の記事で確認できます。

- [Power Automate の制限と構成](/power-automate/limits-and-config)
- [Excel Online (Business) コネクタの既知の問題と制限事項](/connectors/excelonlinebusiness/#known-issues-and-limitations)

## <a name="see-also"></a>関連項目

- [Office スクリプトのトラブルシューティング](troubleshooting.md)
- [Office スクリプトの効果を元に戻す](undo.md)
- [スクリプトのパフォーマンスをOfficeする](../develop/web-client-performance.md)
- [Web 上の Excel Officeスクリプトのスクリプトの基本](../develop/scripting-fundamentals.md)
