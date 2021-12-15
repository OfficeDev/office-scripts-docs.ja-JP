---
title: プラットフォームの制限と要件 (スクリプトOffice)
description: スクリプトと一緒に使用する場合Officeスクリプトのリソース制限とブラウザー Excel on the web
ms.date: 12/14/2021
ms.localizationpriority: medium
ms.openlocfilehash: e0c2c8108bb36dc979b33e195cdc0e0e17a586ba
ms.sourcegitcommit: 1e561d245d9ab7f2529a1b071b1fa0326e37c4ca
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 12/15/2021
ms.locfileid: "61513630"
---
# <a name="platform-limits-and-requirements-with-office-scripts"></a>プラットフォームの制限と要件 (スクリプトOffice)

スクリプトの開発時に注意する必要があるプラットフォームのOfficeがあります。 この記事では、ブラウザーのサポートとデータ制限について詳OfficeスクリプトのExcel on the web。

## <a name="browser-support"></a>ブラウザのサポート

Officeスクリプトは、スクリプトをサポートするブラウザー[でOffice for the web。](https://support.microsoft.com/office/ad1303e0-a318-47aa-b409-d3a5eb44e452) ただし、JavaScript の一部の機能は、11 Internet Explorer (IE 11) ではサポートされていません。 [ES6](https://www.w3schools.com/Js/js_es6.asp)以降で導入された機能は、IE 11 では動作しません。 組織内のユーザーが引き続きそのブラウザーを使用している場合は、共有するときに、その環境でスクリプトをテストしてください。

[!INCLUDE [Teams support note](../includes/teams-support-note.md)]

### <a name="third-party-cookies"></a>サードパーティの Cookie

ブラウザーで [自動化] タブを表示するには、サードパーティの Cookie が **有効になっている必要** Excel on the web。 タブが表示されていない場合は、ブラウザーの設定を確認します。 プライベート ブラウザー セッションを使用している場合は、その度にこの設定を再び有効にする必要があります。

> [!NOTE]
> 一部のブラウザーでは、この設定を "サードパーティ Cookie" ではなく"すべての Cookie" と呼ぶ場合があります。

#### <a name="instructions-for-adjusting-cookie-settings-in-popular-browsers"></a>一般的なブラウザーで Cookie 設定を調整する手順

- [Chrome](https://support.google.com/chrome/answer/95647)
- [Edge](https://support.microsoft.com/microsoft-edge/597f04f2-c0ce-f08c-7c2b-541086362bd2)
- [Firefox](https://support.mozilla.org/kb/disable-third-party-cookies)
- [Safari](https://support.apple.com/guide/safari/manage-cookies-and-website-data-sfri11471/mac)

## <a name="data-limits"></a>データの上限

データを一度にExcelできるデータの量と、トランザクションを実行できるPower Automate制限があります。

### <a name="excel"></a>Excel

Excel for the webを使用してブックを呼び出す場合、次の制限があります。

- 要求と応答は **5 MB に制限されています**。
- 範囲は 500 万 **セルに制限されます**。

大規模なデータセットを扱う際にエラーが発生する場合は、より大きな範囲ではなく、複数の小さい範囲を使用してみてください。 例については、「大規模なデータセットを書 [き込む」サンプルを参照](../resources/samples/write-large-dataset.md) してください。 [Range.getSpecialCells](/javascript/api/office-scripts/excelscript/excelscript.range#getSpecialCells_cellType__cellValueType_)のような API を使用して、大きな範囲ではなく特定のセルをターゲットにすることもできます。

### <a name="power-automate"></a>Power Automate

ユーザーが Office スクリプトPower Automate使用する場合、各ユーザーは 1 日にスクリプトの実行アクションに対して **800 回の呼び出しに制限されます**。 この制限は、UTC の午前 12:00 にリセットされます。

またPower Automateプラットフォームには使用上の制限があります。これは次の記事で確認できます。

- [サーバーの制限と構成Power Automate](/power-automate/limits-and-config)
- [オンライン (Business) コネクタExcel既知の問題と制限事項](/connectors/excelonlinebusiness/#known-issues-and-limitations)

> [!NOTE]
> 実行時間が長いスクリプトがある場合は、同期的なスクリプト操作の[120 秒のタイムアウトPower Automate注意してください](/power-automate/limits-and-config#timeout)。 スクリプトを最適化[するか、オートメーション](../develop/web-client-performance.md)を複数のスクリプトにExcelする必要があります。

## <a name="see-also"></a>関連項目

- [スクリプトOfficeトラブルシューティング](troubleshooting.md)
- [Office スクリプトの効果を元に戻す](undo.md)
- [スクリプトのパフォーマンスをOfficeする](../develop/web-client-performance.md)
- [スクリプトの基本OfficeスクリプトのExcel on the web](../develop/scripting-fundamentals.md)
