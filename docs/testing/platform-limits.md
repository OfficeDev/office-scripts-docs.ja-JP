---
title: Office Scripts を使用したプラットフォームの制限と要件
description: Excel on the Web で使用する場合の Office Scripts のリソース制限とブラウザーのサポート。
ms.date: 11/07/2022
ms.localizationpriority: medium
ms.openlocfilehash: 764d1eddaf303a941a098ec1d3f3056d63e8693f
ms.sourcegitcommit: 7cadf2b637bf62874e43b6e595286101816662aa
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 11/09/2022
ms.locfileid: "68891247"
---
# <a name="platform-limits-and-requirements-with-office-scripts"></a>Office Scripts を使用したプラットフォームの制限と要件

Office Scripts を開発する際に注意する必要があるプラットフォームの制限がいくつかあります。 この記事では、Office Scripts for Excel on the Web のブラウザーサポートとデータ制限について詳しく説明します。

## <a name="browser-support"></a>ブラウザのサポート

Office Scripts は、[Office for the Web をサポート](https://support.microsoft.com/office/ad1303e0-a318-47aa-b409-d3a5eb44e452)する任意のブラウザーで動作します。 ただし、一部の JavaScript 機能は Internet Explorer 11 (IE 11) ではサポートされていません。 [ES6 以降](https://www.w3schools.com/Js/js_es6.asp)で導入された機能は、IE 11 では機能しません。 組織内のユーザーが引き続きそのブラウザーを使用している場合は、共有するときにその環境でスクリプトをテストしてください。

[!INCLUDE [Teams support note](../includes/teams-support-note.md)]

### <a name="third-party-cookies"></a>サード パーティの Cookie

Excel on the Web で **[自動化]** タブを表示するには、ブラウザーでサード パーティの Cookie が有効になっている必要があります。 タブが表示されていない場合は、ブラウザーの設定を確認します。 プライベート ブラウザー セッションを使用している場合は、毎回この設定を再度有効にする必要があります。

> [!NOTE]
> 一部のブラウザーでは、"サード パーティの Cookie" ではなく、この設定を "すべての Cookie" と呼んでいます。

#### <a name="instructions-for-adjusting-cookie-settings-in-popular-browsers"></a>一般的なブラウザーで Cookie 設定を調整する手順

- [Chrome](https://support.google.com/chrome/answer/95647)
- [Edge](https://support.microsoft.com/microsoft-edge/597f04f2-c0ce-f08c-7c2b-541086362bd2)
- [Firefox](https://support.mozilla.org/kb/disable-third-party-cookies)
- [Safari](https://support.apple.com/guide/safari/manage-cookies-and-website-data-sfri11471/mac)

## <a name="data-limits"></a>データの上限

一度に転送できる Excel データの量と、実行できる個々の Power Automate トランザクションの数には制限があります。

### <a name="excel"></a>Excel

Excel for the Web には、スクリプトを使用して Workbook を呼び出すときに、次の制限があります:

- 要求と応答は **5 MB** に制限されています。
- 範囲は **500 万セル** に制限されます。

大規模なデータセットを処理するときにエラーが発生する場合は、より大きな範囲ではなく、複数の小さい範囲を使用してみてください。 例としては、「[大規模なデータセットの書き込み](../resources/samples/write-large-dataset.md)」のサンプルを参照してください。 [Range.getSpecialCells](/javascript/api/office-scripts/excelscript/excelscript.range#excelscript-excelscript-range-getspecialcells-member(1)) などの API を使用して、大きな範囲ではなく特定のセルをターゲットにすることもできます。

Office スクリプトに固有ではない Excel の制限については、 [Excel の仕様と制限](https://support.microsoft.com/office/excel-specifications-and-limits-1672b34d-7043-467e-8e27-269d656771c3)に関する記事を参照してください。

### <a name="power-automate"></a>Power Automate

Power Automate で Office Scripts を使用する場合、各ユーザーは **1 日あたり 1,600 回のRun Script アクションの呼び出し** に制限されます。 この制限は、UTC の午前 12:00 にリセットされます。

Power Automate プラットフォームには、次の記事に記載されている使用制限もあります。

- [Power Automate における制限と構成](/power-automate/limits-and-config)
- [Excel Online (Business) コネクタの既知の問題と制限事項](/connectors/excelonlinebusiness/#known-issues-and-limitations)

> [!NOTE]
> 実行時間の長いスクリプトがある場合は、[同期的 Power Automate 操作における 120 秒のタイムアウト](/power-automate/limits-and-config#timeout)に注意してください。 [スクリプトを最適化](../develop/web-client-performance.md)するか、Excel の自動化を複数のスクリプトに分割する必要があります。

## <a name="see-also"></a>関連項目

- [Excel の仕様と制限](https://support.microsoft.com/office/excel-specifications-and-limits-1672b34d-7043-467e-8e27-269d656771c3)
- [Office Scripts のトラブルシューティング](troubleshooting.md)
- [Office スクリプトの効果を元に戻す](undo.md)
- [Office Scripts のパフォーマンスの改善](../develop/web-client-performance.md)
