---
title: Office スクリプトと Office アドインの違い
description: スクリプトとアドインの動作Office API Office違い。
ms.date: 02/04/2022
ms.localizationpriority: medium
ms.openlocfilehash: 018d210208bc78da894678d21e368864522cb83e
ms.sourcegitcommit: 7023b9e23499806901a5ecf8ebc460b76887cca6
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 03/31/2022
ms.locfileid: "64585610"
---
# <a name="differences-between-office-scripts-and-office-add-ins"></a>Office スクリプトと Office アドインの違い

各スクリプトを使用するOfficeとOfficeアドインの違いを理解します。 Officeスクリプトは、ワークフローを改善するために必要なユーザーが迅速に作成するように設計されています。 Officeアドインは、リボン ボタンOffice作業ウィンドウを通じて、より対話的なエクスペリエンスを提供する UI と統合します。 Officeアドインは、カスタム関数を提供することでExcel機能を拡張できます。

:::image type="content" source="../images/office-programmability-diagram.png" alt-text="さまざまな拡張ソリューションのフォーカス領域を示す 4 象限Office図。Office スクリプトと Office Web アドインの両方が Web とコラボレーションに焦点を当て、Office スクリプトはエンド ユーザーに対応します (一方、Office Web アドインはプロフェッショナル開発者を対象としています)。":::

Officeスクリプトは、手動でボタンを押したり[、Power Automate](https://flow.microsoft.com/) のステップとして実行したりしますが、Office アドインは構成方法に応じて実行を続けています。 たとえば、作業ウィンドウが閉Office実行を続行するアドインを構成できます。 つまり、Officeアドインはセッション中に状態を維持しますが、Officeスクリプトは実行の間に内部状態を維持します。 構築するソリューションで保守された状態が必要な場合は、Office アドインの[](/office/dev/add-ins)ドキュメントを参照して、Office アドインの詳細を確認する必要があります。

この記事の残りの部分では、アドインとスクリプトの主な違OfficeについてOfficeします。

## <a name="platform-support"></a>プラットフォームサポート

Officeはクロスプラットフォームです。 デスクトップ、Mac、Windows Web プラットフォーム間で動作し、それぞれで同じエクスペリエンスを提供します。 この例外は、個々の API のドキュメントに示されています。

Officeスクリプトは現在、ユーザーがサポートしているExcel on the web。 すべての記録、編集、およびスクリプト管理は、Web プラットフォーム上で行われます。

### <a name="script-support-for-excel-on-windows"></a>スクリプトは、ExcelのWindows

[!INCLUDE [Run-from-button support](../includes/run-from-button-desktop-support.md)]

## <a name="apis"></a>API

Office用の JavaScript API Office Office スクリプト API はいくつかの機能を共有しますが、それらは異なるプラットフォームです。 スクリプト Office API は、JavaScript API モデルの最適化された同期Excelサブセットです。 大きな違いは、アドインでの`load`/`sync`パラダイムの使用です。さらに、アドインはイベント用の API と、共通 API と呼ばれる Excel 外部のより広範な機能セットを提供します。

### <a name="events"></a>イベント

Officeスクリプトは、ブック レベルのイベントをサポート[していない](/office/dev/add-ins/excel/excel-add-ins-events)。 スクリプトは、スクリプトの [実行] ボタンを選択するか、スクリプトを使用して実行Power Automate。 すべてのスクリプトでコードが 1 つのメソッドで実行 `main` され、終了します。

### <a name="common-apis"></a>共通 API

Officeは共通 [API を使用できません](/javascript/api/office)。 一般的な API でのみサポートされている認証、ダイアログ ウィンドウ、その他の機能が必要な場合は、Office スクリプトではなく Office アドインを作成する必要があります。

## <a name="see-also"></a>関連項目

- [Excel on the web の Office スクリプト](../overview/excel.md)
- [スクリプトと VBA Officeの違い](vba-differences.md)
- [Office スクリプトのトラブルシューティング](../testing/troubleshooting.md)
- [Excel 作業ウィンドウ アドインを作成する](/office/dev/add-ins/quickstarts/excel-quickstart-jquery)
