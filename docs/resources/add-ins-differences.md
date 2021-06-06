---
title: Office スクリプトと Office アドインの違い
description: スクリプトとアドインの動作Office API Office違い。
ms.date: 06/02/2021
localization_priority: Normal
ms.openlocfilehash: 46f5f2ea6fea15e9506f5c7d30941311fc2e669e
ms.sourcegitcommit: 0bfc9472d107e32c804029659317f8e81fec5d19
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 06/05/2021
ms.locfileid: "52779364"
---
# <a name="differences-between-office-scripts-and-office-add-ins"></a>Office スクリプトと Office アドインの違い

各スクリプトと OfficeアドインOfficeの違いを理解し、各アドインをいつ使用する必要が生じ得るのかについて理解します。 Officeスクリプトは、ワークフローの改善を探しているすべてのユーザーが迅速に作成するように設計されています。 Officeアドインは、リボン ボタンと作業ウィンドウOffice対話型の UI と統合します。 Officeアドインは、カスタム関数を提供することで、組み込Excel機能を拡張できます。

:::image type="content" source="../images/office-programmability-diagram.png" alt-text="さまざまな拡張ソリューションのフォーカス領域を示す 4 象限Office図。Office スクリプトと Office Web アドインの両方が Web とコラボレーションに焦点を当て、Office スクリプトはエンド ユーザーに対応します (一方、Office Web アドインはプロの開発者を対象とします)":::

Officeスクリプトは手動でボタンを押して実行するか[、Power Automate](https://flow.microsoft.com/)でステップとして実行しますが、Office アドインは構成方法に応じて実行を続行します。 たとえば、作業ウィンドウが閉Office実行を続行するアドインを構成できます。 つまり、Officeアドインはセッション中に状態を維持しますが、Officeスクリプトは実行の間に内部状態を維持します。 構築するソリューションに保守状態が必要な場合は、Office アドインの[](/office/dev/add-ins)ドキュメントを参照して、Office アドインの詳細を確認する必要があります。

この記事の残りの部分では、アドインとスクリプトの主なOfficeについてOfficeします。

## <a name="platform-support"></a>プラットフォームサポート

Officeアドインはクロスプラットフォームです。 デスクトップ、Mac、Windows Web プラットフォーム間で動作し、それぞれで同じエクスペリエンスを提供します。 この例外は、個々の API のドキュメントに示されています。

Officeスクリプトは現在、ユーザーがサポートしているExcel on the web。 すべての記録、編集、および実行は、Web プラットフォーム上で行われます。

## <a name="apis"></a>API

OfficeアドインOffice Office スクリプト API の JavaScript API はいくつかの機能を共有しますが、プラットフォームは異なります。 スクリプト Office API は、JavaScript API モデルの最適化された同期Excelサブセットです。 大きな違いは、アドイン `load` / `sync` でのパラダイムの使用です。さらに、アドインはイベント用の API と、共通 API と呼ばれる Excel以外の広範な機能セットを提供します。

### <a name="events"></a>イベント

Officeスクリプトは、ブック レベルのイベントを[サポートしていない](/office/dev/add-ins/excel/excel-add-ins-events)。 スクリプトは、ユーザーがスクリプトの [**実行**] ボタンを押すか、スクリプトを使用してトリガー Power Automate。 すべてのスクリプトでコードが 1 つのメソッドで `main` 実行され、終了します。

### <a name="common-apis"></a>共通 API

Officeスクリプトで共通[API を使用することはできません](/javascript/api/office)。 一般的な API でのみサポートされている認証、ダイアログ ウィンドウ、その他の機能が必要な場合は、Office スクリプトではなく Office アドインを作成する必要があります。

## <a name="see-also"></a>関連項目

- [Excel on the web の Office スクリプト](../overview/excel.md)
- [スクリプトと VBA Officeの違い](vba-differences.md)
- [Office スクリプトのトラブルシューティング](../testing/troubleshooting.md)
- [Excel 作業ウィンドウ アドインを作成する](/office/dev/add-ins/quickstarts/excel-quickstart-jquery)
