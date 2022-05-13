---
title: Office スクリプトと Office アドインの違い
description: Office スクリプトとOffice アドインの動作と API の違い。
ms.date: 02/04/2022
ms.localizationpriority: medium
ms.openlocfilehash: bd483f928e3e153b8a08537f6b333c3ea8d724dd
ms.sourcegitcommit: 34c7740c9bff0e4c7426e01029f967724bfee566
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 05/13/2022
ms.locfileid: "65393622"
---
# <a name="differences-between-office-scripts-and-office-add-ins"></a>Office スクリプトと Office アドインの違い

Office スクリプトとOffice アドインの違いを理解し、それぞれをいつ使用するかを把握します。 Office スクリプトは、ワークフローを改善しようとしているすべてのユーザーが迅速に作成できるように設計されています。 Office アドインはOffice UI と統合され、リボン ボタンや作業ウィンドウを使用して、よりインタラクティブなエクスペリエンスを実現します。 Office アドインは、カスタム関数を提供することで、組み込みのExcel関数を拡張することもできます。

:::image type="content" source="../images/office-programmability-diagram.png" alt-text="さまざまなOffice機能拡張ソリューションのフォーカス領域を示す 4 象限図。Office スクリプトとOffice Web アドインはどちらも Web とコラボレーションに焦点を当てていますが、Office スクリプトはエンド ユーザーに対応します (一方、Office Web アドインはプロの開発者を対象としています)。":::

Officeスクリプトは手動ボタンを押すか[、Power Automate](https://flow.microsoft.com/)のステップとして実行されますが、Officeアドインは構成方法に応じて引き続き実行されます。 たとえば、Office アドインを構成して、作業ウィンドウが閉じても引き続き実行するようにすることができます。 つまり、Office アドインはセッション中に状態を維持しますが、Office スクリプトは実行間の内部状態を保持しません。 ビルドするソリューションに保守状態が必要な場合は、[Office アドインのドキュメント](/office/dev/add-ins)にアクセスして、アドインのOfficeの詳細を確認する必要があります。

この記事の残りの部分では、Office アドインとOffice スクリプトの主な違いについて説明します。

## <a name="platform-support"></a>プラットフォームのサポート

Office アドインはクロスプラットフォームです。 デスクトップ、Mac、iOS、および Web プラットフォームWindowsまたがって動作し、それぞれに同じエクスペリエンスを提供します。 これに対する例外については、個々の API のドキュメントに記載されています。

Office スクリプトは現在、Excel on the webでのみサポートされています。 すべての記録、編集、およびスクリプト管理は、Web プラットフォームで行われます。

### <a name="script-support-for-excel-on-windows"></a>WindowsでのExcelのスクリプトサポート

[!INCLUDE [Run-from-button support](../includes/run-from-button-desktop-support.md)]

## <a name="apis"></a>API

Office アドイン用の javaScript API と Office スクリプト API のOfficeは、いくつかの機能を共有しますが、プラットフォームは異なります。 Office スクリプト API は、Excel JavaScript API モデルの最適化された同期サブセットです。 主な違いは、アドインでの`load`/`sync`パラダイムの使用です。さらに、アドインでは、イベント用の API と、共通 API と呼ばれるExcel以外の広範な機能セットが提供されます。

### <a name="events"></a>イベント

Office スクリプトでは、ブック レベルの[イベント](/office/dev/add-ins/excel/excel-add-ins-events)はサポートされていません。 スクリプトは、ユーザーがスクリプトの **[実行**] ボタンを選択するか、Power Automateを介してトリガーされます。 すべてのスクリプトは、1 つの `main` メソッドでコードを実行してから終了します。

### <a name="common-apis"></a>共通 API

Office スクリプトでは[共通 API を](/javascript/api/office)使用できません。 Common API でのみサポートされている認証、ダイアログ ウィンドウ、またはその他の機能が必要な場合は、Office スクリプトではなく、Office アドインを作成する必要があります。

## <a name="see-also"></a>関連項目

- [ExcelでスクリプトをOfficeする](../overview/excel.md)
- [Office スクリプトと VBA マクロの違い](vba-differences.md)
- [Office スクリプトのトラブルシューティング](../testing/troubleshooting.md)
- [Excel 作業ウィンドウ アドインを作成する](/office/dev/add-ins/quickstarts/excel-quickstart-jquery)
