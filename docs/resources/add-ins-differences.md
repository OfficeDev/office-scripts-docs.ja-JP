---
title: Office スクリプトと Office アドインの違い
description: Office スクリプトと Office アドインの動作と API の違い。
ms.date: 02/04/2022
ms.localizationpriority: medium
ms.openlocfilehash: a3df4daf04f963598d2cb31f82dd2c1c9923fdc8
ms.sourcegitcommit: 33fe0f6807daefb16b148fd73c863de101f47cea
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 08/08/2022
ms.locfileid: "67281911"
---
# <a name="differences-between-office-scripts-and-office-add-ins"></a>Office スクリプトと Office アドインの違い

Office スクリプトと Office アドインの違いを理解し、それぞれをいつ使用するかを把握します。 Office スクリプトは、ワークフローを改善しようとしているすべてのユーザーが迅速に作成できるように設計されています。 Office アドインは Office UI と統合され、リボン ボタンと作業ウィンドウを使用して、より対話型のエクスペリエンスを実現します。 Office アドインでは、カスタム関数を提供することで、組み込みの Excel 関数を拡張することもできます。

:::image type="content" source="../images/office-programmability-diagram.png" alt-text="さまざまな Office 機能拡張ソリューションのフォーカス領域を示す 4 象限図。Office スクリプトと Office Web アドインはどちらも Web とコラボレーションに焦点を当てていますが、Office スクリプトはエンド ユーザーに対応します (一方、Office Web アドインはプロの開発者を対象としています)。":::

Office スクリプトは手動ボタンを押すか [、Power Automate](https://flow.microsoft.com/) の手順で完了するまで実行されますが、Office アドインは構成方法に応じて引き続き実行されます。 たとえば、作業ウィンドウが閉じても引き続き実行するように Office アドインを構成できます。 つまり、Office アドインはセッション中に状態を維持しますが、Office スクリプトは実行間の内部状態を維持しません。 ビルドするソリューションに保守状態が必要な場合は、 [Office アドインのドキュメント](/office/dev/add-ins) を参照して、Office アドインの詳細を確認する必要があります。

この記事の残りの部分では、Office アドインと Office スクリプトの主な違いについて説明します。

## <a name="platform-support"></a>プラットフォームのサポート

Office アドインはクロスプラットフォームです。 Windows デスクトップ、Mac、iOS、Web プラットフォーム間で動作し、それぞれに同じエクスペリエンスを提供します。 これに対する例外については、個々の API のドキュメントに記載されています。

Office スクリプトは現在、Excel on the webでのみサポートされています。 すべての記録、編集、およびスクリプト管理は、Web プラットフォームで行われます。

### <a name="script-support-for-excel-on-windows"></a>Windows 上の Excel のスクリプトサポート

[!INCLUDE [Run-from-button support](../includes/run-from-button-desktop-support.md)]

## <a name="apis"></a>API

Office アドイン用の Office JavaScript API と Office スクリプト API は一部の機能を共有していますが、プラットフォームは異なります。 Office スクリプト API は、Excel JavaScript API モデルの最適化された同期サブセットです。 主な違いは、アドインでの `load`/`sync` パラダイムの使用です。さらに、アドインは、イベント用の API と、共通 API と呼ばれる Excel 以外の広範な機能セットを提供します。

### <a name="events"></a>Events

Office スクリプトでは、ブック レベルの [イベント](/office/dev/add-ins/excel/excel-add-ins-events)はサポートされていません。 スクリプトは、ユーザーがスクリプトの **[実行** ] ボタンを選択するか、Power Automate を使用してトリガーされます。 すべてのスクリプトは、1 つの `main` 関数でコードを実行し、終了します。

### <a name="common-apis"></a>共通 API

Office スクリプトでは [、一般的な API を](/javascript/api/office)使用できません。 一般的な API でのみサポートされている認証、ダイアログ ウィンドウ、その他の機能が必要な場合は、Office スクリプトの代わりに Office アドインを作成する必要があります。

## <a name="see-also"></a>関連項目

- [Excel の Office スクリプト](../overview/excel.md)
- [Office スクリプトと VBA マクロの違い](vba-differences.md)
- [Office スクリプトのトラブルシューティング](../testing/troubleshooting.md)
- [Excel 作業ウィンドウ アドインを作成する](/office/dev/add-ins/quickstarts/excel-quickstart-jquery)
