---
title: Office スクリプトと Office アドインの違い
description: スクリプトとアドインの動作Office API Office違い。
ms.date: 06/01/2020
localization_priority: Normal
ms.openlocfilehash: 5c30406867da05952dedda684f765df5e7a7e53f
ms.sourcegitcommit: 09d8859d5269ada8f1d0e141f6b5a4f96d95a739
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 05/24/2021
ms.locfileid: "52631679"
---
# <a name="differences-between-office-scripts-and-office-add-ins"></a>Office スクリプトと Office アドインの違い

Officeアドインとカスタム スクリプトOffice共通点が多い。 どちらも JavaScript API を使用してブックExcel制御を提供します。 ただし、Officeスクリプト API は、JavaScript API の特殊な同期Officeです。

:::image type="content" source="../images/office-programmability-diagram.png" alt-text="さまざまな拡張ソリューションのフォーカス領域を示す 4 象限Office図。Office スクリプトと Office Web アドインの両方が Web とコラボレーションに焦点を当て、Office スクリプトはエンド ユーザーに対応します (一方、Office Web アドインはプロの開発者を対象とします)":::

Officeスクリプトは、手動ボタンを押して完了するか[、Power Automate](https://flow.microsoft.com/)でステップとして実行しますが、Office アドインは作業ウィンドウを開いている間も保持されます。 つまり、アドインはセッション中に状態を維持できるのに対し、Officeスクリプトは実行の間に内部状態を維持できません。 Excel 拡張機能がスクリプト プラットフォームの機能を超える必要がある場合は、Office アドインのドキュメントを参照して[、Office](/office/dev/add-ins)アドインの詳細を確認してください。

この記事の残りの部分では、アドインとスクリプトの主なOfficeについてOfficeします。

## <a name="platform-support"></a>プラットフォームサポート

Officeアドインはクロスプラットフォームです。 デスクトップ、Mac、Windows Web プラットフォーム間で動作し、それぞれで同じエクスペリエンスを提供します。 この例外は、個々の API のドキュメントに示されています。

Officeスクリプトは現在、ユーザーがサポートしているExcel on the web。 すべての記録、編集、および実行は、Web プラットフォーム上で行われます。

## <a name="apis"></a>API

OfficeアドインOffice Office スクリプト API の JavaScript API はいくつかの機能を共有しますが、プラットフォームは異なります。 スクリプト Office API は、JavaScript API モデルの最適化された同期Excelバージョンです。 大きな違いは、アドイン `load` / `sync` でのパラダイムの使用です。さらに、アドインはイベント用の API と、共通 API と呼ばれる Excel以外の広範な機能セットを提供します。

### <a name="events"></a>イベント

Officeスクリプトはイベントをサポート[していない](/office/dev/add-ins/excel/excel-add-ins-events)。 すべてのスクリプトでコードが 1 つのメソッドで `main` 実行され、終了します。 イベントがトリガーされると再アクティブ化されないので、イベントを登録できません。

### <a name="common-apis"></a>共通 API

Officeスクリプトで共通[API を使用することはできません](/javascript/api/office)。 一般的な API でのみサポートされている認証、ダイアログ ウィンドウ、その他の機能が必要な場合は、Office スクリプトではなく Office アドインを作成する必要があります。

## <a name="see-also"></a>関連項目

- [Excel on the web の Office スクリプト](../overview/excel.md)
- [スクリプトと VBA Officeの違い](vba-differences.md)
- [Office スクリプトのトラブルシューティング](../testing/troubleshooting.md)
- [Excel 作業ウィンドウ アドインを作成する](/office/dev/add-ins/quickstarts/excel-quickstart-jquery)
