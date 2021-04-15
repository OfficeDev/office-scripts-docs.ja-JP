---
title: Office スクリプトと Office アドインの違い
description: スクリプトとアドインの動作Office API のOffice違い。
ms.date: 06/01/2020
localization_priority: Normal
ms.openlocfilehash: 96af98ca9f247406c5cc916f38892c318d33c560
ms.sourcegitcommit: 45ffe3dbd2c834b78592ad35928cf8096f5e80bc
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 04/14/2021
ms.locfileid: "51755099"
---
# <a name="differences-between-office-scripts-and-office-add-ins"></a>Office スクリプトと Office アドインの違い

OfficeアドインとスクリプトOffice共通点が多い。 どちらも JavaScript API の Excel ブックの自動制御を提供します。 ただし、Officeスクリプト API は、JavaScript API の特殊な同期Officeです。

:::image type="content" source="../images/office-programmability-diagram.png" alt-text="さまざまな拡張ソリューションのフォーカス領域を示す 4 象限Office図。Office スクリプトと Office Web アドインの両方が Web とコラボレーションに焦点を当てしていますが、Office スクリプトはエンド ユーザーに対応します (Office Web アドインはプロフェッショナル開発者を対象としています)。":::

Officeスクリプトは、手動ボタンを押して完了するか [、Power Automate](https://flow.microsoft.com/)のステップとして実行しますが、Office アドインは作業ウィンドウが開いている間も保持されます。 つまり、アドインはセッション中に状態を維持できるのに対し、Officeスクリプトは実行の間に内部状態を維持できません。 Excel 拡張機能がスクリプト プラットフォームの機能を超える必要がある場合は [、Office](/office/dev/add-ins) アドインのドキュメントを参照して、Office アドインの詳細を確認してください。

この記事の残りの部分では、アドインとスクリプトの主な違OfficeについてOfficeします。

## <a name="platform-support"></a>プラットフォームサポート

Officeはクロスプラットフォームです。 Windows デスクトップ、Mac、iOS、および Web プラットフォーム間で動作し、それぞれに同じエクスペリエンスを提供します。 この例外は、個々の API のドキュメントに示されています。

Officeスクリプトは現在、Web 上の Excel でのみサポートされています。 すべての記録、編集、および実行は、Web プラットフォーム上で行われます。

## <a name="apis"></a>API

アドイン用の JavaScript API Office同期バージョンOfficeはありません。標準のOfficeスクリプト API はプラットフォームに固有であり、パラダイムの使用を避けるための多数の最適化と変更 `load` / `sync` があります。

[Excel JavaScript API の](/javascript/api/excel?view=excel-js-preview&preserve-view=true)一部は、スクリプト非同期 API Office[互換性があります](../develop/excel-async-model.md)。 一部のサンプルとアドイン コード ブロックは、最小限の変換でブロック `Excel.run` に移植できます。 2 つのプラットフォームは機能を共有しますが、ギャップがあります。 2 つの主要な API セットは、Officeが含まれていますが、スクリプトOfficeイベントと共通 API ではありません。

### <a name="events"></a>イベント

Officeスクリプトはイベントをサポート [していない](/office/dev/add-ins/excel/excel-add-ins-events)。 すべてのスクリプトでコードが 1 つのメソッドで `main` 実行され、終了します。 イベントがトリガーされると再アクティブ化されないので、イベントを登録できません。

### <a name="common-apis"></a>共通 API

Officeは共通 [API を使用できません](/javascript/api/office)。 一般的な API でのみサポートされている認証、ダイアログ ウィンドウ、その他の機能が必要な場合は、Office スクリプトではなく Office アドインを作成する必要があります。

## <a name="see-also"></a>関連項目

- [Excel on the web の Office スクリプト](../overview/excel.md)
- [スクリプトと VBA Officeの違い](vba-differences.md)
- [Office スクリプトのトラブルシューティング](../testing/troubleshooting.md)
- [Excel 作業ウィンドウ アドインを作成する](/office/dev/add-ins/quickstarts/excel-quickstart-jquery)
