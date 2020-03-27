---
title: Office スクリプトと Office アドインの相違点
description: Office スクリプトと Office アドインの動作と API の違い。
ms.date: 03/23/2020
localization_priority: Normal
ms.openlocfilehash: 2290d4e34b7a7286d67443de9e9c64bad4fcd4b7
ms.sourcegitcommit: d556aaefac80e55f53ac56b7f6ecbc657ebd426f
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 03/26/2020
ms.locfileid: "42978729"
---
# <a name="differences-between-office-scripts-and-office-add-ins"></a>Office スクリプトと Office アドインの相違点

Office アドインと Office スクリプトには、多くの共通点があります。 どちらも、Office JavaScript API の名前空間を`Excel`使用して、Excel ブックの自動制御を提供します。 ただし、Office スクリプトの範囲は、より制限されています。

![さまざまな Office 機能拡張ソリューションのフォーカス領域を示す4つの領域の図。 Office スクリプトと Office Web アドインはどちらも Web とコラボレーションに重点が置いていますが、Office スクリプトはエンドユーザーに対して機能します (ただし、Office Web アドインでは、プロフェッショナル開発者が対象となります)。)](../images/office-programmability-diagram.png)

Office スクリプトは、作業ウィンドウが開いている間は Office アドインが保持されるのに対して、手動ボタンを押すか、[電源自動化](https://flow.microsoft.com/)で手順として、完了するために実行します。 これは、アドインがセッション中に状態を維持できることを意味しますが、Office スクリプトでは実行の間に内部状態は保持されません。 Excel 拡張機能がスクリプトプラットフォームの機能を超える必要がある場合は、office アドインの[ドキュメント](/office/dev/add-ins)にアクセスして、office アドインの詳細を確認してください。

この記事の残りの部分では、Office アドインと Office スクリプトの主な違いについて説明します。

## <a name="platform-support"></a>プラットフォームのサポート

Office アドインはプラットフォーム間で機能します。 これらは、Windows デスクトップ、Mac、iOS、および web プラットフォーム間で機能し、それぞれに同じ操作を提供します。 この点については、個々の API のドキュメントに記載されている例外を参照してください。

Office スクリプトは、現在 web 上の Excel でのみサポートされています。 すべての記録、編集、実行は、web プラットフォーム上で実行されます。

## <a name="apis"></a>API

Office スクリプトは、ほとんどの Excel JavaScript Api をサポートしています。これは、2つのプラットフォーム間で多くの機能が重なっていることを意味します。 2つの例外として、イベントと共通 Api があります。

### <a name="events"></a>イベント

Office スクリプトは[イベント](/office/dev/add-ins/excel/excel-add-ins-events)をサポートしていません。 すべてのスクリプトは、コードを 1 `main`つのメソッドで実行し、終了します。 イベントがトリガーされると再アクティブ化されないため、イベントを登録できません。

### <a name="common-apis"></a>共通 API

Office スクリプトでは、[共通 api](/javascript/api/office)を使用できません。 一般的な Api でのみサポートされている認証、ダイアログウィンドウ、またはその他の機能が必要な場合は、Office のスクリプトではなく、Office アドインを作成する必要があります。

## <a name="see-also"></a>関連項目

- [Excel on the web の Office スクリプト](../overview/excel.md)
- [Office スクリプトと VBA マクロの相違点](vba-differences.md)
- [Office スクリプトのトラブルシューティング](../testing/troubleshooting.md)
- [Excel 作業ウィンドウ アドインを作成する](/office/dev/add-ins/quickstarts/excel-quickstart-jquery)
