---
title: Office スクリプトと Office アドインの相違点
description: Office スクリプトと Office アドインの動作と API の違い。
ms.date: 12/12/2019
localization_priority: Normal
ms.openlocfilehash: 4626afb66b54c94a72f29b039c601435c089d64d
ms.sourcegitcommit: b075eed5a6f275274fbbf6d62633219eac416f26
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 03/10/2020
ms.locfileid: "42700395"
---
# <a name="differences-between-office-scripts-and-office-add-ins"></a>Office スクリプトと Office アドインの相違点

Office アドインと Office スクリプトには、多くの共通点があります。 どちらも、Office JavaScript API の名前空間を`Excel`使用して、Excel ブックの自動制御を提供します。 ただし、Office スクリプトの範囲は、より制限されています。

Office スクリプトは、手動のボタンを押すことで完了まで実行されます。 Office アドインは、ユーザーの操作に依存し、ブックの使用中は保持されます。 Excel 拡張機能がスクリプトプラットフォームの機能を超える必要がある場合は、office アドインの[ドキュメント](/office/dev/add-ins)にアクセスして、office アドインの詳細を確認してください。

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

- [Web 上の Excel での Office スクリプト](../overview/excel.md)
- [Office スクリプトのトラブルシューティング](../testing/troubleshooting.md)
- [Excel 作業ウィンドウ アドインを作成する](/office/dev/add-ins/quickstarts/excel-quickstart-jquery)