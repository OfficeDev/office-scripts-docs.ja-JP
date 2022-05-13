---
title: Office スクリプトと VBA マクロの違い
description: Office スクリプトと Excel VBA マクロの動作と API の違い。
ms.date: 02/04/2022
ms.localizationpriority: medium
ms.openlocfilehash: 60e4fba6e63967302066f544b76fb20a8c8630a6
ms.sourcegitcommit: 34c7740c9bff0e4c7426e01029f967724bfee566
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 05/13/2022
ms.locfileid: "65393615"
---
# <a name="differences-between-office-scripts-and-vba-macros"></a>Office スクリプトと VBA マクロの違い

Office スクリプトと VBA マクロには、多くの共通点があります。 どちらも、ユーザーが使いやすいアクション レコーダーを使用してソリューションを自動化し、それらの記録を編集できるようにします。 どちらのフレームワークも、プログラマと思えない人がExcelで小さなプログラムを作成できるように設計されています。

基本的な違いは、VBA マクロはデスクトップ ソリューション用に開発され、Office スクリプトはセキュリティで保護されたクラウドベースのソリューション用に設計されていることです。 現在、Office スクリプトはExcel on the webでのみサポートされています。

:::image type="content" source="../images/office-programmability-diagram.png" alt-text="さまざまなOffice機能拡張ソリューションのフォーカス領域を示す 4 象限図。Office スクリプトと VBA マクロはどちらもエンド ユーザーがソリューションを作成できるように設計されていますが、Office スクリプトは Web とコラボレーション用に構築されています (一方、VBA はデスクトップ用です)。":::

この記事では、VBA マクロ (および VBA 全般) と Office スクリプトの主な違いについて説明します。 Office スクリプトはExcelでのみ使用できるため、ここで説明する唯一のホストです。

## <a name="platform-and-ecosystem"></a>プラットフォームとエコシステム

VBA は、Windowsおよび Mac 上のExcelでサポートされています。 Office スクリプトはExcel on the webでサポートされています。

2 つのソリューションは、それぞれのプラットフォーム向けに設計されています。 VBA はユーザーのデスクトップと対話して、COM や OLE などの同様のテクノロジに接続できます。 ただし、VBA ではインターネットに発信する便利な方法はありません。 Office スクリプトでは、JavaScript 用のユニバーサル ランタイムが使用されます。 これにより、スクリプトの実行に使用されているマシンに関係なく、一貫した動作とアクセシビリティが提供されます。 また、他の Web サービスを呼び出すこともできます。

### <a name="script-support-for-excel-on-windows"></a>WindowsでのExcelのスクリプトサポート

[!INCLUDE [Run-from-button support](../includes/run-from-button-desktop-support.md)]

## <a name="security"></a>セキュリティ

VBA マクロには、Excelと同じセキュリティ上のスペースがあります。 これにより、デスクトップに完全にアクセスできます。 Office スクリプトは、ブックをホストしているコンピューターではなく、ブックにのみアクセスできます。 さらに、JavaScript 認証トークンをスクリプトと共有することはできません。 つまり、スクリプトにはサインインしているユーザーのトークンも外部サービスにサインインするための API 機能もないため、既存のトークンを使用してユーザーに代わって外部呼び出しを行うことができません。

管理者には、VBA マクロの 3 つのオプションがあります。テナント上のすべてのマクロを許可するか、テナントでマクロを許可しないか、署名された証明書を持つマクロのみを許可します。 この粒度の欠如により、単一の不良アクターを分離するのが困難になります。 現在、Office スクリプトは、テナント全体、テナント全体、またはテナント内のユーザーのグループに対してオフにすることができます。 管理者は、他のユーザーとスクリプトを共有できるユーザーと、Power Automateでスクリプトを使用できるユーザーも制御できます。

## <a name="coverage"></a>割合

現在、VBA では、Excel機能 (特にデスクトップ クライアントで使用可能な機能) のより完全なカバレッジが提供されています。 Office スクリプトは、Excel on the webのほぼすべてのシナリオを網羅しています。 さらに、新機能が Web で初めて登場すると、Office スクリプトはアクション レコーダーと JavaScript API の両方でサポートされます。

Office スクリプトでは、Excel レベルの[イベント](/office/vba/excel/concepts/events-worksheetfunctions-shapes/using-events-with-excel-objects)はサポートされていません。 スクリプトは、ユーザーが手動で起動したとき、またはPower Automate フローがスクリプトを呼び出すときにのみ実行されます。

## <a name="power-automate"></a>Power Automate

Office スクリプトは、Power Automateで実行できます。 ブックは、スケジュールされたフローまたはイベント ドリブン フローを使用して更新できるため、Excelを開くことなくワークフローを自動化できます。 つまり、ブックがOneDriveに保存され (Power Automateからアクセスできる) 限り、自分と組織がデスクトップ、Mac、または Web クライアントを使用するかどうかに関係なく、フローでスクリプト Excelを実行できます。

VBA にはPower Automate コネクタがありません。 サポートされているすべての VBA シナリオには、マクロの実行に参加するユーザーが含まれます。

手動の[Power Automate フローのチュートリアルからスクリプトを呼び出](../tutorials/excel-power-automate-manual.md)して、Power Automateの学習を開始してください。 [自動タスクリマインダー](scenarios/task-reminders.md)サンプルを確認して、実際のシナリオでPower Automateを介してTeamsに接続されたOfficeスクリプトを確認することもできます。

## <a name="see-also"></a>関連項目

- [ExcelでスクリプトをOfficeする](../overview/excel.md)
- [Power Automate を使用した Office スクリプトの実行](../develop/power-automate-integration.md)
- [Office スクリプトと Office アドインの違い](add-ins-differences.md)
- [Office スクリプトのトラブルシューティング](../testing/troubleshooting.md)
- [Excel VBA リファレンス](/office/vba/api/overview/excel)
