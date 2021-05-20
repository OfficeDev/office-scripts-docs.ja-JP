---
title: Office スクリプトと VBA マクロの違い
description: OfficeスクリプトとVBAマクロの動作とAPIの違いExcel。
ms.date: 05/18/2021
localization_priority: Normal
ms.openlocfilehash: 612a5f21d935fd262a6e9fd12a3431956105636a
ms.sourcegitcommit: 4687693f02fc90a57ba30c461f35046e02e6f5fb
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 05/19/2021
ms.locfileid: "52545589"
---
# <a name="differences-between-office-scripts-and-vba-macros"></a>Office スクリプトと VBA マクロの違い

Officeスクリプトと VBA マクロには多くの共通点があります。 どちらも、ユーザーが使いやすいアクション レコーダを使用してソリューションを自動化し、それらの記録の編集を可能にします。 どちらのフレームワークも、プログラマとは思わない人がExcelで小さなプログラムを作成する権限を与えるように設計されています。
基本的な違いは、VBA マクロはデスクトップ ソリューション用に開発され、スクリプトは安全なクラウドベースのソリューション用に設計Officeです。 現在、OfficeスクリプトはExcel on the webでのみサポートされています。

:::image type="content" source="../images/office-programmability-diagram.png" alt-text="さまざまなOffice機能拡張ソリューションの対象領域を示す 4 つの作業領域図。Officeスクリプトと VBA マクロは、エンド ユーザーがソリューションを作成できるように設計されていますが、Office スクリプトは Web と共同作業用に構築されています (VBA はデスクトップ用です)。":::

この資料では、VBA マクロ (一般的には VBA) とOffice スクリプトの主な違いについて説明します。 OfficeスクリプトはExcelでのみ使用できるため、ここで説明する唯一のホストです。

## <a name="platform-and-ecosystem"></a>プラットフォームとエコシステム

VBA はデスクトップ用に設計されており、スクリプトは web 用に設計Office。 VBA は、ユーザーのデスクトップと対話して、COM や OLE などの類似のテクノロジに接続できます。 しかし、VBAはインターネットに呼び出す便利な方法はありません。

Officeスクリプトは、JavaScript のユニバーサル ランタイムを使用します。 これにより、スクリプトの実行に使用するマシンに関係なく、一貫した動作とアクセシビリティが得られます。 また、他の Web サービスを呼び出すこともできます。

## <a name="security"></a>セキュリティ

VBA マクロには、Excelと同じセキュリティクリアランスがあります。 これにより、デスクトップにフル アクセスできます。 Officeスクリプトは、ブックへのアクセス権のみを持ち、ブックをホストしているコンピューターにはアクセスできません。 さらに、スクリプトと JavaScript 認証トークンを共有することはできません。 つまり、スクリプトにはサインインしているユーザーのトークンも、外部サービスにサインインするための API 機能もないので、既存のトークンを使用してユーザーに代わって外部呼び出しを行うことができません。

管理者には、テナント上のすべてのマクロを許可する、テナント上のすべてのマクロを許可しない、署名付き証明書を持つマクロのみを許可する、という 3 つの VBA マクロのオプションがあります。 この粒度の欠如は、単一の悪いアクターを分離することが困難になります。 現在、Officeスクリプトはテナントのオンまたはオフのいずれかです。 ただし、管理者は個々のスクリプトやスクリプト作成者をより詳細に制御するように取り組んでいます。

## <a name="coverage"></a>割合

現在、VBA は、特にデスクトップ クライアントで利用できる機能Excel詳細に提供しています。 Officeスクリプトは、Excel on the webのほぼすべてのシナリオをカバーします。 さらに、新しい機能が Web 上でデビューすると、Officeスクリプトはアクション レコーダと JavaScript API の両方でサポートします。

OfficeスクリプトはExcelレベルの[イベント](/office/vba/excel/concepts/events-worksheetfunctions-shapes/using-events-with-excel-objects)をサポートしていません。 スクリプトは、ユーザーが手動でスクリプトを開始した場合、またはPower Automateフローがスクリプトを呼び出したときにのみ実行されます。

## <a name="power-automate"></a>Power Automate

OfficeスクリプトはPower Automateを通して実行できます。 ブックはスケジュールフローまたはイベントドリブンフローで更新できるため、Excelを開くことなくワークフローを自動化できます。 つまり、ワークブックがOneDrive(およびPower Automateにアクセス可能)に保存されている限り、ユーザーと組織がExcelのデスクトップ、Mac、または Web クライアントを使用しているかどうかに関係なく、フローでスクリプトを実行できます。

VBA にはPower Automateコネクタがありません。 サポートされているすべての VBA シナリオには、マクロの実行に参加するユーザーが含まれます。

Power Automateの学習を開始するには[、手動のPower Automateフローチュートリアルからスクリプトを呼び出す](../tutorials/excel-power-automate-manual.md)を試してみてください。 また、[自動化タスクのリマインダー](scenarios/task-reminders.md)サンプルを確認して、実際のシナリオでTeamsに接続されたOfficeスクリプトをPower Automateすることもできます。

## <a name="see-also"></a>関連項目

- [Excel on the web の Office スクリプト](../overview/excel.md)
- [Power Automate を使用した Office スクリプトの実行](../develop/power-automate-integration.md)
- [Office スクリプトと Office アドインの違い](add-ins-differences.md)
- [Office スクリプトのトラブルシューティング](../testing/troubleshooting.md)
- [Excel VBA リファレンス](/office/vba/api/overview/excel)
