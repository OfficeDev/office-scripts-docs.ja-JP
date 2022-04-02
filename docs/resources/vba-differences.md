---
title: スクリプトと VBA Officeの違い
description: スクリプトと VBA マクロの動作Office API Excel違い。
ms.date: 02/04/2022
ms.localizationpriority: medium
ms.openlocfilehash: 53cd2d9b163a3d3c3f9ac9196b5f5126b539611a
ms.sourcegitcommit: 7023b9e23499806901a5ecf8ebc460b76887cca6
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 03/31/2022
ms.locfileid: "64586018"
---
# <a name="differences-between-office-scripts-and-vba-macros"></a>スクリプトと VBA Officeの違い

Office VBA マクロには多くの共通点があります。 どちらも、ユーザーが使いやすいアクション レコーダーを使用してソリューションを自動化し、それらの記録の編集を許可します。 どちらのフレームワークも、プログラマが自分を考慮しない可能性があるユーザーが、ユーザーに小さなプログラムを作成Excel。

基本的な違いは、VBA マクロがデスクトップ ソリューション用に開発され、スクリプトOffice、セキュリティで保護されたクラウドベースのソリューション用に設計されている点です。 現在、Officeスクリプトは、このスクリプトでのみExcel on the web。

:::image type="content" source="../images/office-programmability-diagram.png" alt-text="さまざまな拡張ソリューションのフォーカス領域を示す 4 象限Office図。Officeスクリプトと VBA マクロの両方は、エンド ユーザーがソリューションを作成するのに役立つ設計ですが、Office スクリプトは Web とコラボレーション用に構築されています (一方、VBA はデスクトップ用です)。":::

この記事では、VBA マクロ (および VBA 全般) とスクリプトの主な違いOffice説明します。 このOfficeスクリプトは、Excelでしか使用できないので、ここで説明する唯一のホストです。

## <a name="platform-and-ecosystem"></a>プラットフォームとエコシステム

VBA は、Excel Mac Windowsによってサポートされます。 Officeスクリプトは、ユーザーがサポートExcel on the web。

2 つのソリューションは、それぞれのプラットフォーム用に設計されています。 VBA は、ユーザーのデスクトップと対話して、COM や OLE などの同様のテクノロジと接続できます。 ただし、VBA にはインターネットに呼び出す便利な方法はありません。 Officeスクリプトは、JavaScript のユニバーサル ランタイムを使用します。 これにより、スクリプトの実行に使用されるコンピューターに関係なく、一貫した動作とアクセシビリティが得されます。 また、他の Web サービスを呼び出す場合もあります。

### <a name="script-support-for-excel-on-windows"></a>スクリプトは、ExcelのWindows

[!INCLUDE [Run-from-button support](../includes/run-from-button-desktop-support.md)]

## <a name="security"></a>セキュリティ

VBA マクロは、VBA マクロと同じセキュリティExcel。 これにより、デスクトップにフル アクセスできます。 Officeスクリプトはブックにのみアクセスできます。ブックをホストするコンピューターにはアクセスできません。 さらに、スクリプトと JavaScript 認証トークンを共有できません。 つまり、スクリプトにはサインインしているユーザーのトークンも、外部サービスにサインインする API 機能もないので、既存のトークンを使用してユーザーの代わりに外部呼び出しを行う必要があります。

管理者には、VBA マクロの 3 つのオプションがあります。テナント上のすべてのマクロを許可する、テナントでマクロを許可する、または署名付き証明書を持つマクロのみを許可する。 この粒度が不足すると、1 つの悪いアクターを分離するのは難しいです。 現在、Officeスクリプトは、テナント全体、テナント全体、またはテナント内のユーザーのグループに対してオフにできます。 管理者は、他のユーザーとスクリプトを共有できるユーザーや、ユーザーが他のユーザーとスクリプトを使用できるユーザー Power Automate。

## <a name="coverage"></a>割合

現在、VBA では、特にデスクトップ クライアントでExcel機能に関するより完全な範囲を提供しています。 Officeスクリプトは、すべてのシナリオをサポートExcel on the web。 さらに、Web で新機能が登場すると、Officeスクリプトはアクション レコーダー API と JavaScript API の両方でサポートされます。

Officeスクリプトは、レベルのイベントExcelサポート[されていません](/office/vba/excel/concepts/events-worksheetfunctions-shapes/using-events-with-excel-objects)。 スクリプトは、ユーザーが手動で起動した場合、またはユーザー フローがスクリプトPower Automate呼び出す場合にのみ実行されます。

## <a name="power-automate"></a>Power Automate

Officeスクリプトは、スクリプトを使用Power Automate。 スケジュールされたフローまたはイベント駆動型フローを使用してブックを更新し、スケジュールされたワークフローを開かなくてもワークフローをExcel。 つまり、ブックが OneDrive (および Power Automate からアクセス可能) に保存されている限り、組織が Excel のデスクトップ、Mac、または Web クライアントを使用するかどうかに関係なく、フローでスクリプトを実行できます。

VBA には、新しいコネクタPower Automateがあります。 サポートされている VBA シナリオはすべて、マクロの実行に参加するユーザーを含む。

手順の[詳細については、手動Power Automateから](../tutorials/excel-power-automate-manual.md)スクリプトを呼び出Power Automate。 また、[自動タスク[アラーム] サンプル](scenarios/task-reminders.md)を参照して、Officeシナリオで Teams Power Automateに接続されているスクリプトを確認することもできます。

## <a name="see-also"></a>関連項目

- [Excel on the web の Office スクリプト](../overview/excel.md)
- [Power Automate を使用した Office スクリプトの実行](../develop/power-automate-integration.md)
- [Office スクリプトと Office アドインの違い](add-ins-differences.md)
- [Office スクリプトのトラブルシューティング](../testing/troubleshooting.md)
- [Excel VBA リファレンス](/office/vba/api/overview/excel)
