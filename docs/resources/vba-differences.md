---
title: スクリプトと VBA Officeの違い
description: スクリプトと VBA マクロの動作Office API Excel違い。
ms.date: 05/21/2021
localization_priority: Normal
ms.openlocfilehash: 0d94607902fa62e07ce378b94ec3b9c328937e16535b1882b6cad5bd76212b33
ms.sourcegitcommit: 75f7ed8c2d23a104acc293f8ce29ea580b4fcdc5
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 08/11/2021
ms.locfileid: "57847269"
---
# <a name="differences-between-office-scripts-and-vba-macros"></a>スクリプトと VBA Officeの違い

Officeスクリプトと VBA マクロには多くの共通点があります。 どちらも、ユーザーが使いやすいアクション レコーダーを使用してソリューションを自動化し、それらの記録の編集を許可します。 どちらのフレームワークも、プログラマが自分を考慮しない可能性があるユーザーに、小規模なプログラムを作成する権限を与Excel。
基本的な違いは、VBA マクロがデスクトップ ソリューション用に開発され、スクリプトOfficeクラウドベースのソリューション用に設計されている点です。 現在、Officeスクリプトは、このスクリプトでのみExcel on the web。

:::image type="content" source="../images/office-programmability-diagram.png" alt-text="さまざまな拡張ソリューションのフォーカス領域を示す 4 象限Office図。スクリプトOffice VBA マクロはどちらも、エンド ユーザーがソリューションを作成するのに役立つ設計ですが、Office スクリプトは Web とコラボレーション用に構築されています (VBA はデスクトップ用です)。":::

この記事では、VBA マクロ (および VBA 全般) とスクリプトの主な違いOffice説明します。 このOfficeスクリプトは、Excelでしか使用できないので、ここで説明する唯一のホストです。

## <a name="platform-and-ecosystem"></a>プラットフォームとエコシステム

VBA はデスクトップ用に設計されOfficeスクリプトは Web 用に設計されています。 VBA は、ユーザーのデスクトップと対話して、COM や OLE などの同様のテクノロジと接続できます。 ただし、VBA にはインターネットに呼び出す便利な方法はありません。

Officeスクリプトは JavaScript の汎用ランタイムを使用します。 これにより、スクリプトの実行に使用されるコンピューターに関係なく、一貫した動作とアクセシビリティが得されます。 また、他の Web サービスを呼び出す場合もあります。

## <a name="security"></a>セキュリティ

VBA マクロのセキュリティクリアランスは、VBA マクロと同Excel。 これにより、デスクトップにフル アクセスできます。 Officeスクリプトはブックにのみアクセスできます。ブックをホストしているコンピューターにはアクセスできません。 さらに、スクリプトと JavaScript 認証トークンを共有できません。 つまり、スクリプトにはサインインしているユーザーのトークンも、外部サービスにサインインする API 機能もないので、既存のトークンを使用してユーザーの代わりに外部呼び出しを行う必要があります。

管理者には、VBA マクロの 3 つのオプションがあります。テナント上のすべてのマクロを許可する、テナントでマクロを許可する、または署名付き証明書を持つマクロのみを許可する。 この粒度が不足すると、1 つの悪いアクターを分離するのは難しいです。 現在、Officeスクリプトは、テナント全体、テナント全体、またはテナント内のユーザーのグループに対してオフにできます。 また、管理者は、他のユーザーとスクリプトを共有できるユーザー、およびスクリプトを他のユーザーで使用できるユーザー Power Automate。

## <a name="coverage"></a>割合

現在、VBA では、デスクトップ クライアントで使用できるExcel機能のより完全な範囲を提供しています。 Officeスクリプトは、すべてのシナリオをサポートExcel on the web。 さらに、Web で新機能が登場すると、Officeスクリプトはアクション レコーダー API と JavaScript API の両方でサポートされます。

Officeスクリプトは、レベルのイベントExcelサポート[されていません](/office/vba/excel/concepts/events-worksheetfunctions-shapes/using-events-with-excel-objects)。 スクリプトは、ユーザーが手動で起動した場合、またはスクリプトを呼び出Power Automate場合にのみ実行されます。

## <a name="power-automate"></a>Power Automate

Officeスクリプトは、次の方法でPower Automate。 スケジュールされたフローまたはイベント駆動型フローを使用してブックを更新し、ワークフローを開かなくてもワークフローをExcel。 つまり、ブックが OneDrive (および Power Automate からアクセス可能) に格納されている限り、組織が Excel のデスクトップ、Mac、または Web クライアントを使用するかどうかに関係なく、フローでスクリプトを実行できます。

VBA には、新しいコネクタPower Automateがあります。 サポートされている VBA シナリオはすべて、マクロの実行に参加するユーザーを含む。

手順の[詳細については、手動Power Automateから](../tutorials/excel-power-automate-manual.md)スクリプトを呼び出Power Automate。 また、[自動タスク[アラーム]](scenarios/task-reminders.md)サンプルを参照して、Officeシナリオで Teams Power Automateに接続されているスクリプトを確認することもできます。

## <a name="see-also"></a>関連項目

- [Excel on the web の Office スクリプト](../overview/excel.md)
- [Power Automate を使用した Office スクリプトの実行](../develop/power-automate-integration.md)
- [Office スクリプトと Office アドインの違い](add-ins-differences.md)
- [Office スクリプトのトラブルシューティング](../testing/troubleshooting.md)
- [Excel VBA リファレンス](/office/vba/api/overview/excel)
