---
title: スクリプトと VBA Officeの違い
description: スクリプトと VBA マクロの動作Office API Excel違い。
ms.date: 12/14/2020
localization_priority: Normal
ms.openlocfilehash: ca571e2adad81a87b99696a652a3c49209b870ab
ms.sourcegitcommit: f7a7aebfb687f2a35dbed07ed62ff352a114525a
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 05/06/2021
ms.locfileid: "52232845"
---
# <a name="differences-between-office-scripts-and-vba-macros"></a>スクリプトと VBA Officeの違い

Officeスクリプトと VBA マクロには多くの共通点があります。 どちらも、ユーザーが使いやすいアクション レコーダーを使用してソリューションを自動化し、それらの記録の編集を許可します。 どちらのフレームワークも、プログラマが自分を考慮しない可能性があるユーザーに、小規模なプログラムを作成する権限を与Excel。
基本的な違いは、VBA マクロがデスクトップ ソリューション用に開発され、スクリプトOfficeプラットフォーム間のサポートとセキュリティを基本原則として設計されている点です。 現在、Officeスクリプトは、このスクリプトでのみExcel on the web。

:::image type="content" source="../images/office-programmability-diagram.png" alt-text="さまざまな拡張ソリューションのフォーカス領域を示す 4 象限Office図。Office スクリプトと VBA マクロの両方は、エンド ユーザーがソリューションを作成するのに役立つ設計ですが、Office スクリプトは Web とコラボレーション用に構築されています (VBA はデスクトップ用です)。":::

この記事では、VBA マクロ (および VBA 全般) とスクリプトの主な違いOffice説明します。 このOfficeスクリプトは、Excelでしか使用できないので、ここで説明する唯一のホストです。

## <a name="platform-and-ecosystem"></a>プラットフォームとエコシステム

VBA はデスクトップ用に設計されOfficeスクリプトは Web 用に設計されています。 VBA は、ユーザーのデスクトップと対話して、COM や OLE などの同様のテクノロジと接続できます。 ただし、VBA にはインターネットに呼び出す便利な方法はありません。

Officeスクリプトは JavaScript の汎用ランタイムを使用します。 これにより、スクリプトの実行に使用されるコンピューターに関係なく、一貫した動作とアクセシビリティが得されます。 また、他の Web サービスを呼び出す場合もあります。

## <a name="security"></a>セキュリティ

VBA マクロのセキュリティクリアランスは、VBA マクロと同Excel。 これにより、デスクトップにフル アクセスできます。 Officeスクリプトはブックにのみアクセスできます。ブックをホストしているコンピューターにはアクセスできません。 さらに、スクリプトと JavaScript 認証トークンを共有できません。 つまり、スクリプトにはサインインしているユーザーのトークンも、外部サービスにサインインする API 機能もないので、既存のトークンを使用してユーザーの代わりに外部呼び出しを行う必要があります。

管理者には、VBA マクロの 3 つのオプションがあります。テナント上のすべてのマクロを許可する、テナントでマクロを許可する、または署名付き証明書を持つマクロのみを許可する。 この粒度が不足すると、1 つの悪いアクターを分離するのは難しいです。 現在、Officeスクリプトはテナントのオンまたはオフのどちらかです。 ただし、管理者が個々のスクリプトとスクリプト作成者を詳細に制御する作業を行っています。

## <a name="coverage"></a>割合

現在、VBA では、デスクトップ クライアントで使用できるExcel機能のより完全な範囲を提供しています。 Officeスクリプトは、すべてのシナリオをサポートExcel on the web。 さらに、Web で新機能が登場すると、Officeスクリプトはアクション レコーダー API と JavaScript API の両方でサポートされます。

Officeスクリプトは、レベルのイベントExcelサポート[されていません](/office/vba/excel/concepts/events-worksheetfunctions-shapes/using-events-with-excel-objects)。 スクリプトは、ユーザーが手動で起動した場合、またはスクリプトを呼び出Power Automate場合にのみ実行されます。

## <a name="power-automate"></a>Power Automate

Officeスクリプトは、次の方法でPower Automate。 スケジュールされたフローまたはイベント駆動型フローを使用してブックを更新し、ワークフローを開かなくてもワークフローをExcel。 つまり、ブックが OneDrive (および Power Automate からアクセス可能) に格納されている限り、組織が Excel のデスクトップ、Mac、または Web クライアントを使用するかどうかに関係なく、フローでスクリプトを実行できます。

VBA には、新しいコネクタPower Automateがあります。 サポートされている VBA シナリオはすべて、マクロの実行に参加するユーザーを含む。

## <a name="see-also"></a>こちらもご覧ください

- [Excel on the web の Office スクリプト](../overview/excel.md)
- [Office スクリプトと Office アドインの違い](add-ins-differences.md)
- [Office スクリプトのトラブルシューティング](../testing/troubleshooting.md)
- [Excel VBA リファレンス](/office/vba/api/overview/excel)
