---
title: スクリプトと VBA Officeの違い
description: スクリプトマクロと Excel VBA マクロOffice API の違い。
ms.date: 12/14/2020
localization_priority: Normal
ms.openlocfilehash: a56409a5de3eb07876faa88bfbfe78eeca59f70f
ms.sourcegitcommit: 45ffe3dbd2c834b78592ad35928cf8096f5e80bc
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 04/14/2021
ms.locfileid: "51755022"
---
# <a name="differences-between-office-scripts-and-vba-macros"></a>スクリプトと VBA Officeの違い

Officeと VBA マクロの共通点が多い。 どちらも、ユーザーが使いやすいアクション レコーダーを使用してソリューションを自動化し、それらの記録の編集を許可します。 どちらのフレームワークも、自分自身をプログラマと見なしていない人が Excel で小さなプログラムを作成するように設計されています。
基本的な違いは、VBA マクロがデスクトップ ソリューション用に開発され、スクリプトOfficeプラットフォーム間のサポートとセキュリティを指針として設計されている点です。 現在、Officeスクリプトは Web 上の Excel でのみサポートされています。

:::image type="content" source="../images/office-programmability-diagram.png" alt-text="さまざまな拡張ソリューションのフォーカス領域を示す 4 象限Office図。スクリプトOffice VBA マクロはどちらも、エンド ユーザーがソリューションを作成するのに役立つ設計ですが、Office スクリプトは Web とコラボレーション用に構築されています (一方、VBA はデスクトップ用です)。":::

この記事では、VBA マクロ (VBA 全般と同様) とスクリプトの主な違いOfficeします。 スクリプトOffice Excel でのみ使用できるので、ここで説明する唯一のホストです。

## <a name="platform-and-ecosystem"></a>プラットフォームとエコシステム

VBA はデスクトップ用に設計されOfficeスクリプトは Web 用に設計されています。 VBA は、ユーザーのデスクトップと対話して、COM や OLE などの同様のテクノロジと接続できます。 ただし、VBA にはインターネットに呼び出す便利な方法はありません。

Officeスクリプトは JavaScript のユニバーサル ランタイムを使用します。 これにより、スクリプトの実行に使用されるコンピューターに関係なく、一貫した動作とアクセシビリティが得されます。 また、他の Web サービスを呼び出す場合もあります。

## <a name="security"></a>セキュリティ

VBA マクロは Excel と同じセキュリティクリアランスを持っています。 これにより、デスクトップにフル アクセスできます。 Officeはブックにのみアクセスできます。ブックをホストするコンピューターにはアクセスできません。 さらに、スクリプトと JavaScript 認証トークンを共有できません。 つまり、スクリプトにはサインインしているユーザーのトークンも、外部サービスにサインインする API 機能もないので、既存のトークンを使用してユーザーの代わりに外部呼び出しを行う必要があります。

管理者には、VBA マクロの 3 つのオプションがあります。テナント上のすべてのマクロを許可する、テナントでマクロを許可する、または署名付き証明書を持つマクロのみを許可する。 この粒度が不足すると、1 つの悪いアクターを分離するのは難しいです。 現在、Officeスクリプトはテナントのオンまたはオフのどちらかです。 ただし、管理者が個々のスクリプトとスクリプト作成者を詳細に制御する作業を行っています。

## <a name="coverage"></a>割合

現在、VBA は Excel 機能、特にデスクトップ クライアントで使用できる機能のより完全な範囲を提供しています。 Officeスクリプトは、Web 上の Excel のシナリオのほぼすべてについて説明します。 さらに、Web で新機能が登場すると、Officeスクリプトはアクション レコーダー API と JavaScript API の両方でサポートされます。

Officeスクリプトは Excel レベルのイベントをサポート [しない](/office/vba/excel/concepts/events-worksheetfunctions-shapes/using-events-with-excel-objects)。 スクリプトは、ユーザーが手動で起動するか、Power Automate フローがスクリプトを呼び出す場合にのみ実行されます。

## <a name="power-automate"></a>Power Automate

Officeスクリプトは、Power Automate を使用して実行できます。 スケジュールされたフローまたはイベント駆動型フローを使用してブックを更新し、Excel を開かなくてもワークフローを自動化できます。 つまり、ブックが OneDrive に保存されている (Power Automate からアクセスできる) 限り、ユーザーと組織が Excel のデスクトップ、Mac、または Web クライアントを使用するかどうかに関係なく、フローでスクリプトを実行できます。

VBA には Power Automate コネクタが存在しない。 サポートされている VBA シナリオはすべて、マクロの実行に参加するユーザーを含む。

## <a name="see-also"></a>関連項目

- [Excel on the web の Office スクリプト](../overview/excel.md)
- [Office スクリプトと Office アドインの違い](add-ins-differences.md)
- [Office スクリプトのトラブルシューティング](../testing/troubleshooting.md)
- [Excel VBA リファレンス](/office/vba/api/overview/excel)
