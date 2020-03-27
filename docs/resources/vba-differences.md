---
title: Office スクリプトと VBA マクロの相違点
description: Office スクリプトと Excel VBA マクロの動作と API の違い。
ms.date: 03/23/2020
localization_priority: Normal
ms.openlocfilehash: 3a0f2c9a2ed7181a10e41d1f45b3af695877a680
ms.sourcegitcommit: d556aaefac80e55f53ac56b7f6ecbc657ebd426f
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 03/26/2020
ms.locfileid: "42978710"
---
# <a name="differences-between-office-scripts-and-vba-macros"></a>Office スクリプトと VBA マクロの相違点

Office スクリプトと VBA マクロには、多くの共通点があります。 両方のユーザーが、使いやすいアクションレコーダーを使用してソリューションを自動化し、それらのレコーディングを編集できるようにします。 両方のフレームワークは、プログラマが開発者にとって Excel で小さなプログラムを作成することを考慮していない可能性があるユーザーを支援するためのものです。
基本的な違いは、デスクトップソリューション用に VBA マクロが開発されており、Office スクリプトが、ガイド原則としてクロスプラットフォームのサポートとセキュリティを使用して設計されているということです。 現時点では、Office スクリプトは web 上の Excel でのみサポートされています。

![さまざまな Office 機能拡張ソリューションに対するフォーカスの領域を示す4つの領域の図。 Office スクリプトと VBA マクロはどちらも、エンドユーザーがソリューションを作成できるように設計されていますが、Office スクリプトは web およびコラボレーション用に構築されています (ただし、VBA はデスクトップ用)。)](../images/office-programmability-diagram.png)

この記事では、VBA マクロ (および VBA) と Office スクリプトの主な違いについて説明します。 Office スクリプトは Excel でのみ使用可能であるため、ここで説明する唯一のホストです。

## <a name="platform-and-ecosystem"></a>プラットフォームとエコシステム

VBA はデスクトップ向けに設計されており、Office スクリプトは web 用に設計されています。 VBA は、ユーザーのデスクトップと対話できます。 これにより、COM や OLE などの同様のテクノロジとの統合が可能になります。 ただし、VBA には、インターネットを呼び出すための便利な方法はありません。

Office スクリプトでは、ユニバーサルランタイムまたは JavaScript を使用します。 これにより、スクリプトの実行に使用されているコンピューターに関係なく、一貫性のある動作とアクセスが可能になります。 また、他の web サービスへの呼び出しを行うこともできます。

## <a name="security"></a>セキュリティ

VBA マクロには、Excel と同じセキュリティクリアランスがあります。 これにより、デスクトップへのフルアクセスが可能になります。 Office スクリプトは、ブックをホストするマシンではなく、ブックへのアクセスのみが可能です。 さらに、スクリプトでは、JavaScript 認証トークンを共有できないため、スクリプトは外部サービスで認証されません。

管理者には、VBA マクロに関する3つのオプションがあります。テナントのすべてのマクロを許可するか、テナントにマクロを許可しないか、署名された証明書によるマクロのみを許可します。 このように細分化されていないと、1つの不良アクターを分離するのが困難になります。 現時点では、Office スクリプトはテナントに対してオンまたはオフになっています。 しかし、管理者が個々のスクリプトやスクリプト作成者をより詳細に制御できるようにしています。

## <a name="coverage"></a>割合

現時点では、VBA は、デスクトップクライアントで使用可能な Excel 機能の詳細な範囲を提供しています。 Office スクリプトでは、web 上の Excel のほぼすべてのシナリオについて説明します。 また、web 上の新機能の debut により、Office スクリプトはアクションレコーダーと JavaScript Api の両方に対してそれらをサポートします。

## <a name="power-automate"></a>Power Automate

Office スクリプトは、電源自動化を使用して実行できます。 ブックは、スケジュールまたはイベントドリブンフローによって更新できます。これにより、Excel を開かなくてもワークフローを自動化できます。 これは、ブックが OneDrive に保存されていて、そのユーザーが Excel のデスクトップ、Mac、または web クライアントを使用しているかどうかに関係なく、フローがスクリプトを実行できることを意味します。

VBA には、Power 自動化との統合はありません。 サポートされているすべての VBA シナリオでは、マクロの実行にユーザーが参加していました。

## <a name="see-also"></a>関連項目

- [Excel on the web の Office スクリプト](../overview/excel.md)
- [Office スクリプトと Office アドインの相違点](add-ins-differences.md)
- [Office スクリプトのトラブルシューティング](../testing/troubleshooting.md)
- [Excel VBA リファレンス](/office/vba/api/overview/excel)
