---
title: Power Query または Office スクリプトを使用する場合
description: スクリプト プラットフォームとスクリプト プラットフォームの両方に最適Power QueryシナリオOfficeシナリオ。
ms.date: 11/23/2021
ms.localizationpriority: medium
ms.openlocfilehash: e91077d635d66dde692c129bdd4b2f32657d5283
ms.sourcegitcommit: 7023b9e23499806901a5ecf8ebc460b76887cca6
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 03/31/2022
ms.locfileid: "64585906"
---
# <a name="when-to-use-power-query-or-office-scripts"></a>Power Query または Office スクリプトを使用する場合

[Power Query](https://powerquery.microsoft.com)スクリプトとOfficeスクリプトは、両方とも強力なオートメーション ソリューションです。Excel。 どちらのソリューションでも、Excelのデータをクリーンアップおよび変換できます。 1 Power Queryまたは Office スクリプトを更新して新しいデータを再実行して、一貫性のある結果を生成し、時間を節約し、結果の情報をより速く処理できます。

この記事では、一方のプラットフォームを他のプラットフォームより好む可能性がある場合の概要を説明します。 一般に、Power Query は、大規模な外部データ ソースからデータをプルして変換する場合や、Office スクリプトを使用すると、Excel 中心の迅速なソリューションとPower Automate統合[に最適](../develop/power-automate-integration.md)です。

## <a name="large-data-sources-and-data-retrieval-power-query"></a>大規模なデータ ソースとデータ取得: Power Query

サポートされているPower Queryデータ ソースを処理する場合は、この方法をお勧めします。

Power Queryには[、何百ものソースへの](https://powerquery.microsoft.com/connectors/)組み込みデータ接続があります。 Power Queryは、データ取得、変換、および組み合わせタスク用に特別に設計されています。 これらのソースの 1 つからのデータが必要な場合、Power Query を使用すると、そのデータを必要な図形Excelコードを使用してデータを取得できます。

これらのPower Query接続は、大規模なデータセット用に設計されています。 これらの転送制限は[、ユーザーまたは](../testing/platform-limits.md)ユーザーのPower AutomateとExcel for the web。

Officeスクリプトは、小規模なデータ ソースまたはデータ ソースがコネクタでカバーされない軽量Power Query提供します。 これには、[アダプティブ カード`fetch`などのアドホック](../develop/external-calls.md) データ ソースからの情報の取得や REST API [の使用Teams含まれます](../resources/scenarios/task-reminders.md)。

## <a name="formatting-visualizations-and-programmatic-control-office-scripts"></a>書式設定、視覚化、およびプログラムによる制御: Office スクリプト

データのインポートOffice変換以外のニーズに対応する場合は、スクリプトを使用することをお勧めします。

UI を使用して手動で実行できるExcel、スクリプトを使用Officeできます。 ブックに一貫性のある書式を適用する場合に最適です。 スクリプトは、グラフ、ピボットテーブル、図形、画像、その他のワークシートの視覚化を作成します。 スクリプトを使用すると、これらの視覚エフェクトの位置、サイズ、色、その他の属性を正確に制御できます。

TypeScript コードを組み込むと、カスタマイズの度合いが高い。 ステートメントのようなプログラムによる制御ロジック `if...else` により、スクリプトは堅牢になります。 これにより、複雑な数式に依存せずに条件付きでデータを読み取Excelブックをスキャンして、ブックを変更する前に予期しない変更を行えます。

書式は、複数のテンプレートPower Query使用Excel[適用できます](https://templates.office.com/power-query-tutorial-tm11414620)。 ただし、テンプレートは個々または組織レベルで更新されます。一方、Officeスクリプトは、より詳細なアクセス制御を提供します。

## <a name="power-automate-integrations"></a>Power Automate統合

Officeスクリプトには、統合のためのより多くのPower Automateがあります。 スクリプトはソリューションに合わせて調整されます。 スクリプトの [入力と出力を定義](../develop/power-automate-integration.md#data-transfer-in-flows-for-scripts)して、フロー内の他のコネクタまたはデータと一緒に動作します。 次のスクリーン ショットは、Power Automateアダプティブ カードからデータをTeamsスクリプトに渡すOffice示しています。

:::image type="content" source="../images/scenario-task-reminders-last-flow-step.png" alt-text="フロー デザイナーの [Excel] (Business) コネクタを示すスクリーンショット。コネクタは、[スクリプトの実行] アクションを使用して、アダプティブ カードから入力Teamsスクリプトに提供します。":::

Power Queryコネクタ[でSQL Server Power Automate](https://powerquery.microsoft.com/flow/)されます。 [[データの変換] Power Query](/connectors/sql/#transform-data-using-power-query)アクションを使用すると、クエリを作成してクエリを作成Power Automate。 このツールは、SQL Serverで使用する強力なツールですが、次Power Queryのフロー スクリーンショットに示すように、入力ソースに制限されます。

:::image type="content" source="../images/power-query-flow-option.png" alt-text="フロー デザイナーのSQL Serverを示すスクリーンショット。コネクタは、[データの変換] アクションを使用してデータPower Queryしています。":::

## <a name="platform-dependencies"></a>プラットフォームの依存関係

Officeスクリプトは現在、ユーザーが使用Excel on the web。 Power Queryデスクトップ上のユーザーのみExcel使用できます。 両方とも、Power Automateを使用して使用できます。これにより、フローは、Excelに格納されているブックとOneDrive。

## <a name="see-also"></a>関連項目

- [Power Query ポータル](https://powerquery.microsoft.com/)
- [Power QueryのExcel](https://powerquery.microsoft.com/excel/)
- [Power Automate を使用した Office スクリプトの実行](../develop/power-automate-integration.md)
