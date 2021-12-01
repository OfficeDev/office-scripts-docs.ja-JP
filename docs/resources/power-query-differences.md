---
title: Power Query または Office スクリプトを使用する場合
description: Power Query プラットフォームとスクリプト プラットフォームの両方に最適Officeシナリオ。
ms.date: 11/23/2021
ms.localizationpriority: medium
ms.openlocfilehash: 1812b508b2cde4d304ecf228adfdd8f68de9808a
ms.sourcegitcommit: 383880e0dc0d09b8f76884675531e462a292d747
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 12/01/2021
ms.locfileid: "61245610"
---
# <a name="when-to-use-power-query-or-office-scripts"></a>Power Query または Office スクリプトを使用する場合

[Power Query スクリプト](https://powerquery.microsoft.com)と Office スクリプトは、両方とも強力なオートメーション ソリューションです。Excel。 どちらのソリューションでも、Excelのデータをクリーンアップおよび変換できます。 単一の Power Query または Office スクリプトを更新し、新しいデータを再実行して一貫性のある結果を生成し、時間を節約し、結果の情報をより速く処理できます。

この記事では、一方のプラットフォームを他のプラットフォームより好む可能性がある場合の概要を説明します。 一般に、Power Query は、大規模な外部データ ソースからデータをプルして変換する場合に最適です。Office スクリプトは、Excel 中心のソリューションと Power Automate 統合を迅速に行う[のに最適](../develop/power-automate-integration.md)です。

## <a name="large-data-sources-and-data-retrieval-power-query"></a>大規模なデータ ソースとデータ取得: Power Query

サポートされているプラットフォームのデータ ソースを処理する場合は、Power Query をお勧めします。

Power Query には、 [何百ものソースへの](https://powerquery.microsoft.com/connectors/) 組み込みデータ接続があります。 Power Query は、データ取得、変換、および組み合わせタスク用に特別に設計されています。 これらのソースの 1 つからのデータが必要な場合は、Power Query を使用すると、そのデータを必要な図形Excelコードを使用してデータを取得できます。

これらの Power Query 接続は、大規模なデータセット用に設計されています。 これらの転送制限は[、データの](../testing/platform-limits.md)転送または設定と同Power Automate制限Excel for the web。

Officeスクリプトは、Power Query コネクタでカバーされていない小規模なデータ ソースまたはデータ ソースに対して軽量なソリューションを提供します。 これには[、API `fetch` の使用や REST、](../develop/external-calls.md)アダプティブ カードなどのアドホック データ ソースからの情報の取得[Teams含まれます](../resources/scenarios/task-reminders.md)。

## <a name="formatting-visualizations-and-programmatic-control-office-scripts"></a>書式設定、視覚化、およびプログラムによる制御: Office スクリプト

データのインポートOffice変換以外のニーズに対応する場合は、スクリプトを使用することをお勧めします。

UI を使用して手動で実行できるExcel、スクリプトを使用Officeできます。 ブックに一貫性のある書式を適用する場合に最適です。 スクリプトは、グラフ、ピボットテーブル、図形、画像、その他のワークシートの視覚化を作成します。 スクリプトを使用すると、これらの視覚エフェクトの位置、サイズ、色、その他の属性を正確に制御できます。

TypeScript コードを組み込むと、カスタマイズの度合いが高い。 ステートメントのようなプログラムによる制御 `if...else` ロジックにより、スクリプトは堅牢になります。 これにより、複雑な数式に依存せずに条件付きでデータを読み取Excelブックをスキャンして、ブックを変更する前に予期しない変更を行えます。

書式は、Power Query を使用して、複数のテンプレートExcel[適用できます](https://templates.office.com/power-query-tutorial-tm11414620)。 ただし、テンプレートは個人または組織レベルで更新されます。一方、Officeスクリプトは、より詳細なアクセス制御を提供します。

## <a name="power-automate-integrations"></a>Power Automate統合

Officeスクリプトには、統合のためのより多くのPower Automateがあります。 スクリプトはソリューションに合わせて調整されます。 スクリプトの [入力と出力を定義](../develop/power-automate-integration.md#data-transfer-in-flows-for-scripts)して、フロー内の他のコネクタまたはデータと動作します。 次のスクリーンショットは、アダプティブ Power AutomateからデータをTeamsスクリプトに渡すOffice示しています。

:::image type="content" source="../images/scenario-task-reminders-last-flow-step.png" alt-text="フロー デザイナーの [Excel] (Business) コネクタを示すスクリーンショット。コネクタは、[スクリプトの実行] アクションを使用して、アダプティブ カードから入力Teamsスクリプトに提供します。":::

Power Query は、新[しいコネクタ](https://powerquery.microsoft.com/flow/)SQL Server Power Automateされます。 [Power [Query を使用してデータを変換](/connectors/sql/#transform-data-using-power-query)する] アクションを使用すると、クエリをデータベースにPower Automate。 このツールは、SQL Serverで使用する強力なツールですが、次のフロー スクリーンショットに示すように、Power Query を入力ソースに制限します。

:::image type="content" source="../images/power-query-flow-option.png" alt-text="フロー デザイナーのSQL Serverを示すスクリーンショット。コネクタは、[Power Query を使用してデータを変換する] アクションを使用しています。":::

## <a name="platform-dependencies"></a>プラットフォームの依存関係

Officeスクリプトは現在、ユーザーが使用Excel on the web。 Power Query は現在、デスクトップ上のExcelでのみ使用できます。 両方とも、Power Automateを使用して使用できます。これにより、フローは、Excelに格納されているブックとOneDrive。

## <a name="see-also"></a>関連項目

- [Power Query Portal](https://powerquery.microsoft.com/)
- [Power Query with Excel](https://powerquery.microsoft.com/excel/)
- [Power Automate を使用した Office スクリプトの実行](../develop/power-automate-integration.md)
