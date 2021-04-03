---
title: Officeスクリプトのサンプル
description: 使用可能なOfficeスクリプトのサンプルとシナリオです。
ms.date: 04/01/2021
localization_priority: Normal
ms.openlocfilehash: de0e99cbac7fcdeb1a3d3c43dd72ce53ed5847dd
ms.sourcegitcommit: 5d24e77df70aa2c1c982275d53213c2a9323ff86
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 04/02/2021
ms.locfileid: "51571426"
---
# <a name="office-scripts-samples-and-scenarios"></a>Officeスクリプトのサンプルとシナリオ

このセクションでは、エンド [Officeタスク](../../overview/excel.md) の自動化を実現するためのスクリプト ベースのオートメーション ソリューションについて説明します。 これは、ビジネス ユーザーが直面する現実的なシナリオが含まれています。詳細なソリューションと、ステップバイステップの説明ビデオ リンクを提供します。

「基本」および「基本 [](#basics)」の各 [](#beyond-the-basics)プロジェクトについて、ソース コード、ステップ バイ ステップ [**の YouTube**](https://www.youtube.com/playlist?list=PLr3zVPZrMOUMl88fs8uc2GGAePRnNe6m0)ビデオなどをご覧ください。

シナリオ [では](#scenarios)、実際の使用例を示すいくつかの大きなシナリオ サンプルが含まれています。

コミュニティからの [投稿も歓迎します](#community-contributions)。

[!INCLUDE [Preview note](../../includes/preview-note.md)]

## <a name="basics"></a>基本事項

| Project | 詳細 |
|---------|---------|
| [スクリプトの基本](../excel-samples.md) | これらのサンプルでは、スクリプトの基本的な構成要素Office示します。 |
| [スクリプトで Range オブジェクトを使用する方法のOfficeする](range-basics.md) | この記事では、Range オブジェクトとその API の使用の基本について説明します。 これは、他のすべてのプロジェクトで使用される基礎トピックです。 |

## <a name="beyond-the-basics"></a>基本を超えて

完全なスクリプト、使用されているサンプル Excel ファイル、およびビデオと共にサンプル シナリオを自動化する次のエンドツーエンド プロジェクトを確認 [してください](https://www.youtube.com/playlist?list=PLr3zVPZrMOUMl88fs8uc2GGAePRnNe6m0)。

| Project | 詳細 |
|---------|---------|
| [Excel でコメントを追加する](add-excel-comments.md) | このサンプルでは、同僚を含むセルにコメントを追加@mentioning示します。 |
| [特定のシートまたはすべてのシートの空白行を数える](count-blank-rows.md) | このサンプルでは、データが存在すると予測されるシートに空白行が含まれるかを検出し、Power Automate フローで使用するために空白行数を報告します。 |
| [Excel ファイルの相互参照と書式設定](excel-cross-reference.md) | このソリューションは、スクリプトと Power Automate を使用して 2 つの Excel ファイルを相互参照および書式設定Office示します。 |
| [電子メール グラフと表の画像](email-images-chart-table.md) | このサンプルでは、Officeおよび Power Automate アクションを使用してグラフを作成し、そのグラフを画像として電子メールで送信します。 |
| [Excel テーブルをフィルター処理し、表示範囲を取得する](filter-table-get-visible-range.md) | このサンプルでは、Excel テーブルをフィルター処理し、表示範囲を JSON オブジェクトとして返します。 この JSON は、大規模なソリューションの一部として Power Automate フローに提供できます。 |
| [ブックで一意の識別子を生成する](document-number-generator.md) | このシナリオは、ユーザーが特定の形式の一意の文書番号を生成し、範囲またはテーブルにエントリを追加するのに役立ちます。 |
| [Excel で計算モードを管理する](excel-calculation.md) | このサンプルでは、Excel on the web で計算モードを使用し、スクリプトを使用してメソッドを計算するOfficeします。 |
| [複数の Excel テーブルを 1 つのテーブルに結合する](copy-tables-combine.md) | このサンプルでは、複数の Excel テーブルのデータを、すべての行を含む 1 つのテーブルに結合します。 |
| [テーブル間で行を移動する](move-rows-across-tables.md) | このサンプルでは、フィルターを保存し、フィルターを処理して再適用することで、テーブル間で行を移動する方法を示します。 |
| [Excel データを JSON として出力する](get-table-data.md) | このソリューションでは、Power Automate で使用する EXCEL テーブル データを JSON として出力する方法を示します。 |
| [Excel ワークシートの各セルからハイパーリンクを削除する](remove-hyperlinks-from-cells.md) | このサンプルでは、現在のワークシートからすべてのハイパーリンクをクリアします。 |
| [フォルダー内のすべての Excel ファイルでスクリプトを実行する](automate-tasks-on-all-excel-files-in-folder.md) | このプロジェクトは、OneDrive for Business のフォルダー内のすべてのファイル (SharePoint フォルダーにも使用できます) に対して一連の自動化タスクを実行します。 Excel ファイルの計算を実行し、書式設定を追加し、同僚にコメント@mentions挿入します。 |
| [Excel データから Teams 会議を送信する](send-teams-invite-from-excel-data.md) | このソリューションでは、Excel ファイルから行Officeを選択し、Teams 会議の招待を送信して Excel を更新するために、スクリプトと Power Automate アクションを使用する方法を示します。 |

## <a name="scenarios"></a>シナリオ

Officeスクリプトは、日常の一部を自動化できます。 これらの日次タスクは、多くの場合、固有のエコシステムに存在し、Excel ブックは特定の方法で設定されます。 これらの大規模なシナリオ サンプルでは、このような実際の使用例を示します。 これらのスクリプトには、Officeスクリプトとブックの両方が含まれるので、シナリオを最後から最後まで確認できます。

| シナリオ | 詳細 |
|---------|---------|
| [Web ダウンロードの分析](../scenarios/analyze-web-downloads.md) | このシナリオでは、Web トラフィック レコードを解析してユーザーの原産国を決定するスクリプトを備えます。 スクリプトのサブ関数の使用、条件付き書式の適用、テーブルの操作など、テキスト解析のスキルを紹介します。 |
| [NOAA の水位データを取得してグラフ化する](../scenarios/noaa-data-fetch.md) | このシナリオでは、Officeスクリプトを使用して、外部ソース [(NOAA Tides](https://tidesandcurrents.noaa.gov/)および Currents データベース) からデータを取得し、結果の情報をグラフ化します。 データを取得し、グラフを使用 `fetch` するスキルを強調します。 |
| [グレード計算機](../scenarios/grade-calculator.md) | このシナリオでは、教員の成績を検証するスクリプトを備えます。 エラー チェック、セルの書式設定、および正規表現のスキルを紹介します。 |
| [タスクのリマインダー](../scenarios/task-reminders.md) | このシナリオでは、power Automate フロー Officeスクリプトを使用して、同僚にリマインダーを送信してプロジェクトの状態を更新します。 Power Automate の統合とスクリプトとの間のデータ転送のスキルを強調しています。 |

## <a name="community-contributions"></a>コミュニティへの投稿

スクリプト コミュニティ [からの投稿](https://github.com/OfficeDev/office-scripts-docs/blob/master/Contributing.md) Office歓迎します。 レビューのプル要求を自由に作成してください。

| Project | 詳細 |
|---------|---------|
| [シーズンの案内応答アニメーション](community-seasons-greetings.md) | このスクリプトは、ホリデー シーズンの精神 [でレス](https://www.linkedin.com/in/lesblackconsultant/) リー ブラックによって投稿されました。 このスクリプトは、Web 上の Excel で歌うクリスマス ツリーを、Web スクリプトを使用して表示Officeです。 |

## <a name="try-it-out"></a>試してみる

これらのサンプルはオープンソースです。 自分で試してみてください。 Microsoft 365 サブスクリプション (E3 以上) のライセンスを持つ、仕事または学校の Microsoft の仕事または学校のアカウントが必要です。 アカウントにサインイン https://office.com して開始してください。

## <a name="leave-a-comment"></a>コメントを残す

特定のサンプルのドキュメント ページの下部にある [フィードバック] セクションを使用して、コメントを残したり、提案をしたり、問題を記録したりしてください。
