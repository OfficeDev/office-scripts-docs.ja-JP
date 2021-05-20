---
title: Officeスクリプトサンプル
description: 使用可能なOffice スクリプトのサンプルとシナリオ。
ms.date: 05/17/2021
localization_priority: Normal
ms.openlocfilehash: 0ea9a8a8986681fca0e45784e2923c1d3b34576d
ms.sourcegitcommit: 4687693f02fc90a57ba30c461f35046e02e6f5fb
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 05/19/2021
ms.locfileid: "52545710"
---
# <a name="office-scripts-samples-and-scenarios"></a>Officeスクリプトのサンプルとシナリオ

このセクションでは[、Officeのスクリプト](../../overview/excel.md)ベースの自動化ソリューションを含み、エンドユーザーが日常業務の自動化を実現できるようにします。 ビジネス ユーザーが直面する現実的なシナリオが含まれ、詳細なソリューションと段階的な説明ビデオ リンクが提供されます。

[「基本](#basics)」と [「その先」](#beyond-the-basics)の各プロジェクトについては、ソースコード、ステップバイステップの [**YouTube動画**](https://www.youtube.com/playlist?list=PLr3zVPZrMOUMl88fs8uc2GGAePRnNe6m0)などをチェックしてください。

シナリオ では [、](#scenarios)実際のユース ケースを示すいくつかの大規模なシナリオ サンプルが含まれています。

また、 [コミュニティからの貢献を](#community-contributions-and-fun-samples)歓迎します。

[!INCLUDE [Preview note](../../includes/preview-note.md)]

## <a name="basics"></a>基本事項

| Project | 詳細 |
|---------|---------|
| [スクリプトの基本](../excel-samples.md) | これらのサンプルでは、Office スクリプトの基本的な構成要素を示します。 |
| [Excelにコメントを追加する](add-excel-comments.md) | このサンプルでは、同僚@mentioningを含むセルにコメントを追加します。 |
| [ブックにイメージを追加する](add-image-to-workbook.md) | このサンプルでは、ブックに画像を追加し、シート間で画像をコピーします。|
| [複数のExcelテーブルを 1 つのテーブルにコピーする](copy-tables-combine.md) | このサンプルでは、複数のExcel テーブルのデータを、すべての行を含む単一のテーブルに結合します。 |

## <a name="beyond-the-basics"></a>応用

完全なスクリプト、使用されるサンプルExcelファイル、[および動画(YouTube でホストされている)](https://www.youtube.com/playlist?list=PLr3zVPZrMOUMl88fs8uc2GGAePRnNe6m0)と共にサンプル シナリオを自動化する、次のエンド ツー エンド プロジェクトをチェックしてください。

| Project | 詳細 |
|---------|---------|
| [特定のシートまたはすべてのシートの空白行をカウントする](count-blank-rows.md) | このサンプルでは、データが存在することを予測するシートに空白行があるかどうかを検出し、Power Automate フローでの使用に関する空白行数を報告します。 |
| [電子メールチャートとテーブル画像](email-images-chart-table.md) | このサンプルでは、OfficeスクリプトとPower Automateアクションを使用して、グラフを作成し、そのグラフを画像として電子メールで送信します。 |
| [外部フェッチ呼び出し](external-fetch-calls.md) | このサンプルでは `fetch` 、スクリプトのGitHubから情報を取得するために使用します。 |
| [テーブルExcelフィルターして表示範囲を取得する](filter-table-get-visible-range.md) | このサンプルでは、Excelテーブルをフィルター処理し、表示範囲を JSON オブジェクトとして返します。 この JSON は、より大きなソリューションの一部としてPower Automateフローに提供できます。 |
| [Excelで計算モードを管理する](excel-calculation.md) | このサンプルでは、Office スクリプトを使用して、計算モードを使用してExcel on the webメソッドを計算する方法を示します。 |
| [テーブル間で行を移動する](move-rows-across-tables.md) | このサンプルでは、フィルターを保存し、フィルターを処理して再適用することにより、テーブル間で行を移動する方法を示します。 |
| [データExcel JSON として出力する](get-table-data.md) | このソリューションでは、Power Automateで使用する json としてExcel表データを出力する方法を示します。 |
| [Excel ワークシートの各セルからハイパーリンクを削除する](remove-hyperlinks-from-cells.md) | このサンプルでは、現在のワークシートからすべてのハイパーリンクを消去します。 |
| [フォルダー内のすべての Excel ファイルでスクリプトを実行する](automate-tasks-on-all-excel-files-in-folder.md) | このプロジェクトは、OneDrive for Business上のフォルダにあるすべてのファイルに対して一連のオートメーション タスクを実行します (SharePoint フォルダにも使用できます)。 Excelファイルに対して計算を実行し、書式を追加し、同僚@mentionsコメントを挿入します。 |
| [大規模なデータセットを記述する](write-large-dataset.md) | このサンプルでは、大きな範囲を小さい部分範囲として送信する方法を示します。 |

## <a name="scenarios"></a>シナリオ

Officeスクリプトは、毎日のルーチンの一部を自動化できます。 これらの日常のタスクは、多くの場合、特定の方法で設定されたExcelブックを使用して、独自のエコシステムに存在します。 これらの大規模なシナリオ サンプルは、このような実際のユース ケースを示しています。 スクリプトとワークブックのOfficeの両方が含まれているので、シナリオをエンドツーエンドで確認できます。

| シナリオ | 詳細 |
|---------|---------|
| [Web ダウンロードの分析](../scenarios/analyze-web-downloads.md) | このシナリオでは、Web トラフィック レコードを解析してユーザーの出身国を決定するスクリプトを使用します。 テキストの解析、スクリプトでのサブ機能の使用、条件付き書式の適用、およびテーブルの操作のスキルを紹介します。 |
| [NOAA の水位データを取得してグラフ化する](../scenarios/noaa-data-fetch.md) | このシナリオでは、Office スクリプトを使用して外部ソース[(NOAA Tides と Currents データベース](https://tidesandcurrents.noaa.gov/)) からデータを取得し、結果の情報をグラフ化します。 `fetch`データの取得やグラフの使用に使用するスキルを強調しています。 |
| [グレード計算機](../scenarios/grade-calculator.md) | このシナリオでは、クラスの成績に対する教員の記録を検証するスクリプトを使用します。 エラー チェック、セルの書式設定、および正規表現のスキルを紹介します。 |
| [タスクのリマインダー](../scenarios/task-reminders.md) | このシナリオでは、Power Automate フローでOffice スクリプトを使用して、プロジェクトのステータスを更新するリマインダーを同僚に送信します。 スクリプトとの間でのPower Automate統合とデータ転送のスキルを強調します。 |

## <a name="community-contributions-and-fun-samples"></a>Community貢献と楽しいサンプル

Officeスクリプトコミュニティからの[貢献](https://github.com/OfficeDev/office-scripts-docs/blob/master/Contributing.md)を歓迎します! レビューのためのプルリクエストを自由に作成してください。

| Project | 詳細 |
|---------|---------|
| [人生のゲーム](https://techcommunity.microsoft.com/t5/excel-blog/ready-player-zero/ba-p/2246208) | Excel技術Communityのユタオ・ホアンの「レディ・プレイヤー・ゼロ」ブログには、ジョン・コンウェイ [*の『人生のゲーム*](https://en.wikipedia.org/wiki/Conway%27s_Game_of_Life)』をモデル化するスクリプトが含まれています。 |
| [季節の挨拶アニメーション](community-seasons-greetings.md) | この脚本は、ホリデーシーズンの精神で [レスリー・ブラック](https://www.linkedin.com/in/lesblackconsultant/) によって貢献されました! これは、Officeスクリプトを使用してExcel on the webで歌うクリスマスツリーを示す楽しいスクリプトです。 |

## <a name="try-it-out"></a>試してみる

これらのサンプルはオープンソースです。 自分で試してみてください。 職場または学校の職場または学校のアカウントが必要で、サブスクリプション (E3 以上) をMicrosoft 365するライセンスを持っています。 https://office.comアカウントにサインインして開始してください。

## <a name="leave-a-comment"></a>コメントを残す

特定のサンプルのドキュメント ページの下部にある **[フィードバック** ] セクションを使用して、コメントを残したり、提案をしたり、問題を記録したりできます。
