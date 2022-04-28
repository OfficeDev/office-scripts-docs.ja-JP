---
title: Office スクリプトのサンプル
description: スクリプトのサンプルとシナリオOffice使用できます。
ms.date: 04/26/2022
ms.localizationpriority: medium
ms.openlocfilehash: 7c9bbe9b6f7eb8abad2995dac72ccf636d585d69
ms.sourcegitcommit: e6428a5214fa38aef036a952a0e3c09dbf6e4d3e
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 04/28/2022
ms.locfileid: "65109159"
---
# <a name="office-scripts-samples-and-scenarios"></a>Office スクリプトのサンプルとシナリオ

このセクションには、エンド ユーザーが毎日のタスクの自動化を実現するのに役立つ[、スクリプト](../../overview/excel.md) ベースの自動化ソリューションOfficeが含まれています。 ビジネス ユーザーが直面する現実的なシナリオが含まれており、詳細なソリューションとステップバイステップの手順ビデオ リンクが提供されます。

[[基本]](#basics) と [[基本以外](#beyond-the-basics)] の各プロジェクトについて、ソース コード、ステップ バイ ステップ [**の YouTube ビデオ**](https://www.youtube.com/playlist?list=PLr3zVPZrMOUMl88fs8uc2GGAePRnNe6m0)などを確認してください。

シナリオには、実際 [の](#scenarios)ユース ケースを示す大規模なシナリオ サンプルがいくつか含まれています。

また、 [コミュニティからの貢献](#community-contributions-and-fun-samples)も歓迎します。

## <a name="basics"></a>基本事項

| Project | 詳細 |
|---------|---------|
| [スクリプトの基本](../excel-samples.md) | これらのサンプルでは、Office スクリプトの基本的な構成要素を示します。 |
| [Excelにコメントを追加する](add-excel-comments.md) | このサンプルでは、同僚@mentioning含むコメントをセルに追加します。 |
| [ブックに画像を追加する](add-image-to-workbook.md) | このサンプルでは、ブックに画像を追加し、シート間で画像をコピーします。|
| [複数のExcel テーブルを 1 つのテーブルにコピーする](copy-tables-combine.md) | このサンプルでは、複数のExcel テーブルのデータを、すべての行を含む 1 つのテーブルに結合します。 |
| [ブックの目次を作成する](table-of-contents.md) | このサンプルでは、各ワークシートへのリンクを含む目次を作成します。 |

## <a name="beyond-the-basics"></a>応用

サンプル シナリオを自動化する次のエンド ツー エンド プロジェクトと、完全なスクリプト、使用されるサンプル Excel ファイル、[およびビデオ (YouTube でホスト) を](https://www.youtube.com/playlist?list=PLr3zVPZrMOUMl88fs8uc2GGAePRnNe6m0)確認してください。

| Project | 詳細 |
|---------|---------|
| [ワークシートを 1 つのブックに結合する](combine-worksheets-into-single-workbook.md) | このサンプルでは、Office スクリプトとPower Automateを使用して、他のブックから 1 つのブックにデータをプルします。 |
| [CSV ファイルをExcel ブックに変換する](convert-csv.md) | このサンプルでは、Office スクリプトとPower Automateを使用して、.csv ファイルから.xlsx ファイルを作成します。 |
| [相互参照ブック](excel-cross-reference.md) | このサンプルでは、Office スクリプトとPower Automateを使用して、さまざまなブックの情報を相互参照および検証します。 |
| [特定のシートまたはすべてのシート内の空白行をカウントする](count-blank-rows.md) | このサンプルでは、シートにデータが存在すると予測される空白行があるかどうかを検出し、Power Automate フローでの使用状況の空白行数を報告します。 |
| [電子メール グラフと表の画像](email-images-chart-table.md) | このサンプルでは、Office スクリプトとPower Automateアクションを使用してグラフを作成し、そのグラフを電子メールで画像として送信します。 |
| [外部フェッチ呼び出し](external-fetch-calls.md) | このサンプルでは、スクリプトのGitHubから情報を取得するために使用`fetch`します。 |
| [テーブルExcelフィルター処理し、表示範囲を取得する](filter-table-get-visible-range.md) | このサンプルでは、Excel テーブルをフィルター処理し、表示範囲を JSON オブジェクトとして返します。 この JSON は、大規模なソリューションの一部としてPower Automate フローに提供できます。 |
| [Excelで計算モードを管理する](excel-calculation.md) | このサンプルでは、Office スクリプトを使用して計算モードを使用し、Excel on the webでメソッドを計算する方法を示します。 |
| [テーブル間で行を移動する](move-rows-across-tables.md) | このサンプルでは、フィルターを保存し、フィルターを処理して再適用することで、テーブル間で行を移動する方法を示します。 |
| [JSON としてデータExcel出力する](get-table-data.md) | このソリューションでは、Power Automateで使用Excelテーブル データを JSON として出力する方法を示します。 |
| [Excel ワークシート内の各セルからハイパーリンクを削除する](remove-hyperlinks-from-cells.md) | このサンプルでは、現在のワークシートからすべてのハイパーリンクをクリアします。 |
| [フォルダー内のすべての Excel ファイルでスクリプトを実行する](automate-tasks-on-all-excel-files-in-folder.md) | このプロジェクトは、OneDrive for Business上のフォルダーに置かれたすべてのファイルに対して一連の自動化タスクを実行します (SharePoint フォルダーにも使用できます)。 Excel ファイルに対して計算を実行し、書式設定を追加し、同僚@mentionsコメントを挿入します。 |
| [大規模データセットを書き込む](write-large-dataset.md) | このサンプルでは、より小さい部分範囲として大きな範囲を送信する方法を示します。 |

## <a name="scenarios"></a>シナリオ

Office スクリプトは、毎日のルーチンの一部を自動化できます。 これらの日常的なタスクは、多くの場合、固有のエコシステムに存在し、特定の方法で設定されたExcelブックがあります。 これらの大規模なシナリオ サンプルは、このような実際のユース ケースを示しています。 Office スクリプトとブックの両方が含まれているので、シナリオを最後から最後まで確認できます。

| シナリオ | 詳細 |
|---------|---------|
| [Web ダウンロードの分析](../scenarios/analyze-web-downloads.md) | このシナリオでは、Web トラフィック レコードを解析してユーザーの配信元の国を特定するスクリプトが用意されています。 スクリプトでサブ関数を使用し、条件付き書式を適用し、テーブルを操作するテキスト解析のスキルを紹介します。 |
| [NOAA の水位データを取得してグラフ化する](../scenarios/noaa-data-fetch.md) | このシナリオでは、Office スクリプトを使用して外部ソース ([NOAA Tides および Currents データベース](https://tidesandcurrents.noaa.gov/)) からデータをプルし、結果の情報をグラフ化します。 データの取得とグラフの使用に使用 `fetch` するスキルが強調されています。 |
| [グレード計算機](../scenarios/grade-calculator.md) | このシナリオでは、クラスの成績についてインストラクターのレコードを検証するスクリプトが用意されています。 エラーチェック、セルの書式設定、正規表現のスキルを紹介します。 |
| [Teams で面接をスケジュールする](../scenarios/schedule-interviews-in-teams.md) | このシナリオでは、Excel スプレッドシートを使用して面接会議の時間を管理し、Teamsで会議をスケジュールするフローを作成する方法を示します。 |
| [タスクのリマインダー](../scenarios/task-reminders.md) | このシナリオでは、Power Automate フローのOffice スクリプトを使用して、同僚にリマインダーを送信してプロジェクトの状態を更新します。 これは、Power Automate統合とスクリプトとの間のデータ転送のスキルを強調しています。 |

## <a name="community-contributions-and-fun-samples"></a>Community投稿と楽しいサンプル

Office スクリプト コミュニティからの[貢献](https://github.com/OfficeDev/office-scripts-docs/blob/master/Contributing.md)を歓迎します。 レビュー用のプル要求を自由に作成してください。

| Project | 詳細 |
|---------|---------|
| [Game of Life](https://techcommunity.microsoft.com/t5/excel-blog/ready-player-zero/ba-p/2246208) | Excel Tech Communityの Yutao Raspberr による "Ready Player Zero" ブログには、John Conway [*の The Game of Life*](https://en.wikipedia.org/wiki/Conway%27s_Game_of_Life) をモデル化するスクリプトが含まれています。 |
| [パンチ クロック ボタン](../scenarios/punch-clock.md) | このスクリプトは [、Brian Gonzalez](https://github.com/b-gonzalez) によって提供されました。 シナリオには、現在の時刻を記録するスクリプトとスクリプト ボタンが用意されています。 |
| [シーズンのあいさつアニメーション](community-seasons-greetings.md) | このスクリプトは、休日の季節の精神で [Leslie Black](https://www.linkedin.com/in/lesblackconsultant/) によって提供されました。 Office スクリプトを使用して、Excel on the webでクリスマス ツリーを歌うのを示す楽しいスクリプトです。 |

## <a name="try-it-out"></a>試してみる

これらのサンプルはオープンソース。 自分で試してみてください。 サブスクリプションをMicrosoft 365するライセンスを持つ職場または学校の Microsoft 職場または学校アカウント (E3 以上) が必要です。 アカウントに https://office.com サインインして作業を開始します。

## <a name="leave-a-comment"></a>コメントを残す

特定のサンプルのドキュメント ページの下部にある **フィードバック** セクションを使用して、コメントを残したり、提案をしたり、問題をログに記録したりできます。
