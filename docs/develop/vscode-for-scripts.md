---
title: Visual Studio Code for Office Scripts (プレビュー)
description: Web 用 VS Code に接続するように Office スクリプト コード エディターをセットアップする方法。
ms.date: 11/04/2022
ms.localizationpriority: medium
ms.openlocfilehash: fd9dd417610c8ad64fbd3fc50048ce56afdb4e28
ms.sourcegitcommit: 7cadf2b637bf62874e43b6e595286101816662aa
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 11/09/2022
ms.locfileid: "68892045"
---
# <a name="visual-studio-code-for-office-scripts-preview"></a>Visual Studio Code for Office Scripts (プレビュー)

[Web 用 Visual Studio Code](https://vscode.dev/) を使用すると、ユーザーはどこからでも何かを編集できます。 Office スクリプト エクスペリエンスをこの一般的なコード エディターに接続して、ブックの外部でスクリプトを開始します。

:::image type="content" source="../images/vscode-script-editor.png" alt-text="コード エディターが開いているExcel on the web ウィンドウが、開いているスクリプトを含む Web ウィンドウの VS Code の横に開きます。":::

Visual Studio Code には、組み込みのコード エディターよりもいくつかの利点があります。

- 全画面編集! スクリプトは、ブックと画面領域を共有する必要はありません。
- 複数のスクリプトを一度に編集! スクリプトをすばやく切り替えて、他のオートメーションからコードを共有します。
- 拡張 機能！ スペル チェック、書式設定、その他の作業を完了するのに役立つその他の機能については、お気に入りの VS Code 拡張機能を使用してください。

> [!NOTE]
> この機能はプレビュー段階です。 フィードバックに基づいて変更される場合があります。 問題が発生した場合は、Excel の **[フィードバック** ] ボタンから報告してください。 現在のバージョンの機能に関する既知の問題を次に示します。
>
> - Visual Studio Code は、Excel on the web経由でのみ Office スクリプトに接続できます。
> - この Office スクリプト接続は、英語の Excel クライアントでのみ使用できます。

## <a name="connect-visual-studio-code-to-office-scripts"></a>Visual Studio Code を Office スクリプトに接続する

Visual Studio Code と Excel on the webを接続するには、次の 1 回限りの手順に従います。

1. Office スクリプト **コード エディター** を開きます。
2. [ **その他のオプション (...)]** メニューの [ **エディター設定**] を選択します。
3. **[(プレビュー)] Visual Studio Code 接続を選択します**。

:::image type="content" source="../images/vscode-enable-option.png" alt-text="Visual Studio Code 接続というラベルが付いたチェック ボックスが表示されているエディター設定作業ウィンドウ。":::

これで、Visual Studio Code からスクリプトを編集して実行できます。 任意のスクリプトから、[ **その他のオプション (...)]** メニューに移動し、[ **VS Code で開く**] を選択します。

:::image type="content" source="../images/vscode-open-option.png" alt-text="開いているスクリプトの横にあるリストから選択されている [VS Code で開く] オプション。":::

## <a name="see-also"></a>関連項目

- [Office スクリプト コード エディター環境](../overview/code-editor-environment.md)
- [Visual Studio Code for the Web (ドキュメント)](https://code.visualstudio.com/docs/editor/vscode-web)
