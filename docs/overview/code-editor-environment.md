---
title: Office スクリプト コード エディター環境
description: Excel on the webの Office スクリプトの前提条件と環境情報。
ms.date: 11/08/2022
ms.localizationpriority: medium
ms.openlocfilehash: a5a7601285553b1da4001a1870b6120f21bf5f2c
ms.sourcegitcommit: 7cadf2b637bf62874e43b6e595286101816662aa
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 11/09/2022
ms.locfileid: "68891254"
---
# <a name="office-scripts-code-editor-environment"></a>Office スクリプト コード エディター環境

Office スクリプトは TypeScript または JavaScript のいずれかで記述され、Office スクリプト JavaScript API を使用して Excel ブックと対話します。 コード エディターは Visual Studio Code に基づいているため、その環境を以前に使用したことがある場合は、自宅にいるように感じます。

> [!TIP]
> Visual Studio Code に慣れている場合は、それを使用してスクリプトを記述できるようになりました。 この機能を試すには、「 [Visual Studio Code for Office Scripts (プレビュー)」](../develop/vscode-for-scripts.md) を参照してください。

## <a name="scripting-language-typescript-or-javascript"></a>スクリプト言語: TypeScript または JavaScript

オフィス スクリプトは [TypeScript](https://www.typescriptlang.org/docs/home.html) で書かれており、[JavaScript](https://developer.mozilla.org/docs/Web/JavaScript) のスーパーセットです。 アクション レコーダーは TypeScript でコードを生成し、Office スクリプトのドキュメントでは TypeScript を使用します。 TypeScript は JavaScript のスーパーセットであるため、JavaScript で記述するすべてのスクリプト コードは正常に動作します。

Office スクリプトは、主に自己完結型のコードです。 TypeScript の機能のごく一部のみが使用されます。 そのため、TypeScript の複雑さを学習することなくスクリプトを編集できます。 コード エディターでは、コードのインストール、コンパイル、実行も処理されるため、スクリプト自体以外の何も心配する必要はありません。 以前のプログラミング知識がなくても、言語を学習し、スクリプトを作成できます。 ただし、プログラミングを初めて使用する場合は、Office スクリプトを続行する前に、いくつかの基礎を学ぶことをお勧めします。

[!INCLUDE [Recommended coding resources](../includes/coding-basics-references.md)]

## <a name="office-scripts-javascript-api"></a>Office スクリプト JavaScript API

Office スクリプトでは、Office アドイン用の Office JavaScript API の特殊なバージョンが使用 [されます](/office/dev/add-ins/overview/index)。2 つの API には類似点がありますが、2 つのプラットフォーム間でコードを移植できるとは想定しないでください。 2 つのプラットフォームの違いについては、「 [Office スクリプトと Office アドインの違い」](../resources/add-ins-differences.md#apis) の記事を参照してください。 スクリプトで使用できるすべての API は、 [Office Scripts API リファレンス ドキュメント](/javascript/api/office-scripts/overview)で確認できます。

## <a name="external-library-support"></a>外部ライブラリのサポート

Office スクリプトでは、外部のサード パーティの JavaScript ライブラリの使用はサポートされていません。 現時点では、スクリプトから Office Scripts API 以外のライブラリを呼び出すことはできません。 [Math](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Math) などの[組み込みの JavaScript オブジェクト](../develop/javascript-objects.md)には引き続きアクセスできます。

## <a name="intellisense"></a>Intellisense

IntelliSense は、コードを記述するのに役立つ一連のコード エディター機能です。 オートコンプリート、構文エラーの強調表示、インライン API ドキュメントを提供します。

IntelliSense では、Excel で推奨されるテキストと同様に、入力時に候補が表示されます。 Tab キーまたは Enter キーを押すと、推奨されるメンバーが挿入されます。 Ctrl + Space キーを押して、現在のカーソル位置で IntelliSense をトリガーします。 これらの提案は、メソッドを完了するときに特に役立ちます。 IntelliSense によって表示されるメソッド シグネチャには、必要な引数の一覧、各引数の型、指定された引数が必須か省略可能か、メソッドの戻り値の型が含まれます。

メソッド、クラス、またはその他のコード オブジェクトの上にカーソルを合わせると、詳細が表示されます。 赤または黄色の波線で表される構文エラーまたはコードの提案にカーソルを合わせると、問題を解決する方法に関する提案が表示されます。 多くの場合、IntelliSense には、コードを自動的に変更するための "クイック修正" オプションが用意されています。

:::image type="content" source="../images/implicit-any-editor-message.png" alt-text="コード エディターのホバー テキストに &quot;クイック修正&quot; ボタンが表示されたエラー メッセージ。":::

Office スクリプト コード エディターでは、Visual Studio Code と同じ IntelliSense エンジンが使用されます。 この機能の詳細については、 [Visual Studio Code の IntelliSense 機能](https://code.visualstudio.com/docs/editor/intellisense#_intellisense-features)に関するページを参照してください。

## <a name="keyboard-shortcuts"></a>キーボード ショートカット

Visual Studio Code のキーボード ショートカットのほとんどは、Office スクリプト コード エディターでも機能します。 次の PDF を使用して、使用可能なオプションについて学習し、コード エディターを最大限に活用します。

- [macOS のキーボード ショートカット](https://code.visualstudio.com/shortcuts/keyboard-shortcuts-macos.pdf)。
- [Windows のキーボード ショートカット](https://code.visualstudio.com/shortcuts/keyboard-shortcuts-windows.pdf)。

## <a name="see-also"></a>関連項目

- [Office スクリプト API リファレンス](/javascript/api/office-scripts/overview)
- [Office スクリプトのトラブルシューティング](../testing/troubleshooting.md)
- [Office スクリプトでの組み込みの JavaScript オブジェクトの使用](../develop/javascript-objects.md)
- [Visual Studio Code for Office Scripts (プレビュー)](../develop/vscode-for-scripts.md)
