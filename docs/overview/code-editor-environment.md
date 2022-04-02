---
title: Office スクリプト コード エディター環境
description: スクリプトの前提条件と環境OfficeをExcel on the web。
ms.date: 05/27/2021
ms.localizationpriority: medium
ms.openlocfilehash: 165365d82aa838f6651461f6edee2389c44e90b1
ms.sourcegitcommit: 7023b9e23499806901a5ecf8ebc460b76887cca6
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 03/31/2022
ms.locfileid: "64585934"
---
# <a name="office-scripts-code-editor-environment"></a>Office スクリプト コード エディター環境

Officeスクリプトは TypeScript または JavaScript のどちらかで記述され、Office スクリプト JavaScript API を使用してブックを操作Excelします。 コード エディターは、Visual Studio Codeに基づいており、以前に環境を使用した場合は、自宅のように感じるでしょう。

## <a name="scripting-language-typescript-or-javascript"></a>スクリプト言語: TypeScript または JavaScript

オフィス スクリプトは [TypeScript](https://www.typescriptlang.org/docs/home.html) で書かれており、[JavaScript](https://developer.mozilla.org/docs/Web/JavaScript) のスーパーセットです。 アクション レコーダーは TypeScript でコードを生成し、スクリプトのドキュメントOffice TypeScript を使用します。 TypeScript は JavaScript のスーパーセットですから、JavaScript で記述するスクリプト コードはうまく動作します。

Officeスクリプトは、主に自己格納型のコードです。 TypeScript の機能の一部だけが使用されます。 したがって、TypeScript の内容を学習せずにスクリプトを編集できます。 コード エディターは、コードのインストール、コンパイル、および実行も処理します。そのため、スクリプト自体以外は心配する必要はありません。 言語を学び、以前のプログラミング知識なしでスクリプトを作成することができます。 ただし、プログラミングが新しい場合は、次のスクリプトを実行する前に、いくつかの基本的なOffice勧めします。

[!INCLUDE [Recommended coding resources](../includes/coding-basics-references.md)]

## <a name="office-scripts-javascript-api"></a>Officeスクリプト JavaScript API

Officeスクリプトは、JavaScript API の専用バージョンをOfficeアドイン用Office[使用します](/office/dev/add-ins/overview/index)。2 つの API には類似点がありますが、2 つのプラットフォーム間でコードを移植できると想定する必要はありません。 2 つのプラットフォームの違いについては、「スクリプトとアドインのOfficeの違Office[」を参照](../resources/add-ins-differences.md#apis)してください。 スクリプトで使用可能なすべての API は、「スクリプト API リファレンス」Office[参照ドキュメントで確認できます](/javascript/api/office-scripts/overview)。

## <a name="external-library-support"></a>外部ライブラリのサポート

Office外部のサードパーティ JavaScript ライブラリの使用はサポートされていません。 現在、スクリプトからスクリプト API 以外のライブラリOffice呼び出す必要があります。 まだ、Math などの組み込みの [JavaScript](../develop/javascript-objects.md) オブジェクトにアクセス [できます](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Math)。

## <a name="intellisense"></a>IntelliSense

IntelliSenseは、コードの記述に役立つ一連のコード エディター機能です。 オートコンプリート、構文エラー強調表示、インライン API ドキュメントを提供します。

IntelliSense入力時に、入力時に候補テキストと同様の候補が表示Excel。 Tab キーまたは Enter キーを押すと、候補メンバーが挿入されます。 Ctrl IntelliSenseスペース キーを押して、現在のカーソル位置でトリガーします。 これらの提案は、メソッドを完了するときに特に便利です。 IntelliSense によって表示されるメソッドシグネチャには、必要な引数の一覧、各引数の型、指定した引数が必須か省略可能か、メソッドの戻り値の型が含まれる。

メソッド、クラス、または他のコード オブジェクトの上にカーソルを置くと、詳細が表示されます。 赤または黄色の線で表される構文エラーまたはコード候補にカーソルを合わせると、問題の解決方法に関する提案が表示されます。 多くの場合IntelliSenseコードを自動的に変更するための "クイック修正" オプションが提供されます。

:::image type="content" source="../images/implicit-any-editor-message.png" alt-text="コード エディターのホバー テキストに 'Quick Fix' ボタンが表示されたエラー メッセージ。":::

スクリプト Office コード エディターは、スクリプト コード エディターと同IntelliSenseエンジンをVisual Studio Code。 この機能の詳細については、「Visual Studio Code[機能」をIntelliSenseしてください](https://code.visualstudio.com/docs/editor/intellisense#_intellisense-features)。

## <a name="keyboard-shortcuts"></a>キーボード ショートカット

ユーザーのキーボード ショートカットの大部分Visual Studio Codeスクリプト コード エディター Office機能します。 次の PDF を使用して、使用可能なオプションについて説明し、コード エディターを利用できます。

- [macOS のキーボード ショートカット](https://code.visualstudio.com/shortcuts/keyboard-shortcuts-macos.pdf)。
- [ユーザーのキーボード ショートカットWindows](https://code.visualstudio.com/shortcuts/keyboard-shortcuts-windows.pdf)。

## <a name="see-also"></a>関連項目

- [Office スクリプト API リファレンス](/javascript/api/office-scripts/overview)
- [Office スクリプトのトラブルシューティング](../testing/troubleshooting.md)
- [Office スクリプトでの組み込みの JavaScript オブジェクトの使用](../develop/javascript-objects.md)
