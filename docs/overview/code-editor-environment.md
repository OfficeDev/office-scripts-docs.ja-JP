---
title: Officeスクリプト コード エディター環境
description: スクリプトの前提条件と環境情報は、OfficeスクリプトExcel on the web。
ms.date: 05/27/2021
localization_priority: Normal
ms.openlocfilehash: 4a8adc03e372bc769fb44b1c4e3e98c7a4531756
ms.sourcegitcommit: 4693c8f79428ec74695328275703af0ba1bfea8f
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 06/23/2021
ms.locfileid: "53074467"
---
# <a name="office-scripts-code-editor-environment"></a>Officeスクリプト コード エディター環境

Officeスクリプトは TypeScript または JavaScript で記述され、Office スクリプト JavaScript API を使用して、Excelブックを操作します。 コード エディターは、Visual Studio Codeに基づいており、以前に環境を使用した場合は、自宅ですぐに使用できます。

## <a name="scripting-language-typescript-or-javascript"></a>スクリプト言語: TypeScript または JavaScript

オフィス スクリプトは [TypeScript](https://www.typescriptlang.org/docs/home.html) で書かれており、[JavaScript](https://developer.mozilla.org/docs/Web/JavaScript) のスーパーセットです。 アクション レコーダーは TypeScript でコードを生成し、スクリプトOffice TypeScript を使用します。 TypeScript は JavaScript のスーパーセットですから、JavaScript で記述するスクリプト コードはうまく動作します。

Officeスクリプトは、主に自己格納型のコードです。 TypeScript の機能の一部だけが使用されます。 したがって、TypeScript の内容を学習せずにスクリプトを編集できます。 コード エディターは、コードのインストール、コンパイル、および実行も処理します。そのため、スクリプト自体以外は心配する必要はありません。 言語を学び、以前のプログラミング知識なしでスクリプトを作成することができます。 ただし、プログラミングを始めてお勧めする場合は、次のスクリプトを実行する前に、いくつかのOffice勧めします。

[!INCLUDE [Recommended coding resources](../includes/coding-basics-references.md)]

## <a name="office-scripts-javascript-api"></a>Officeスクリプト JavaScript API

Officeスクリプトでは、JavaScript API の専用Officeバージョンを使用Office[アドインを使用します](/office/dev/add-ins/overview/index)。2 つの API には類似点がありますが、2 つのプラットフォーム間でコードを移植できると想定する必要はありません。 2 つのプラットフォームの違いについては、「スクリプトとアドインのOffice[の違Office」を参照](../resources/add-ins-differences.md#apis)してください。 スクリプトで使用可能なすべての API は、「スクリプト API リファレンス」Office[で確認できます](/javascript/api/office-scripts/overview)。

## <a name="external-library-support"></a>外部ライブラリのサポート

Officeスクリプトは、外部のサードパーティの JavaScript ライブラリの使用をサポートしていない。 現在、スクリプトからスクリプト API 以外のライブラリOffice呼び出す必要があります。 まだ、Math などの組み込みの [JavaScript](../develop/javascript-objects.md)オブジェクトにアクセス [できます](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Math)。

## <a name="intellisense"></a>IntelliSense

IntelliSenseは、コードの記述に役立つ一連のコード エディター機能です。 オートコンプリート、構文エラー強調表示、インライン API ドキュメントを提供します。

IntelliSense入力時に、入力時の候補テキストと同様に、候補が表示Excel。 Tab キーまたは Enter キーを押すと、候補メンバーが挿入されます。 Ctrl IntelliSenseスペース キーを押して、現在のカーソル位置でトリガーします。 これらの提案は、メソッドを完了するときに特に便利です。 IntelliSense によって表示されるメソッドシグネチャには、必要な引数の一覧、各引数の型、指定した引数が必須か省略可能か、メソッドの戻り値の型が含まれる。

メソッド、クラス、または他のコード オブジェクトの上にカーソルを置くと、詳細が表示されます。 赤または黄色の線で表される構文エラーまたはコード候補にカーソルを合わせると、問題の解決方法に関する提案が表示されます。 多くの場合IntelliSenseコードを自動的に変更するための "クイック修正" オプションが提供されます。

:::image type="content" source="../images/implicit-any-editor-message.png" alt-text="コード エディターのホバー テキストに 'Quick Fix' ボタンが表示されたエラー メッセージ。":::

スクリプト Office コード エディターは、スクリプト コード エディターと同IntelliSenseエンジンをVisual Studio Code。 この機能の詳細については、「Visual Studio Code[機能」をIntelliSenseしてください](https://code.visualstudio.com/docs/editor/intellisense#_intellisense-features)。

## <a name="keyboard-shortcuts"></a>キーボード ショートカット

ユーザーのキーボード ショートカットの大部分Visual Studio Codeスクリプト コード エディター Office機能します。 次の PDF を使用して、使用可能なオプションについて説明し、コード エディターを利用できます。

- [macOS のキーボード ショートカット](https://code.visualstudio.com/shortcuts/keyboard-shortcuts-macos.pdf)。
- [ユーザーのキーボード ショートカットWindows。](https://code.visualstudio.com/shortcuts/keyboard-shortcuts-windows.pdf)

## <a name="see-also"></a>関連項目

- [Office スクリプト API リファレンス](/javascript/api/office-scripts/overview)
- [Office スクリプトのトラブルシューティング](../testing/troubleshooting.md)
- [Office スクリプトでの組み込みの JavaScript オブジェクトの使用](../develop/javascript-objects.md)
