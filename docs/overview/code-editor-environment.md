---
title: Officeスクリプト コード エディター環境
description: スクリプトの前提条件と環境情報は、OfficeスクリプトExcel on the web。
ms.date: 05/10/2021
localization_priority: Normal
ms.openlocfilehash: aa54939826f8dda2a068df0f3fabf0fd3a2c842b
ms.sourcegitcommit: 4687693f02fc90a57ba30c461f35046e02e6f5fb
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 05/19/2021
ms.locfileid: "52545823"
---
# <a name="office-scripts-code-editor-environment"></a>Officeスクリプト コード エディター環境

Officeスクリプトは TypeScript または JavaScript で記述され、Office スクリプト JavaScript API を使用して、Excelブックを操作します。 コード エディターは、Visual Studio Codeに基づいており、以前に環境を使用した場合は、自宅ですぐに使用できます。

## <a name="scripting-language-typescript-or-javascript"></a>スクリプト言語: TypeScript または JavaScript

オフィス スクリプトは [TypeScript](https://www.typescriptlang.org/docs/home.html) で書かれており、[JavaScript](https://developer.mozilla.org/docs/Web/JavaScript) のスーパーセットです。 アクション レコーダーは TypeScript でコードを生成し、スクリプトOffice TypeScript を使用します。 TypeScript は JavaScript のスーパーセットですから、JavaScript で記述するスクリプト コードはうまく動作します。

Officeスクリプトは、主に自己格納型のコードです。 TypeScript の機能の一部だけが使用されます。 したがって、TypeScript の内容を学習せずにスクリプトを編集できます。 コード エディターは、コードのインストール、コンパイル、および実行も処理します。そのため、スクリプト自体以外は心配する必要はありません。 言語を学び、以前のプログラミング知識なしでスクリプトを作成することができます。 ただし、プログラミングを始めてお勧めする場合は、次のスクリプトを実行する前に、いくつかのOffice勧めします。

[!INCLUDE [Preview note](../includes/coding-basics-references.md)]

## <a name="office-scripts-javascript-api"></a>Officeスクリプト JavaScript API

Officeスクリプトでは、JavaScript API の専用Officeバージョンを使用Office[アドインを使用します](/office/dev/add-ins/overview/index)。2 つの API には類似点がありますが、2 つのプラットフォーム間でコードを移植できると想定する必要はありません。 2 つのプラットフォームの違いについては、「スクリプトとアドインのOffice[の違Office」を参照](../resources/add-ins-differences.md#apis)してください。 スクリプトで使用可能なすべての API は、「スクリプト API リファレンス」Office[で確認できます](/javascript/api/office-scripts/overview)。

## <a name="external-library-support"></a>外部ライブラリのサポート

Officeスクリプトは、外部のサードパーティの JavaScript ライブラリの使用をサポートしていない。 現在、スクリプトからスクリプト API 以外のライブラリOffice呼び出す必要があります。 まだ、Math などの組み込みの [JavaScript](../develop/javascript-objects.md)オブジェクトにアクセス [できます](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Math)。

## <a name="intellisense"></a>IntelliSense

IntelliSenseは、スクリプトの編集時に入力ミスや構文エラーを防ぐのに役立つコード エディター機能です。 入力時に可能なオブジェクト名とフィールド名、およびすべての API のインライン ドキュメントが表示されます。

コード Excelエディターは、コード エディターと同IntelliSenseエンジンをVisual Studio Code。 この機能の詳細については、「Visual Studio Code[機能」をIntelliSenseしてください](https://code.visualstudio.com/docs/editor/intellisense#_intellisense-features)。

## <a name="keyboard-shortcuts"></a>キーボード ショートカット

ユーザーのキーボード ショートカットの大部分Visual Studio Codeスクリプト コード エディター Office機能します。 次の PDF を使用して、使用可能なオプションについて説明し、コード エディターを利用できます。

- [macOS のキーボード ショートカット](https://code.visualstudio.com/shortcuts/keyboard-shortcuts-macos.pdf)。
- [ユーザーのキーボード ショートカットWindows。](https://code.visualstudio.com/shortcuts/keyboard-shortcuts-windows.pdf)

## <a name="see-also"></a>関連項目

- [Office スクリプト API リファレンス](/javascript/api/office-scripts/overview)
- [Office スクリプトのトラブルシューティング](../testing/troubleshooting.md)
- [Office スクリプトでの組み込みの JavaScript オブジェクトの使用](../develop/javascript-objects.md)
