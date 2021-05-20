---
title: Officeスクリプト コード エディタ環境
description: Excel on the webのスクリプトOffice前提条件と環境情報。
ms.date: 05/10/2021
localization_priority: Normal
ms.openlocfilehash: aa54939826f8dda2a068df0f3fabf0fd3a2c842b
ms.sourcegitcommit: 4687693f02fc90a57ba30c461f35046e02e6f5fb
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 05/19/2021
ms.locfileid: "52545823"
---
# <a name="office-scripts-code-editor-environment"></a>Officeスクリプト コード エディタ環境

Officeスクリプトは TypeScript または JavaScript のいずれかで記述され、Officeスクリプト JavaScript API を使用してExcelワークブックと対話します。 コード エディターはVisual Studio Code基づいているため、以前にその環境を使用したことがある場合は、気持ちよく過ごせます。

## <a name="scripting-language-typescript-or-javascript"></a>スクリプト言語: タイプスクリプトまたは JavaScript

Officeスクリプトは[、JavaScript](https://developer.mozilla.org/docs/Web/JavaScript)のスーパーセットである[TypeScript](https://www.typescriptlang.org/docs/home.html)で記述されています。 アクション レコーダは TypeScript でコードを生成し、Officeスクリプトのドキュメントでは TypeScript を使用します。 TypeScript は JavaScript のスーパーセットなので、JavaScript で記述したスクリプト コードは正常に動作します。

Officeスクリプトは、主に自己完結型のコードです。 TypeScript の機能のごく一部のみが使用されます。 したがって、TypeScript の複雑さを学習することなくスクリプトを編集できます。 コード エディタは、コードのインストール、コンパイル、実行も処理するため、スクリプト自体以外の問題を心配する必要はありません。 プログラミングの知識がなくても、言語を学習し、スクリプトを作成することができます。 ただし、プログラミングを始めた場合は、Officeスクリプトを使用する前に、いくつかの基礎を学ぶことをお勧めします。

[!INCLUDE [Preview note](../includes/coding-basics-references.md)]

## <a name="office-scripts-javascript-api"></a>Officeスクリプト Java スクリプト API

Officeスクリプトでは[、Office](/office/dev/add-ins/overview/index)アドイン用の Office JavaScript API の特殊なバージョンを使用します。2 つの API には類似点がありますが、2 つのプラットフォーム間でコードを移植できると想定しないでください。 2 つのプラットフォームの違いについては[、「Office スクリプトとOfficeアドインの違い」を参照してください](../resources/add-ins-differences.md#apis)。 スクリプトで使用できるすべての API については[、「Officeスクリプト API リファレンス」のドキュメントを参照](/javascript/api/office-scripts/overview)できます。

## <a name="external-library-support"></a>外部ライブラリのサポート

Officeスクリプトは、外部のサードパーティ製 JavaScript ライブラリの使用法をサポートしていません。 現在、スクリプトからOfficeスクリプト API 以外のライブラリを呼び出すことはできません。 [Math](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Math)などの[組み込みの JavaScript オブジェクト](../develop/javascript-objects.md)に引き続きアクセスできます。

## <a name="intellisense"></a>IntelliSense

IntelliSenseは、スクリプトを編集するときに誤字や構文エラーを防ぐコード エディター機能です。 入力時に使用可能なオブジェクト名とフィールド名、およびすべての API のインライン ドキュメントが表示されます。

Excel コード エディターは、Visual Studio Codeと同じIntelliSense エンジンを使用します。 この機能の詳細については[、Visual Studio CodeのIntelliSense機能](https://code.visualstudio.com/docs/editor/intellisense#_intellisense-features)を参照してください。

## <a name="keyboard-shortcuts"></a>キーボード ショートカット

Visual Studio Codeのキーボード ショートカットのほとんどは、Office スクリプト コード エディターでも機能します。 次の PDF を使用して、使用可能なオプションについて学習し、コード エディターを最大限に活用してください。

- [macOS のキーボードショートカット](https://code.visualstudio.com/shortcuts/keyboard-shortcuts-macos.pdf).
- [Windowsのキーボード ショートカット](https://code.visualstudio.com/shortcuts/keyboard-shortcuts-windows.pdf)。

## <a name="see-also"></a>関連項目

- [Office スクリプト API リファレンス](/javascript/api/office-scripts/overview)
- [Office スクリプトのトラブルシューティング](../testing/troubleshooting.md)
- [Office スクリプトでの組み込みの JavaScript オブジェクトの使用](../develop/javascript-objects.md)
