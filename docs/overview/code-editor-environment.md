---
title: Office スクリプトのコードエディター環境
description: Web 上の Excel の Office スクリプトの前提条件と環境情報。
ms.date: 07/10/2020
localization_priority: Normal
ms.openlocfilehash: 643ea2d5bd69adf4311546465ccd65c08dacf4b4
ms.sourcegitcommit: ebd1079c7e2695ac0e7e4c616f2439975e196875
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 07/17/2020
ms.locfileid: "45160496"
---
# <a name="office-scripts-code-editor-environment"></a>Office スクリプトのコードエディター環境

Office スクリプトは、 [TypeScript または javascript](#scripting-language-typescript-or-javascript)で記述され、 [Office スクリプト javascript api](#office-scripts-javascript-api)を使用して Excel ブックを操作します。

## <a name="scripting-language-typescript-or-javascript"></a>スクリプト言語: TypeScript または JavaScript

Office スクリプトは、 [TypeScript](https://www.typescriptlang.org/docs/home.html)または[JavaScript](https://developer.mozilla.org/docs/Web/JavaScript)で記述されています。 アクションレコーダーは、TypeScript のコード (JavaScript のスーパーセット) を生成します。 Office スクリプトドキュメントでは TypeScript を使用していますが、JavaScript を使用した方が快適な場合は、その代わりに使用できます。

Office スクリプトには、主に自己完結型のコード部分があります。 TypeScript の機能のごく一部のみが使用されます。 そのため、TypeScript の複雑な部分を学ばずにスクリプトを編集できます。 コードエディターでは、コードのインストール、コンパイル、実行も処理されるため、スクリプト自体について心配する必要はありません。 以前のプログラミング知識がなくても、言語を理解し、スクリプトを作成することができます。 ただし、プログラミングに慣れていない場合は、Office スクリプトを続行する前に、いくつかの基本事項を理解することをお勧めします。

- JavaScript の基本事項について説明します。 変数、制御フロー、関数、データ型などの概念に慣れている必要があります。 [Mozilla には、JavaScript に関する優れた総合的なチュートリアルが用意](https://developer.mozilla.org/docs/Web/JavaScript/Guide/Introduction)されています。
- TypeScript の種類について説明します。 TypeScript は、コンパイル時に適切な型をメソッドの呼び出しと割り当てに使用することによって、JavaScript で構築されます。 [インターフェイス](https://www.typescriptlang.org/docs/handbook/interfaces.html)、[クラス](https://www.typescriptlang.org/docs/handbook/classes.html)、[型の推論](https://www.typescriptlang.org/docs/handbook/type-inference.html)、および型の[互換性](https://www.typescriptlang.org/docs/handbook/type-compatibility.html)に関する TypeScript ドキュメントは、最も有用です。

## <a name="office-scripts-javascript-api"></a>Office スクリプト JavaScript API

Office スクリプトは、office[アドイン](/office/dev/add-ins/overview/index)用の Office JavaScript api である特別なバージョンを使用します。2つの Api には類似点がありますが、2つのプラットフォーム間でコードを移植できるとは想定しないでください。 2つのプラットフォーム間の相違点については、「 [Office スクリプトと Office アドインの相違点](../resources/add-ins-differences.md#apis)」を参照してください。 スクリプトで使用可能なすべての Api は、「 [Office スクリプト API リファレンス」ドキュメント](/javascript/api/office-scripts/overview)で確認できます。

## <a name="intellisense"></a>インテリジェンス

IntelliSense は、スクリプトを編集するときの入力ミスや構文エラーを防止するために役立つコードエディターの機能です。 入力したオブジェクトとフィールド名、および各 API のインラインドキュメントが表示されます。

Excel コードエディターは、Visual Studio Code と同じ IntelliSense エンジンを使用します。 この機能の詳細については、 [Visual Studio Code の IntelliSense 機能](https://code.visualstudio.com/docs/editor/intellisense#_intellisense-features)を参照してください。

## <a name="external-library-support"></a>外部ライブラリのサポート

Office スクリプトでは、外部のサードパーティの JavaScript ライブラリの使用はサポートされていません。 現在、スクリプトから Office スクリプト Api 以外のライブラリを呼び出すことはできません。 [Math](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Math)などの[組み込みの JavaScript オブジェクト](../develop/javascript-objects.md)には、まだアクセスできます。

## <a name="browser-support"></a>ブラウザのサポート

Office スクリプト[は、web 用の office をサポート](https://support.microsoft.com/office/ad1303e0-a318-47aa-b409-d3a5eb44e452)する任意のブラウザーで動作します。 ただし、一部の JavaScript 機能は Internet Explorer 11 (IE 11) ではサポートされていません。 ES6 以降で導入された機能は、IE 11 で[は](https://www.w3schools.com/Js/js_es6.asp)動作しません。 組織内のユーザーが依然としてそのブラウザーを使用している場合は、その環境でスクリプトを共有するときに必ずテストしてください。

## <a name="see-also"></a>関連項目

- [Office スクリプト API リファレンス](/javascript/api/office-scripts/overview)
- [Office スクリプトのトラブルシューティング](../testing/troubleshooting.md)
- [Office スクリプトでの組み込みの JavaScript オブジェクトの使用](../develop/javascript-objects.md)
