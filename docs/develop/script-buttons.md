---
title: ボタンを使用してExcelで Office スクリプトを実行する
description: ExcelのスクリプトOffice制御するブックにボタンを追加します。
ms.topic: overview
ms.date: 05/09/2022
ms.localizationpriority: medium
ms.openlocfilehash: fde34d62f9abe897a8b93195ab37a75cfc73f619
ms.sourcegitcommit: 34c7740c9bff0e4c7426e01029f967724bfee566
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 05/13/2022
ms.locfileid: "65393685"
---
# <a name="run-office-scripts-in-excel-with-buttons"></a>ボタンを使用してExcelで Office スクリプトを実行する

ブックにスクリプト ボタンを追加することにより、同僚がスクリプトを見つけて実行するのに役立ちます。

:::image type="content" source="../images/run-from-button.png" alt-text="クリックするとスクリプトを実行するワークシート内のボタン。":::

## <a name="create-script-buttons"></a>スクリプト ボタンの作成

任意のスクリプトで、スクリプトの詳細ページまたはコード エディターの作業ウィンドウの **[その他のオプション (...)]** メニューに移動し、[ **追加] ボタン** を選択します。 これにより、選択した場合に関連付けられたスクリプトを実行するブックにボタンが作成されます。 また、スクリプトをブックと共有します。ブックへの書き込みアクセス許可を持つすべてのユーザーが、便利な自動処理を使用できます。

次のスクリーンショットは、[**ピボットテーブルの作成**] というタイトルのスクリプトのスクリプトの詳細ページを示しています。[**その他のオプション ] (...)** メニューの [**追加] ボタン** オプションが強調表示されています。

:::image type="content" source="../images/add-button.png" alt-text="スクリプトの詳細ページ メニューの [追加] ボタンオプション。":::

## <a name="remove-script-buttons"></a>スクリプト ボタンを削除する

ボタンを使用してスクリプトの共有を停止するには、スクリプトの詳細ページの **[その他のオプション (...)]** メニューに移動し、[ **共有の停止**] を選択します。 これにより、スクリプトを実行しているすべてのボタンが削除されます。 1 つのボタンを削除すると、操作が元に戻された場合や、ボタンが切り取って貼り付けられた場合でも、その 1 つのボタンからスクリプトが削除されます。

## <a name="script-buttons-with-excel-on-windows"></a>WindowsのExcelを含むスクリプト ボタン

これらのスクリプト ボタンは Windows でも機能します。 Excel on the webでボタンを作成し、Windowsのユーザーはボタンをクリックしてスクリプトを実行できます。 WindowsのExcelでスクリプトを編集することはできません。 スクリプトは、Excel on the webでのみ編集できます。

一部のOffice スクリプト API は、Windows (特に古いビルド) のExcelではサポートされない場合があります。 これには、Web 専用機能用の新しい API と API が含まれます。 スクリプトにサポートされていない API が含まれている場合、スクリプトは実行されません。代わりに、[**スクリプトの実行状態**] 作業ウィンドウに警告メッセージが表示されます。"このスクリプトは、現在、Excel for the webで実行する必要があります。 ブラウザーでブックを開いてからもう一度試すか、スクリプト所有者に問い合わせてください。  

> [!IMPORTANT]
> スクリプト ボタンでは[、webView2](/deployoffice/webview2-install) がWindowsでExcelを操作する必要があります。 これは、デスクトップ上の最新バージョンのExcelで既定でインストールされますが、スクリプト ボタンをクリックできない場合は、[WebView2 ランタイムのダウンロード](https://developer.microsoft.com/en-us/microsoft-edge/webview2/#download-section)とブラウザー エンジンのダウンロードに関するページを参照してください。
