---
title: Office スクリプトファイルの保存と所有権
description: Office スクリプトが Microsoft OneDrive に格納され、所有者間で転送される方法について説明します。
ms.date: 11/13/2020
localization_priority: Normal
ms.openlocfilehash: 648f3b2cf7e7d8d3bab2cf07a090e116e267a99a
ms.sourcegitcommit: 82d3c0ef1e187bcdeceb2b5fc3411186674fe150
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 11/18/2020
ms.locfileid: "49346866"
---
# <a name="office-scripts-file-storage-and-ownership"></a>Office スクリプトファイルの保存と所有権

Office スクリプトは、Microsoft OneDrive に **ost** ファイルとして保存されます。 これにより、スクリプトは特定のブックの外に存在することができます。 OneDrive の設定は、すべてのスクリプトの **ost** ファイルの共有アクセスとアクセス許可を制御します。Excel のすべての設定に依存しません。

## <a name="file-storage"></a>ファイルの記憶域

Office スクリプトは OneDrive に保存されています。 この **ost** ファイルは、/ **ドキュメント/Office スクリプト/** フォルダーにあります。 ファイル名の変更や削除など、これらの **ost** ファイルに対して行われた編集は、コードエディターとスクリプトギャラリーに反映されます。

ブックの1つと共有されているスクリプトは、スクリプト作成者の OneDrive に残ります。 これらのフォルダーは、Excel で共有スクリプトを実行しても、ローカルフォルダーや OneDrive フォルダーにはコピーされません。 コードエディターの [ **コピーの作成** ] ボタンをクリックすると、スクリプトの別のコピーが OneDrive に保存されます。 コピーを変更しても、元のスクリプトには影響しません。

### <a name="script-folders"></a>スクリプトフォルダー

OneDrive にフォルダーを追加すると、スクリプトを整理できます。 / **ドキュメント/Office スクリプト/** の下のフォルダーは、コードエディターの [ **マイスクリプト** ] セクションの下に表示されます。 これらのフォルダーは、コードエディターを使用して作成または削除できないことに注意してください。 同様に、スクリプトはフォルダーに配置したり、コードエディターを使用してフォルダー間で移動したりすることはできません。

![[コードエディター] 作業ウィンドウに表示されているフォルダー内の一部のスクリプト](../images/script-folders.png)

## <a name="file-ownership-and-retention"></a>ファイルの所有権と保持

Office スクリプトは、ユーザーの OneDrive に保存されます。 これらは、Microsoft OneDrive で指定されているアイテム保持ポリシーと削除ポリシーに従います。 組織から削除されるユーザーによって作成および共有されたスクリプトを処理する方法については、[OneDrive の保持と削除](/onedrive/retention-and-deletion)を参照してください。

## <a name="see-also"></a>関連項目

- [Excel on the web での Office スクリプトの共有](https://support.microsoft.com/office/sharing-office-scripts-in-excel-for-the-web-226eddbc-3a44-4540-acfe-fccda3d1122b)
- [Office スクリプトのトラブルシューティング](../testing/troubleshooting.md)
- [M365 での Office スクリプトの設定](https://support.office.com/article/office-scripts-settings-in-m365-19d3c51a-6ca2-40ab-978d-60fa49554dcf)
- [Office スクリプトの効果を元に戻す](../testing/undo.md)
