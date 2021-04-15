---
title: Office スクリプト ファイルのストレージと所有権
description: Microsoft OneDrive にOfficeし、所有者間で転送する方法に関する情報。
ms.date: 11/13/2020
localization_priority: Normal
ms.openlocfilehash: bd868c1dbfd0b33d3cd9fc4ee774c654d86f9b07
ms.sourcegitcommit: 45ffe3dbd2c834b78592ad35928cf8096f5e80bc
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 04/14/2021
ms.locfileid: "51755106"
---
# <a name="office-scripts-file-storage-and-ownership"></a>Office スクリプト ファイルのストレージと所有権

Officeスクリプトは **、Microsoft OneDrive に .osts** ファイルとして保存されます。 これにより、スクリプトを特定のブックの外部に存在できます。 OneDrive 設定は、すべてのスクリプト **.osts** ファイルの共有アクセスとアクセス許可を制御します。Excel の設定とは独立しています。

## <a name="file-storage"></a>ファイルの記憶域

スクリプトOffice OneDrive に保存されます。 **.osts ファイル** は **、/Documents/Officeフォルダーにあります**。 ファイルの名前の変更や削除など、これらの **.osts** ファイルに対して行われた編集は、コード エディターとスクリプト ギャラリーに反映されます。

ブックの 1 つと共有されているスクリプトは、スクリプト作成者の OneDrive に残ります。 Excel で共有スクリプトを実行すると、ローカル フォルダーまたは OneDrive フォルダーにはコピーされません。 コード **エディターの [コピーの** 作成] ボタンは、OneDrive にスクリプトの個別のコピーを保存します。 コピーに対する変更は、元のスクリプトには影響を与えかねない。

### <a name="script-folders"></a>スクリプト フォルダー

OneDrive にフォルダーを追加すると、スクリプトを整理し続けるのに役立ちます。 **/Documents/Office スクリプト/ の下の** フォルダーは、コード エディターの **[マイ スクリプト**] セクションに表示されます。 これらのフォルダーは、コード エディターを使用して作成または削除することはできません。 同様に、スクリプトをフォルダーに配置したり、コード エディターを使用してフォルダー間で移動したりすることはできません。

:::image type="content" source="../images/script-folders.png" alt-text="作業ウィンドウに表示されるフォルダーに含まれるスクリプトを表示するコード エディターの [新しいスクリプト] ダイアログ。":::

## <a name="file-ownership-and-retention"></a>ファイルの所有権と保持

Officeスクリプトは、ユーザーの OneDrive に格納されます。 Microsoft OneDrive で指定された保持ポリシーと削除ポリシーに従います。 組織から削除されるユーザーによって作成および共有されたスクリプトを処理する方法については、[OneDrive の保持と削除](/onedrive/retention-and-deletion)を参照してください。

## <a name="see-also"></a>関連項目

- [Excel on the web での Office スクリプトの共有](https://support.microsoft.com/office/sharing-office-scripts-in-excel-for-the-web-226eddbc-3a44-4540-acfe-fccda3d1122b)
- [Office スクリプトのトラブルシューティング](../testing/troubleshooting.md)
- [M365 での Office スクリプトの設定](https://support.office.com/article/office-scripts-settings-in-m365-19d3c51a-6ca2-40ab-978d-60fa49554dcf)
- [Office スクリプトの効果を元に戻す](../testing/undo.md)
