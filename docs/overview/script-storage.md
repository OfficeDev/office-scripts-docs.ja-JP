---
title: Officeスクリプト ファイルのストレージと所有権
description: スクリプトを管理者Officeに格納し、所有者Microsoft OneDrive転送する方法に関する情報。
ms.date: 11/13/2020
localization_priority: Normal
ms.openlocfilehash: 47b732399c3068bea78b027e01324bbd73a83bc7
ms.sourcegitcommit: f7a7aebfb687f2a35dbed07ed62ff352a114525a
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 05/06/2021
ms.locfileid: "52232530"
---
# <a name="office-scripts-file-storage-and-ownership"></a>Officeスクリプト ファイルのストレージと所有権

Officeスクリプトは、ユーザーの **ファイルに .osts** ファイルとしてMicrosoft OneDrive。 これにより、スクリプトを特定のブックの外部に存在できます。 ユーザー OneDrive設定は、すべてのスクリプト **.osts** ファイルの共有アクセスとアクセス許可を制御します。任意の設定にExcelします。

## <a name="file-storage"></a>ファイルの記憶域

スクリプトOfficeは、ユーザーのサーバーにOneDrive。 **.osts ファイル** は **、/Documents/Officeフォルダーにあります**。 ファイルの名前の変更や削除など、これらの **.osts** ファイルに対して行われた編集は、コード エディターとスクリプト ギャラリーに反映されます。

ブックの 1 つと共有されているスクリプトは、スクリプト作成者のデータベースに残OneDrive。 共有スクリプトを OneDrive で実行すると、ローカル フォルダーまたはローカル フォルダーにはコピー Excel。 コード **エディターの [コピー** を作成] ボタンをクリックすると、スクリプトの別のコピーがユーザーのページにOneDrive。 コピーに対する変更は、元のスクリプトには影響を与えかねない。

### <a name="script-folders"></a>スクリプト フォルダー

フォルダーをフォルダーに追加OneDriveスクリプトを整理するのに役立ちます。 **/Documents/Office スクリプト/ の** 下のフォルダーは、コード エディターの **[マイ スクリプト**] セクションに表示されます。 これらのフォルダーは、コード エディターを使用して作成または削除することはできません。 同様に、スクリプトをフォルダーに配置したり、コード エディターを使用してフォルダー間で移動したりすることはできません。

:::image type="content" source="../images/script-folders.png" alt-text="作業ウィンドウに表示されるフォルダーに含まれるスクリプトを表示するコード エディターの [新しいスクリプト] ダイアログ":::

## <a name="file-ownership-and-retention"></a>ファイルの所有権と保持

Officeスクリプトは、ユーザーのデータベースにOneDrive。 ユーザーは、ユーザーが指定した保持ポリシーと削除ポリシー Microsoft OneDrive。 組織から削除されるユーザーによって作成および共有されたスクリプトを処理する方法については、[OneDrive の保持と削除](/onedrive/retention-and-deletion)を参照してください。

## <a name="see-also"></a>こちらもご覧ください

- [Excel on the web での Office スクリプトの共有](https://support.microsoft.com/office/sharing-office-scripts-in-excel-for-the-web-226eddbc-3a44-4540-acfe-fccda3d1122b)
- [Office スクリプトのトラブルシューティング](../testing/troubleshooting.md)
- [M365 での Office スクリプトの設定](https://support.office.com/article/office-scripts-settings-in-m365-19d3c51a-6ca2-40ab-978d-60fa49554dcf)
- [Office スクリプトの効果を元に戻す](../testing/undo.md)
