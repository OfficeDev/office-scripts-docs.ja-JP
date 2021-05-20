---
title: Officeスクリプト ファイルの保存と所有権
description: Microsoft OneDriveでスクリプトOffice格納し、所有者間で転送する方法に関する情報。
ms.date: 05/17/2021
localization_priority: Normal
ms.openlocfilehash: 556d784dc1fe64873866c49ab2726a4c68abc1a7
ms.sourcegitcommit: 4687693f02fc90a57ba30c461f35046e02e6f5fb
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 05/19/2021
ms.locfileid: "52545802"
---
# <a name="office-scripts-file-storage-and-ownership"></a>Officeスクリプト ファイルの保存と所有権

Officeスクリプトは、Microsoft OneDriveに **.osts** ファイルとして保存されます。 ブックとは別に保存されます。 他のユーザーにアクセス権を付与するには、[スクリプトを Excel ブックと共有](excel.md#sharing-scripts)します。 つまり、スクリプトを添付せず、ファイルにリンクします。 Excelファイルにアクセスできるユーザーは、スクリプトの表示、実行、またはコピーの作成も可能です。

スクリプトを共有しない限り、他の誰もスクリプトにアクセスできません。 OneDrive設定は、すべてのスクリプト **.osts** ファイルに対する共有アクセスとアクセス許可を、Excel設定とは無関係に制御します。 スクリプトは、ローカル ディスクまたはカスタム クラウドの場所からリンクすることはできません。 Officeスクリプトは、スクリプトがOneDrive フォルダー内にあるか、ブックと共有されている場合にのみスクリプトを認識して実行します。

## <a name="file-storage"></a>ファイルの記憶域

スクリプトはOneDriveに格納Office。 **osts** ファイルは **、/ドキュメント/Office スクリプト/** フォルダーにあります。 これらの **.osts** ファイルに対する編集 (ファイルの名前の変更や削除など) は、コード エディターおよびスクリプト ギャラリーに反映されます。

ワークブックの 1 つと共有されるスクリプトは、スクリプト作成者のOneDriveに残ります。 Excelで共有スクリプトを実行しても、ローカルフォルダまたはOneDrive フォルダにコピーされません。 コード エディター **の [コピーの作成**] ボタンをクリックすると、スクリプトのコピーがOneDriveに保存されます。 コピーに対する変更は、元のスクリプトには影響しません。

## <a name="file-ownership-and-retention"></a>ファイルの所有権と保存

OfficeスクリプトはユーザーのOneDriveに格納されます。 Microsoft OneDriveで指定された保存と削除のポリシーに従います。 組織から削除されるユーザーによって作成および共有されたスクリプトを処理する方法については、[OneDrive の保持と削除](/onedrive/retention-and-deletion)を参照してください。

編集中、ファイルは一時的にブラウザに保存されます。 Excel ウィンドウを閉じる前にスクリプトを保存して、OneDrive場所に保存する必要があります。 編集後にファイルを保存することを忘れないでください。

## <a name="see-also"></a>関連項目

- [Excel on the web での Office スクリプトの共有](https://support.microsoft.com/office/sharing-office-scripts-in-excel-for-the-web-226eddbc-3a44-4540-acfe-fccda3d1122b)
- [Office スクリプトのトラブルシューティング](../testing/troubleshooting.md)
- [M365 での Office スクリプトの設定](https://support.office.com/article/office-scripts-settings-in-m365-19d3c51a-6ca2-40ab-978d-60fa49554dcf)
- [Office スクリプトの効果を元に戻す](../testing/undo.md)
