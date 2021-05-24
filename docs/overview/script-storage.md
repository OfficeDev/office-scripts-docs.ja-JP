---
title: Officeスクリプト ファイルのストレージと所有権
description: スクリプトを管理者Officeに格納し、所有者Microsoft OneDrive転送する方法に関する情報。
ms.date: 05/17/2021
localization_priority: Normal
ms.openlocfilehash: 556d784dc1fe64873866c49ab2726a4c68abc1a7
ms.sourcegitcommit: 4687693f02fc90a57ba30c461f35046e02e6f5fb
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 05/19/2021
ms.locfileid: "52545802"
---
# <a name="office-scripts-file-storage-and-ownership"></a>Officeスクリプト ファイルのストレージと所有権

Officeスクリプトは、ユーザーの **ファイルに .osts** ファイルとしてMicrosoft OneDrive。 ブックとは別に格納されます。 他のユーザーにアクセス権を与えるために[、スクリプトをブックExcelします](excel.md#sharing-scripts)。 つまり、スクリプトをファイルにリンクしているのではなく、ファイルとリンクします。 Excelファイルにアクセスできるユーザーは、スクリプトの表示、実行、またはコピーの作成を行えます。

スクリプトを共有しない限り、他のユーザーはスクリプトにアクセスできません。 ユーザー OneDrive設定は、すべてのスクリプト **.osts** ファイルに対する共有アクセスとアクセス許可を、すべてのスクリプト設定にExcelします。 スクリプトは、ローカル ディスクまたはカスタム クラウドの場所からリンクできません。 Officeスクリプトは、スクリプトを認識して実行できるのは、スクリプトが OneDriveフォルダーにある場合、またはブックと共有されている場合のみです。

## <a name="file-storage"></a>ファイルの記憶域

スクリプトOfficeは、ユーザーのサーバーにOneDrive。 **.osts ファイル** は **、/Documents/Officeフォルダーにあります**。 ファイルの名前の変更や削除など、これらの **.osts** ファイルに対して行われた編集は、コード エディターとスクリプト ギャラリーに反映されます。

ブックの 1 つと共有されているスクリプトは、スクリプト作成者のデータベースに残OneDrive。 共有スクリプトを OneDrive で実行すると、ローカル フォルダーまたはローカル フォルダーにはコピー Excel。 コード **エディターの [コピー** を作成] ボタンをクリックすると、スクリプトの別のコピーがユーザーのページにOneDrive。 コピーに対する変更は、元のスクリプトには影響を与えかねない。

## <a name="file-ownership-and-retention"></a>ファイルの所有権と保持

Officeスクリプトは、ユーザーのデータベースにOneDrive。 ユーザーは、ユーザーが指定した保持ポリシーと削除ポリシー Microsoft OneDrive。 組織から削除されるユーザーによって作成および共有されたスクリプトを処理する方法については、[OneDrive の保持と削除](/onedrive/retention-and-deletion)を参照してください。

編集中、ファイルはブラウザーに一時的に保存されます。 スクリプトを保存してからウィンドウを閉じる前に、Excelの場所に保存OneDriveがあります。 編集後にファイルを保存することを忘れないでください。それ以外の場合、これらの編集はブラウザーのバージョンのファイルにのみ含されます。

## <a name="see-also"></a>関連項目

- [Excel on the web での Office スクリプトの共有](https://support.microsoft.com/office/sharing-office-scripts-in-excel-for-the-web-226eddbc-3a44-4540-acfe-fccda3d1122b)
- [Office スクリプトのトラブルシューティング](../testing/troubleshooting.md)
- [M365 での Office スクリプトの設定](https://support.office.com/article/office-scripts-settings-in-m365-19d3c51a-6ca2-40ab-978d-60fa49554dcf)
- [Office スクリプトの効果を元に戻す](../testing/undo.md)
