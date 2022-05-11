---
title: Office スクリプト ファイルのストレージと所有権
description: Office スクリプトをMicrosoft OneDriveに格納し、所有者間で転送する方法に関する情報。
ms.date: 05/10/2022
ms.localizationpriority: medium
ms.openlocfilehash: 5e2bc89db54ee5520c3b911ebd0f182777a78e2b
ms.sourcegitcommit: 8ae932e8b4e521fec8576ab16126eb9fe22a8dd7
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 05/11/2022
ms.locfileid: "65310758"
---
# <a name="office-scripts-file-storage-and-ownership"></a>Office スクリプト ファイルのストレージと所有権

Office スクリプトは、Microsoft OneDriveに **.osts** ファイルとして格納されます。 ブックとは別に格納されます。 他のユーザーにアクセス権を付与するには、[スクリプトをExcel ブックと共有します](excel.md#share-office-scripts)。 つまり、スクリプトを添付せず、ファイルとリンクしています。 Excel ファイルにアクセスできるユーザーは、スクリプトを表示、実行、または作成することもできます。

スクリプトを共有しない限り、他のユーザーはアクセスできません。 OneDrive設定は、Excel設定に関係なく、すべてのスクリプト **.osts** ファイルに対する共有アクセスとアクセス許可を制御します。 スクリプトをローカル ディスクまたはカスタム クラウドの場所からリンクすることはできません。 Office スクリプトは、OneDrive フォルダー内にある場合、またはブックと共有されている場合にのみ、スクリプトを認識して実行します。

## <a name="file-storage"></a>ファイルの記憶域

スクリプトはOneDriveに格納Office。 **.osts** ファイルは、**/Documents/Office Scripts/** フォルダーにあります。 これらの **.osts** ファイルに対して行われた編集 (ファイルの名前変更や削除など) は、コード エディターとスクリプト ギャラリーに反映されます。

ブックの 1 つと共有されるスクリプトは、スクリプト作成者のOneDriveに残ります。 共有スクリプトを Excel で実行すると、ローカル フォルダーまたはOneDrive フォルダーにはコピーされません。 コード エディターの **[コピーを作成**] ボタンをクリックすると、スクリプトの別のコピーがOneDriveに保存されます。 コピーに対する変更は、元のスクリプトには影響しません。

### <a name="restore-deleted-scripts"></a>削除されたスクリプトを復元する

Excelでスクリプトを削除すると、OneDriveのごみ箱に移動します。 削除されたスクリプトを復元するには、「[OneDriveで削除されたファイルまたはフォルダーを復元する」に記載されている手順に](https://support.microsoft.com/office/949ada80-0026-4db3-a953-c99083e6a84f)従います。 **.osts** ファイルを復元すると、[**すべてのスクリプト**] リストに返されます。

削除されたスクリプトは、ブックと共有されません。 スクリプトを復元しても、スクリプトへのアクセスは保持 **されません** 。 スクリプトをもう一度共有する必要があります。

復元されたスクリプトは、Power Automate フローで引き続き期待どおりに動作します。 フロー コネクタを再作成する必要はありません。

## <a name="file-ownership-and-retention"></a>ファイルの所有権と保持期間

Office スクリプトは、ユーザーのOneDriveに格納されます。 これらのポリシーは、Microsoft OneDriveで指定された保持ポリシーと削除ポリシーに従います。 組織から削除されるユーザーによって作成および共有されたスクリプトを処理する方法については、[OneDrive の保持と削除](/onedrive/retention-and-deletion)を参照してください。

編集中、ファイルは一時的にブラウザーに保存されます。 Excel ウィンドウを閉じる前にスクリプトを保存して、OneDriveの場所に保存する必要があります。 編集後にファイルを保存することを忘れないでください。それ以外の場合、これらの編集はブラウザーのバージョンのファイルにのみ含まれます。

## <a name="audit-office-scripts-usage-at-the-admin-level"></a>管理者レベルでのスクリプトの使用Office監査する

コンプライアンス センターの監査ログでOffice スクリプトを使用しているテナントを確認します。 このツールの使用方法については、 [セキュリティ & コンプライアンス センターの監査ログを検索](/microsoft-365/compliance/search-the-audit-log-in-security-and-compliance?view=o365-worldwide&preserve-view=true#search-the-audit-log)するを参照してください。

検索ツールでOffice スクリプトを使用しているユーザーを検索するには、[**ファイル]、[フォルダー]、または [サイト**] フィールドに追加`.osts`します。 これにより、Office Scripts ファイル拡張子を持つすべてのファイルが検索されます。 組織内のだれかが Office スクリプト機能を使用した場合、ユーザー アクティビティは監査ログの検索結果に表示されます。

## <a name="see-also"></a>関連項目

- [Excel on the web での Office スクリプトの共有](https://support.microsoft.com/office/226eddbc-3a44-4540-acfe-fccda3d1122b)
- [Office スクリプトのトラブルシューティング](../testing/troubleshooting.md)
- [M365 での Office スクリプトの設定](/microsoft-365/admin/manage/manage-office-scripts-settings)
- [Office スクリプトの効果を元に戻す](../testing/undo.md)
