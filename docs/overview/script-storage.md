---
title: Officeスクリプト ファイルのストレージと所有権
description: スクリプトを管理者Officeに格納し、所有者Microsoft OneDrive転送する方法に関する情報。
ms.date: 06/04/2021
ms.localizationpriority: medium
ms.openlocfilehash: 6b82fc041c97288feefa85f2a9c9efeab0cb5705
ms.sourcegitcommit: 5ec904cbb1f2cc00a301a5ba7ccb8ae303341267
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 09/18/2021
ms.locfileid: "59447451"
---
# <a name="office-scripts-file-storage-and-ownership"></a>Officeスクリプト ファイルのストレージと所有権

Officeスクリプトは、ユーザーの **ファイルに .osts** ファイルとしてMicrosoft OneDrive。 ブックとは別に格納されます。 他のユーザーにアクセス権を与えるために[、スクリプトをブックExcelします](excel.md#share-scripts)。 つまり、スクリプトをファイルにリンクしているのではなく、ファイルとリンクします。 Excelファイルにアクセスできるユーザーは、スクリプトの表示、実行、またはコピーの作成を行えます。

スクリプトを共有しない限り、他のユーザーはスクリプトにアクセスできません。 ユーザー OneDrive設定は、すべてのスクリプト **.osts** ファイルに対する共有アクセスとアクセス許可を、すべてのスクリプト設定にExcelします。 スクリプトは、ローカル ディスクまたはカスタム クラウドの場所からリンクできません。 Officeスクリプトは、スクリプトを認識して実行できるのは、スクリプトが OneDriveフォルダーにある場合、またはブックと共有されている場合のみです。

## <a name="file-storage"></a>ファイルの記憶域

スクリプトOfficeは、ユーザーのサーバーにOneDrive。 **.osts ファイル** は **、/Documents/Officeフォルダーにあります**。 ファイルの名前の変更や削除など、これらの **.osts** ファイルに対して行われた編集は、コード エディターとスクリプト ギャラリーに反映されます。

ブックの 1 つと共有されているスクリプトは、スクリプト作成者のデータベースに残OneDrive。 共有スクリプトを OneDrive で実行すると、ローカル フォルダーまたはローカル フォルダーにはコピー Excel。 コード **エディターの [コピー** を作成] ボタンをクリックすると、スクリプトの別のコピーがユーザーのページにOneDrive。 コピーに対する変更は、元のスクリプトには影響を与えかねない。

### <a name="restore-deleted-scripts"></a>削除されたスクリプトを復元する

スクリプトを削除すると、Excelごみ箱にOneDriveされます。 削除されたスクリプトを復元するには、「削除されたファイルまたはフォルダーを復元する」に記載されている手順に[従](https://support.microsoft.com/office/949ada80-0026-4db3-a953-c99083e6a84f)OneDrive。 **.osts ファイルを復元すると**、[すべてのスクリプト] リスト **に戻** されます。

削除されたスクリプトはブックと共有されません。 スクリプトを復元しても、スクリプト **へのアクセス** は保持されます。 スクリプトを再び共有する必要があります。

復元されたスクリプトは、引き続き必要にPower Automateします。 フロー コネクタを再作成する必要はない。

## <a name="file-ownership-and-retention"></a>ファイルの所有権と保持

Officeスクリプトは、ユーザーのデータベースにOneDrive。 ユーザーは、ユーザーが指定した保持ポリシーと削除ポリシー Microsoft OneDrive。 組織から削除されるユーザーによって作成および共有されたスクリプトを処理する方法については、[OneDrive の保持と削除](/onedrive/retention-and-deletion)を参照してください。

編集中、ファイルはブラウザーに一時的に保存されます。 スクリプトを保存してからウィンドウを閉じる前に、Excelの場所に保存OneDriveがあります。 編集後にファイルを保存することを忘れないでください。それ以外の場合、これらの編集はブラウザーのバージョンのファイルにのみ含されます。

## <a name="audit-office-scripts-usage-at-the-admin-level"></a>管理者Officeスクリプトの使用状況を監査する

コンプライアンス センターで監査ログOfficeスクリプトを使用しているテナントを確認します。 このツールの使い方については、「セキュリティ コンプライアンス センターで監査ログを検索する」 [を&してください](/microsoft-365/compliance/search-the-audit-log-in-security-and-compliance?view=o365-worldwide&preserve-view=true#search-the-audit-log)。

検索ツールでスクリプトをOfficeするユーザーを検索するには、[ファイル]、フォルダー、またはサイト フィールド `.osts` **に追加** します。 これにより、スクリプト ファイル拡張子が Officeファイルが検索されます。 組織内のユーザーが [スクリプト] 機能を使用Office、監査ログの検索結果にユーザー アクティビティが表示されます。

> [!NOTE]
> スクリプトの実行は現在ログに記録されません。 作成、表示、および変更のアクションだけがログに記録されます。

## <a name="see-also"></a>関連項目

- [Excel on the web での Office スクリプトの共有](https://support.microsoft.com/office/226eddbc-3a44-4540-acfe-fccda3d1122b)
- [Office スクリプトのトラブルシューティング](../testing/troubleshooting.md)
- [M365 での Office スクリプトの設定](/microsoft-365/admin/manage/manage-office-scripts-settings)
- [Office スクリプトの効果を元に戻す](../testing/undo.md)
