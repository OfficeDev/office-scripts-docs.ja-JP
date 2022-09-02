---
title: Office スクリプト ファイルのストレージと所有権
description: Office スクリプトが Microsoft OneDrive に格納され、所有者間で転送される方法に関する情報。
ms.date: 08/31/2022
ms.localizationpriority: medium
ms.openlocfilehash: 573f65f299c29b4f481c9a2e23ebe7e36181706b
ms.sourcegitcommit: a6504f8b0d6b717457c6e0b5306c35ad3900914e
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 09/02/2022
ms.locfileid: "67572508"
---
# <a name="office-scripts-file-storage-and-ownership"></a>Office スクリプト ファイルのストレージと所有権

Office スクリプトは、Microsoft OneDrive または SharePoint フォルダーに **.osts** ファイルとして格納されます。 ブックとは別に格納されます。 SharePoint サイトの外部にいるユーザーにスクリプトへのアクセス権を付与するには、 [スクリプトを Excel ブックと共有します](excel.md#share-office-scripts)。 つまり、スクリプトを添付せず、ファイルとリンクしています。 Excel ファイルにアクセスできるユーザーは、スクリプトを表示、実行、または作成することもできます。

Excel は、スクリプトが OneDrive フォルダー、Sharepoint フォルダー、またはブックと共有されている場合にのみ、スクリプトを認識して実行します。

## <a name="onedrive"></a>OneDrive

既定の動作では、Office スクリプトは OneDrive に格納されます。 **.osts** ファイルは、**/Documents/Office スクリプト/** フォルダーにあります。 これらの **.osts** ファイルに対して行われた編集 (ファイルの名前変更や削除など) は、コード エディターとスクリプト ギャラリーに反映されます。

ブックの 1 つと共有されるスクリプトは、スクリプト作成者の OneDrive に残ります。 Excel で共有スクリプトを実行しても、ローカルフォルダーまたは OneDrive フォルダーにはコピーされません。 コード エディターの **[コピーの作成** ] ボタンは、OneDrive にスクリプトの別のコピーを保存します。 コピーに対する変更は、元のスクリプトには影響しません。

個人用スクリプトを共有しない限り、他のユーザーはアクセスできません。 OneDrive の設定は、Excel 設定とは関係なく、すべてのスクリプト **.osts** ファイルに対する共有アクセスとアクセス許可を制御します。 スクリプトをローカル ディスクまたはカスタム クラウドの場所からリンクすることはできません。

## <a name="sharepoint"></a>SharePoint

SharePoint サイトに保存される Office スクリプトは、チームが所有します。 適切なアクセス権を持つ自分と組織のメンバーは、SharePoint からスクリプトを実行および編集できます。 これらのスクリプトは、[ **自動化** ] タブの [スクリプト ギャラリー] にも表示されます。

SharePoint からスクリプトを読み込むには、[ **すべてのスクリプト** ] に移動し、一覧の下部にある [ **その他のスクリプトを表示** ] を選択します。 これにより、アクセスできる任意の SharePoint サイトから **.osts** ファイルを選択できるファイル ピッカーが表示されます。 既に開いている SharePoint のスクリプトは、最近使用したスクリプトの一覧に表示されることに注意してください。

スクリプトを SharePoint に保存するには、 **その他のオプション (...)** メニューに移動し、[ **名前を付けて保存]** を選択します。 これにより、SharePoint サイト内のフォルダーを選択できるファイル ピッカーが開きます。 新しい場所に保存すると、その場所にスクリプトのコピーが作成されます。 元のバージョンは引き続き OneDrive またはその他の SharePoint の場所にあります。

> [!IMPORTANT]
> [外部呼び出し](../develop/external-calls.md)を含むスクリプトを SharePoint から実行することはできません。 "現時点では、SharePoint サイトに保存されたスクリプトではネットワーク アクセス呼び出しはサポートされていません" というエラーが表示されます。

> [!IMPORTANT]
> Power Automate では、現時点では SharePoint に格納されているスクリプトはサポート **されていません** 。

## <a name="restore-deleted-scripts"></a>削除されたスクリプトを復元する

Excel でスクリプトを削除すると、OneDrive または SharePoint のごみ箱に移動します。 削除されたスクリプトを復元するには、「 [SharePoint と OneDrive で職場または学校の紛失、削除、または破損したアイテムを回復する方法](https://support.microsoft.com/office/how-to-recover-missing-deleted-or-corrupted-items-in-sharepoint-and-onedrive-for-work-or-school-3d748edf-c072-46c9-81a4-4989056ebc87)」に記載されている手順に従います。 **.osts** ファイルを復元すると、[**すべてのスクリプト**] リストに返されます。

削除されたスクリプトは、ブックと共有されません。 スクリプトを復元しても、スクリプトへのアクセスは保持 **されません** 。 スクリプトをもう一度共有する必要があります。

復元されたスクリプトは、Power Automate フローで引き続き想定どおりに動作します。 フロー コネクタを再作成する必要はありません。

## <a name="file-ownership-and-retention"></a>ファイルの所有権と保持期間

Office スクリプトは、Microsoft OneDrive と Microsoft SharePoint で指定された保持ポリシーと削除ポリシーに従います。 組織から削除されたユーザーによって作成され、共有されたスクリプトを処理する方法については、「 [SharePoint と OneDrive のリテンション期間の詳細](/microsoft-365/compliance/retention-policies-sharepoint?view=o365-worldwide&preserve-view=true)」を参照してください。

編集中、ファイルは一時的にブラウザーに保存されます。 Excel ウィンドウを閉じる前にスクリプトを保存して、OneDrive の場所に保存する必要があります。 編集後にファイルを保存することを忘れないでください。それ以外の場合、これらの編集はブラウザーのバージョンのファイルにのみ含まれます。

## <a name="audit-office-scripts-usage-at-the-admin-level"></a>管理者レベルで Office スクリプトの使用状況を監査する

コンプライアンス センターの監査ログを使用して、組織内で Office スクリプトを使用しているユーザーを検出します。 監査ログの詳細については、 [セキュリティ & コンプライアンス センターの監査ログの検索](/microsoft-365/compliance/search-the-audit-log-in-security-and-compliance?view=o365-worldwide&preserve-view=true#search-the-audit-log)に関するページを参照してください。

管理者として Office スクリプト関連のアクティビティを具体的に監査するには、次の手順に従います。

1. InPrivate ブラウザー ウィンドウ (または Incognito またはその他のブラウザー固有の制限付き追跡モード) で、 [コンプライアンス センター](https://compliance.microsoft.com/)を開いてログインします。
1. **[監査**] ページに移動します。
1. *(1 回限り)* [ **検索** ] タブで、[ **ユーザーと管理者のアクティビティの記録を開始する**] を選択します。

    > [!IMPORTANT]
    > テナント全体のすべてのアクティビティが記録されるまでに、記録を有効にしてから 1 時間または 2 時間かかる場合があります。

1. 目的の検索オプションを設定し、 **Search** キーを押します。 **アクティビティ** をフィルター処理して **ブックでスクリプトを実行** し、スクリプトが実行された日時を確認します。 [ **ファイル]、[フォルダー]、または [サイト** ] フィールドをフィルター処理して `.osts`、 . これにより、組織内のだれがスクリプトを作成または変更しているかがわかります。

    :::image type="content" source="../images/audit-log-example.png" alt-text="&quot;ブックにスクリプトを実行する&quot; アクションや .osts ファイルのアップロードと変更など、いくつかの監査ログ検索結果行。":::

## <a name="see-also"></a>関連項目

- [Excel on the web での Office スクリプトの共有](https://support.microsoft.com/office/226eddbc-3a44-4540-acfe-fccda3d1122b)
- [Office スクリプトのトラブルシューティング](../testing/troubleshooting.md)
- [M365 での Office スクリプトの設定](/microsoft-365/admin/manage/manage-office-scripts-settings)
- [Office スクリプトの効果を元に戻す](../testing/undo.md)
