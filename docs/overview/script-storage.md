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
# <a name="office-scripts-file-storage-and-ownership"></a><span data-ttu-id="c0b6a-103">Office スクリプトファイルの保存と所有権</span><span class="sxs-lookup"><span data-stu-id="c0b6a-103">Office Scripts file storage and ownership</span></span>

<span data-ttu-id="c0b6a-104">Office スクリプトは、Microsoft OneDrive に **ost** ファイルとして保存されます。</span><span class="sxs-lookup"><span data-stu-id="c0b6a-104">Office Scripts are stored as **.osts** files in your Microsoft OneDrive.</span></span> <span data-ttu-id="c0b6a-105">これにより、スクリプトは特定のブックの外に存在することができます。</span><span class="sxs-lookup"><span data-stu-id="c0b6a-105">This allows your scripts to exist outside any particular workbook.</span></span> <span data-ttu-id="c0b6a-106">OneDrive の設定は、すべてのスクリプトの **ost** ファイルの共有アクセスとアクセス許可を制御します。Excel のすべての設定に依存しません。</span><span class="sxs-lookup"><span data-stu-id="c0b6a-106">Your OneDrive settings control the shared access and permissions for all script **.osts** files; independent of any Excel settings.</span></span>

## <a name="file-storage"></a><span data-ttu-id="c0b6a-107">ファイルの記憶域</span><span class="sxs-lookup"><span data-stu-id="c0b6a-107">File storage</span></span>

<span data-ttu-id="c0b6a-108">Office スクリプトは OneDrive に保存されています。</span><span class="sxs-lookup"><span data-stu-id="c0b6a-108">You Office Scripts are stored in your OneDrive.</span></span> <span data-ttu-id="c0b6a-109">この **ost** ファイルは、/ **ドキュメント/Office スクリプト/** フォルダーにあります。</span><span class="sxs-lookup"><span data-stu-id="c0b6a-109">The **.osts** files are found in the **/Documents/Office Scripts/** folder.</span></span> <span data-ttu-id="c0b6a-110">ファイル名の変更や削除など、これらの **ost** ファイルに対して行われた編集は、コードエディターとスクリプトギャラリーに反映されます。</span><span class="sxs-lookup"><span data-stu-id="c0b6a-110">Any edits made to these **.osts** files, such as renaming or deleting files, will be reflected in the Code Editor and Script Gallery.</span></span>

<span data-ttu-id="c0b6a-111">ブックの1つと共有されているスクリプトは、スクリプト作成者の OneDrive に残ります。</span><span class="sxs-lookup"><span data-stu-id="c0b6a-111">Scripts that are shared with one of your workbooks remain in the script creator's OneDrive.</span></span> <span data-ttu-id="c0b6a-112">これらのフォルダーは、Excel で共有スクリプトを実行しても、ローカルフォルダーや OneDrive フォルダーにはコピーされません。</span><span class="sxs-lookup"><span data-stu-id="c0b6a-112">They are not copied to any of your local or OneDrive folders when you run the shared script in Excel.</span></span> <span data-ttu-id="c0b6a-113">コードエディターの [ **コピーの作成** ] ボタンをクリックすると、スクリプトの別のコピーが OneDrive に保存されます。</span><span class="sxs-lookup"><span data-stu-id="c0b6a-113">The **Make a Copy** button of the Code Editor saves a separate copy of the script in your OneDrive.</span></span> <span data-ttu-id="c0b6a-114">コピーを変更しても、元のスクリプトには影響しません。</span><span class="sxs-lookup"><span data-stu-id="c0b6a-114">Changes to the copy don't affect the original script.</span></span>

### <a name="script-folders"></a><span data-ttu-id="c0b6a-115">スクリプトフォルダー</span><span class="sxs-lookup"><span data-stu-id="c0b6a-115">Script folders</span></span>

<span data-ttu-id="c0b6a-116">OneDrive にフォルダーを追加すると、スクリプトを整理できます。</span><span class="sxs-lookup"><span data-stu-id="c0b6a-116">Adding folders to your OneDrive helps keep your scripts organized.</span></span> <span data-ttu-id="c0b6a-117">/ **ドキュメント/Office スクリプト/** の下のフォルダーは、コードエディターの [ **マイスクリプト** ] セクションの下に表示されます。</span><span class="sxs-lookup"><span data-stu-id="c0b6a-117">Any folders under **/Documents/Office Scripts/** are displayed under the **My Scripts** section of the Code Editor.</span></span> <span data-ttu-id="c0b6a-118">これらのフォルダーは、コードエディターを使用して作成または削除できないことに注意してください。</span><span class="sxs-lookup"><span data-stu-id="c0b6a-118">Please note that these folders cannot be created or deleted by using the Code Editor.</span></span> <span data-ttu-id="c0b6a-119">同様に、スクリプトはフォルダーに配置したり、コードエディターを使用してフォルダー間で移動したりすることはできません。</span><span class="sxs-lookup"><span data-stu-id="c0b6a-119">Likewise, scripts cannot be placed in folders, or moved across folders by using the Code Editor.</span></span>

![[コードエディター] 作業ウィンドウに表示されているフォルダー内の一部のスクリプト](../images/script-folders.png)

## <a name="file-ownership-and-retention"></a><span data-ttu-id="c0b6a-121">ファイルの所有権と保持</span><span class="sxs-lookup"><span data-stu-id="c0b6a-121">File ownership and retention</span></span>

<span data-ttu-id="c0b6a-122">Office スクリプトは、ユーザーの OneDrive に保存されます。</span><span class="sxs-lookup"><span data-stu-id="c0b6a-122">Office Scripts are stored in a user's OneDrive.</span></span> <span data-ttu-id="c0b6a-123">これらは、Microsoft OneDrive で指定されているアイテム保持ポリシーと削除ポリシーに従います。</span><span class="sxs-lookup"><span data-stu-id="c0b6a-123">They follow the retention and deletion policies specified by Microsoft OneDrive.</span></span> <span data-ttu-id="c0b6a-124">組織から削除されるユーザーによって作成および共有されたスクリプトを処理する方法については、[OneDrive の保持と削除](/onedrive/retention-and-deletion)を参照してください。</span><span class="sxs-lookup"><span data-stu-id="c0b6a-124">To learn how to handle scripts that were created and shared by a user being removed from your organization, see [OneDrive retention and deletion](/onedrive/retention-and-deletion).</span></span>

## <a name="see-also"></a><span data-ttu-id="c0b6a-125">関連項目</span><span class="sxs-lookup"><span data-stu-id="c0b6a-125">See also</span></span>

- [<span data-ttu-id="c0b6a-126">Excel on the web での Office スクリプトの共有</span><span class="sxs-lookup"><span data-stu-id="c0b6a-126">Sharing Office Scripts in Excel for the Web</span></span>](https://support.microsoft.com/office/sharing-office-scripts-in-excel-for-the-web-226eddbc-3a44-4540-acfe-fccda3d1122b)
- [<span data-ttu-id="c0b6a-127">Office スクリプトのトラブルシューティング</span><span class="sxs-lookup"><span data-stu-id="c0b6a-127">Troubleshooting Office Scripts</span></span>](../testing/troubleshooting.md)
- [<span data-ttu-id="c0b6a-128">M365 での Office スクリプトの設定</span><span class="sxs-lookup"><span data-stu-id="c0b6a-128">Office Scripts settings in M365</span></span>](https://support.office.com/article/office-scripts-settings-in-m365-19d3c51a-6ca2-40ab-978d-60fa49554dcf)
- [<span data-ttu-id="c0b6a-129">Office スクリプトの効果を元に戻す</span><span class="sxs-lookup"><span data-stu-id="c0b6a-129">Undo the effects of an Office Script</span></span>](../testing/undo.md)
