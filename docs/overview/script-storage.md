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
# <a name="office-scripts-file-storage-and-ownership"></a><span data-ttu-id="79a51-103">Office スクリプト ファイルのストレージと所有権</span><span class="sxs-lookup"><span data-stu-id="79a51-103">Office Scripts file storage and ownership</span></span>

<span data-ttu-id="79a51-104">Officeスクリプトは **、Microsoft OneDrive に .osts** ファイルとして保存されます。</span><span class="sxs-lookup"><span data-stu-id="79a51-104">Office Scripts are stored as **.osts** files in your Microsoft OneDrive.</span></span> <span data-ttu-id="79a51-105">これにより、スクリプトを特定のブックの外部に存在できます。</span><span class="sxs-lookup"><span data-stu-id="79a51-105">This allows your scripts to exist outside any particular workbook.</span></span> <span data-ttu-id="79a51-106">OneDrive 設定は、すべてのスクリプト **.osts** ファイルの共有アクセスとアクセス許可を制御します。Excel の設定とは独立しています。</span><span class="sxs-lookup"><span data-stu-id="79a51-106">Your OneDrive settings control the shared access and permissions for all script **.osts** files; independent of any Excel settings.</span></span>

## <a name="file-storage"></a><span data-ttu-id="79a51-107">ファイルの記憶域</span><span class="sxs-lookup"><span data-stu-id="79a51-107">File storage</span></span>

<span data-ttu-id="79a51-108">スクリプトOffice OneDrive に保存されます。</span><span class="sxs-lookup"><span data-stu-id="79a51-108">You Office Scripts are stored in your OneDrive.</span></span> <span data-ttu-id="79a51-109">**.osts ファイル** は **、/Documents/Officeフォルダーにあります**。</span><span class="sxs-lookup"><span data-stu-id="79a51-109">The **.osts** files are found in the **/Documents/Office Scripts/** folder.</span></span> <span data-ttu-id="79a51-110">ファイルの名前の変更や削除など、これらの **.osts** ファイルに対して行われた編集は、コード エディターとスクリプト ギャラリーに反映されます。</span><span class="sxs-lookup"><span data-stu-id="79a51-110">Any edits made to these **.osts** files, such as renaming or deleting files, will be reflected in the Code Editor and Script Gallery.</span></span>

<span data-ttu-id="79a51-111">ブックの 1 つと共有されているスクリプトは、スクリプト作成者の OneDrive に残ります。</span><span class="sxs-lookup"><span data-stu-id="79a51-111">Scripts that are shared with one of your workbooks remain in the script creator's OneDrive.</span></span> <span data-ttu-id="79a51-112">Excel で共有スクリプトを実行すると、ローカル フォルダーまたは OneDrive フォルダーにはコピーされません。</span><span class="sxs-lookup"><span data-stu-id="79a51-112">They are not copied to any of your local or OneDrive folders when you run the shared script in Excel.</span></span> <span data-ttu-id="79a51-113">コード **エディターの [コピーの** 作成] ボタンは、OneDrive にスクリプトの個別のコピーを保存します。</span><span class="sxs-lookup"><span data-stu-id="79a51-113">The **Make a Copy** button of the Code Editor saves a separate copy of the script in your OneDrive.</span></span> <span data-ttu-id="79a51-114">コピーに対する変更は、元のスクリプトには影響を与えかねない。</span><span class="sxs-lookup"><span data-stu-id="79a51-114">Changes to the copy don't affect the original script.</span></span>

### <a name="script-folders"></a><span data-ttu-id="79a51-115">スクリプト フォルダー</span><span class="sxs-lookup"><span data-stu-id="79a51-115">Script folders</span></span>

<span data-ttu-id="79a51-116">OneDrive にフォルダーを追加すると、スクリプトを整理し続けるのに役立ちます。</span><span class="sxs-lookup"><span data-stu-id="79a51-116">Adding folders to your OneDrive helps keep your scripts organized.</span></span> <span data-ttu-id="79a51-117">**/Documents/Office スクリプト/ の下の** フォルダーは、コード エディターの **[マイ スクリプト**] セクションに表示されます。</span><span class="sxs-lookup"><span data-stu-id="79a51-117">Any folders under **/Documents/Office Scripts/** are displayed under the **My Scripts** section of the Code Editor.</span></span> <span data-ttu-id="79a51-118">これらのフォルダーは、コード エディターを使用して作成または削除することはできません。</span><span class="sxs-lookup"><span data-stu-id="79a51-118">Please note that these folders cannot be created or deleted by using the Code Editor.</span></span> <span data-ttu-id="79a51-119">同様に、スクリプトをフォルダーに配置したり、コード エディターを使用してフォルダー間で移動したりすることはできません。</span><span class="sxs-lookup"><span data-stu-id="79a51-119">Likewise, scripts cannot be placed in folders, or moved across folders by using the Code Editor.</span></span>

:::image type="content" source="../images/script-folders.png" alt-text="作業ウィンドウに表示されるフォルダーに含まれるスクリプトを表示するコード エディターの [新しいスクリプト] ダイアログ。":::

## <a name="file-ownership-and-retention"></a><span data-ttu-id="79a51-121">ファイルの所有権と保持</span><span class="sxs-lookup"><span data-stu-id="79a51-121">File ownership and retention</span></span>

<span data-ttu-id="79a51-122">Officeスクリプトは、ユーザーの OneDrive に格納されます。</span><span class="sxs-lookup"><span data-stu-id="79a51-122">Office Scripts are stored in a user's OneDrive.</span></span> <span data-ttu-id="79a51-123">Microsoft OneDrive で指定された保持ポリシーと削除ポリシーに従います。</span><span class="sxs-lookup"><span data-stu-id="79a51-123">They follow the retention and deletion policies specified by Microsoft OneDrive.</span></span> <span data-ttu-id="79a51-124">組織から削除されるユーザーによって作成および共有されたスクリプトを処理する方法については、[OneDrive の保持と削除](/onedrive/retention-and-deletion)を参照してください。</span><span class="sxs-lookup"><span data-stu-id="79a51-124">To learn how to handle scripts that were created and shared by a user being removed from your organization, see [OneDrive retention and deletion](/onedrive/retention-and-deletion).</span></span>

## <a name="see-also"></a><span data-ttu-id="79a51-125">関連項目</span><span class="sxs-lookup"><span data-stu-id="79a51-125">See also</span></span>

- [<span data-ttu-id="79a51-126">Excel on the web での Office スクリプトの共有</span><span class="sxs-lookup"><span data-stu-id="79a51-126">Sharing Office Scripts in Excel for the Web</span></span>](https://support.microsoft.com/office/sharing-office-scripts-in-excel-for-the-web-226eddbc-3a44-4540-acfe-fccda3d1122b)
- [<span data-ttu-id="79a51-127">Office スクリプトのトラブルシューティング</span><span class="sxs-lookup"><span data-stu-id="79a51-127">Troubleshooting Office Scripts</span></span>](../testing/troubleshooting.md)
- [<span data-ttu-id="79a51-128">M365 での Office スクリプトの設定</span><span class="sxs-lookup"><span data-stu-id="79a51-128">Office Scripts settings in M365</span></span>](https://support.office.com/article/office-scripts-settings-in-m365-19d3c51a-6ca2-40ab-978d-60fa49554dcf)
- [<span data-ttu-id="79a51-129">Office スクリプトの効果を元に戻す</span><span class="sxs-lookup"><span data-stu-id="79a51-129">Undo the effects of an Office Script</span></span>](../testing/undo.md)
