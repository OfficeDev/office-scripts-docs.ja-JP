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
# <a name="office-scripts-file-storage-and-ownership"></a><span data-ttu-id="ae903-103">Officeスクリプト ファイルのストレージと所有権</span><span class="sxs-lookup"><span data-stu-id="ae903-103">Office Scripts file storage and ownership</span></span>

<span data-ttu-id="ae903-104">Officeスクリプトは、ユーザーの **ファイルに .osts** ファイルとしてMicrosoft OneDrive。</span><span class="sxs-lookup"><span data-stu-id="ae903-104">Office Scripts are stored as **.osts** files in your Microsoft OneDrive.</span></span> <span data-ttu-id="ae903-105">これにより、スクリプトを特定のブックの外部に存在できます。</span><span class="sxs-lookup"><span data-stu-id="ae903-105">This allows your scripts to exist outside any particular workbook.</span></span> <span data-ttu-id="ae903-106">ユーザー OneDrive設定は、すべてのスクリプト **.osts** ファイルの共有アクセスとアクセス許可を制御します。任意の設定にExcelします。</span><span class="sxs-lookup"><span data-stu-id="ae903-106">Your OneDrive settings control the shared access and permissions for all script **.osts** files; independent of any Excel settings.</span></span>

## <a name="file-storage"></a><span data-ttu-id="ae903-107">ファイルの記憶域</span><span class="sxs-lookup"><span data-stu-id="ae903-107">File storage</span></span>

<span data-ttu-id="ae903-108">スクリプトOfficeは、ユーザーのサーバーにOneDrive。</span><span class="sxs-lookup"><span data-stu-id="ae903-108">You Office Scripts are stored in your OneDrive.</span></span> <span data-ttu-id="ae903-109">**.osts ファイル** は **、/Documents/Officeフォルダーにあります**。</span><span class="sxs-lookup"><span data-stu-id="ae903-109">The **.osts** files are found in the **/Documents/Office Scripts/** folder.</span></span> <span data-ttu-id="ae903-110">ファイルの名前の変更や削除など、これらの **.osts** ファイルに対して行われた編集は、コード エディターとスクリプト ギャラリーに反映されます。</span><span class="sxs-lookup"><span data-stu-id="ae903-110">Any edits made to these **.osts** files, such as renaming or deleting files, will be reflected in the Code Editor and Script Gallery.</span></span>

<span data-ttu-id="ae903-111">ブックの 1 つと共有されているスクリプトは、スクリプト作成者のデータベースに残OneDrive。</span><span class="sxs-lookup"><span data-stu-id="ae903-111">Scripts that are shared with one of your workbooks remain in the script creator's OneDrive.</span></span> <span data-ttu-id="ae903-112">共有スクリプトを OneDrive で実行すると、ローカル フォルダーまたはローカル フォルダーにはコピー Excel。</span><span class="sxs-lookup"><span data-stu-id="ae903-112">They are not copied to any of your local or OneDrive folders when you run the shared script in Excel.</span></span> <span data-ttu-id="ae903-113">コード **エディターの [コピー** を作成] ボタンをクリックすると、スクリプトの別のコピーがユーザーのページにOneDrive。</span><span class="sxs-lookup"><span data-stu-id="ae903-113">The **Make a Copy** button of the Code Editor saves a separate copy of the script in your OneDrive.</span></span> <span data-ttu-id="ae903-114">コピーに対する変更は、元のスクリプトには影響を与えかねない。</span><span class="sxs-lookup"><span data-stu-id="ae903-114">Changes to the copy don't affect the original script.</span></span>

### <a name="script-folders"></a><span data-ttu-id="ae903-115">スクリプト フォルダー</span><span class="sxs-lookup"><span data-stu-id="ae903-115">Script folders</span></span>

<span data-ttu-id="ae903-116">フォルダーをフォルダーに追加OneDriveスクリプトを整理するのに役立ちます。</span><span class="sxs-lookup"><span data-stu-id="ae903-116">Adding folders to your OneDrive helps keep your scripts organized.</span></span> <span data-ttu-id="ae903-117">**/Documents/Office スクリプト/ の** 下のフォルダーは、コード エディターの **[マイ スクリプト**] セクションに表示されます。</span><span class="sxs-lookup"><span data-stu-id="ae903-117">Any folders under **/Documents/Office Scripts/** are displayed under the **My Scripts** section of the Code Editor.</span></span> <span data-ttu-id="ae903-118">これらのフォルダーは、コード エディターを使用して作成または削除することはできません。</span><span class="sxs-lookup"><span data-stu-id="ae903-118">Please note that these folders cannot be created or deleted by using the Code Editor.</span></span> <span data-ttu-id="ae903-119">同様に、スクリプトをフォルダーに配置したり、コード エディターを使用してフォルダー間で移動したりすることはできません。</span><span class="sxs-lookup"><span data-stu-id="ae903-119">Likewise, scripts cannot be placed in folders, or moved across folders by using the Code Editor.</span></span>

:::image type="content" source="../images/script-folders.png" alt-text="作業ウィンドウに表示されるフォルダーに含まれるスクリプトを表示するコード エディターの [新しいスクリプト] ダイアログ":::

## <a name="file-ownership-and-retention"></a><span data-ttu-id="ae903-121">ファイルの所有権と保持</span><span class="sxs-lookup"><span data-stu-id="ae903-121">File ownership and retention</span></span>

<span data-ttu-id="ae903-122">Officeスクリプトは、ユーザーのデータベースにOneDrive。</span><span class="sxs-lookup"><span data-stu-id="ae903-122">Office Scripts are stored in a user's OneDrive.</span></span> <span data-ttu-id="ae903-123">ユーザーは、ユーザーが指定した保持ポリシーと削除ポリシー Microsoft OneDrive。</span><span class="sxs-lookup"><span data-stu-id="ae903-123">They follow the retention and deletion policies specified by Microsoft OneDrive.</span></span> <span data-ttu-id="ae903-124">組織から削除されるユーザーによって作成および共有されたスクリプトを処理する方法については、[OneDrive の保持と削除](/onedrive/retention-and-deletion)を参照してください。</span><span class="sxs-lookup"><span data-stu-id="ae903-124">To learn how to handle scripts that were created and shared by a user being removed from your organization, see [OneDrive retention and deletion](/onedrive/retention-and-deletion).</span></span>

## <a name="see-also"></a><span data-ttu-id="ae903-125">こちらもご覧ください</span><span class="sxs-lookup"><span data-stu-id="ae903-125">See also</span></span>

- [<span data-ttu-id="ae903-126">Excel on the web での Office スクリプトの共有</span><span class="sxs-lookup"><span data-stu-id="ae903-126">Sharing Office Scripts in Excel for the Web</span></span>](https://support.microsoft.com/office/sharing-office-scripts-in-excel-for-the-web-226eddbc-3a44-4540-acfe-fccda3d1122b)
- [<span data-ttu-id="ae903-127">Office スクリプトのトラブルシューティング</span><span class="sxs-lookup"><span data-stu-id="ae903-127">Troubleshooting Office Scripts</span></span>](../testing/troubleshooting.md)
- [<span data-ttu-id="ae903-128">M365 での Office スクリプトの設定</span><span class="sxs-lookup"><span data-stu-id="ae903-128">Office Scripts settings in M365</span></span>](https://support.office.com/article/office-scripts-settings-in-m365-19d3c51a-6ca2-40ab-978d-60fa49554dcf)
- [<span data-ttu-id="ae903-129">Office スクリプトの効果を元に戻す</span><span class="sxs-lookup"><span data-stu-id="ae903-129">Undo the effects of an Office Script</span></span>](../testing/undo.md)
