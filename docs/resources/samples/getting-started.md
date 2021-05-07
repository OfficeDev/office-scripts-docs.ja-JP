---
title: スクリプトの使用Officeする
description: アクセス、環境Officeスクリプト パターンを含むスクリプトの概要。
ms.date: 04/01/2021
localization_priority: Normal
ROBOTS: NOINDEX
ms.openlocfilehash: d30c4fb4523c49b559e057eede4d5de162b74f9c
ms.sourcegitcommit: f7a7aebfb687f2a35dbed07ed62ff352a114525a
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 05/06/2021
ms.locfileid: "52232761"
---
# <a name="getting-started"></a><span data-ttu-id="19c3e-103">はじめに</span><span class="sxs-lookup"><span data-stu-id="19c3e-103">Getting started</span></span>

<span data-ttu-id="19c3e-104">このセクションでは、アクセス、環境、スクリプトの基本、およびいくつかの基本的なスクリプト パターンOfficeスクリプトの基本について説明します。</span><span class="sxs-lookup"><span data-stu-id="19c3e-104">This section provides details about the basics of Office Scripts including access, environment, script fundamentals, and few basic script patterns.</span></span>

## <a name="environment-setup"></a><span data-ttu-id="19c3e-105">環境のセットアップ</span><span class="sxs-lookup"><span data-stu-id="19c3e-105">Environment setup</span></span>

<span data-ttu-id="19c3e-106">アクセス、環境、スクリプト エディターの基本について説明します。</span><span class="sxs-lookup"><span data-stu-id="19c3e-106">Learn about the basics of access, environment, and script editor.</span></span>

<span data-ttu-id="19c3e-107">[![スクリプト アプリケーションOfficeの基本](../../images/getting-started-env.png)](https://youtu.be/vvCtxsjPxo8 "スクリプト アプリケーションOfficeの基本")</span><span class="sxs-lookup"><span data-stu-id="19c3e-107">[![Basics of Office Scripts application](../../images/getting-started-env.png)](https://youtu.be/vvCtxsjPxo8 "Basics of Office Scripts application")</span></span>

### <a name="access"></a><span data-ttu-id="19c3e-108">Access</span><span class="sxs-lookup"><span data-stu-id="19c3e-108">Access</span></span>

<span data-ttu-id="19c3e-109">Officeスクリプトでは、[スクリプト] の [組織の設定] **Microsoft 365の下** 設定管理者がOffice  >    >  **する必要があります**。</span><span class="sxs-lookup"><span data-stu-id="19c3e-109">Office Scripts requires admin settings available for Microsoft 365 administrator under **Settings** > **Org settings** > **Office Scripts**.</span></span> <span data-ttu-id="19c3e-110">既定では、すべてのユーザーに対して有効になっています。</span><span class="sxs-lookup"><span data-stu-id="19c3e-110">By default, it's turned on for all users.</span></span> <span data-ttu-id="19c3e-111">2 つのサブ設定があります。管理者はオンとオフを切り替えます。</span><span class="sxs-lookup"><span data-stu-id="19c3e-111">There are two sub-settings, which the admin can turn on and off.</span></span>

* <span data-ttu-id="19c3e-112">組織内でスクリプトを共有する機能</span><span class="sxs-lookup"><span data-stu-id="19c3e-112">Ability to share scripts within the organization</span></span>
* <span data-ttu-id="19c3e-113">スクリプトを使用する機能は、Power Automate</span><span class="sxs-lookup"><span data-stu-id="19c3e-113">Ability to use scripts in Power Automate</span></span>

<span data-ttu-id="19c3e-114">Office スクリプトにアクセスできる場合は、Excel on the web (ブラウザー) でファイルを開き、[自動化] タブが [Excel] リボンに表示されるのかを確認します。 </span><span class="sxs-lookup"><span data-stu-id="19c3e-114">You can tell if you have access to Office Scripts by opening a file in Excel on the web (browser) and seeing if the **Automate** tab appears in the Excel ribbon or not.</span></span>
<span data-ttu-id="19c3e-115">[自動化] タブが表示できない \*\*場合は、[\*\* このトラブルシューティング [] セクションを確認してください](../../testing/troubleshooting.md#automate-tab-not-appearing-or-office-scripts-unavailable)。</span><span class="sxs-lookup"><span data-stu-id="19c3e-115">If you still can't see the **Automate** tab, check [this troubleshooting section](../../testing/troubleshooting.md#automate-tab-not-appearing-or-office-scripts-unavailable).</span></span>

### <a name="availability"></a><span data-ttu-id="19c3e-116">可用性</span><span class="sxs-lookup"><span data-stu-id="19c3e-116">Availability</span></span>

<span data-ttu-id="19c3e-117">Officeスクリプトは、E3+ ライセンスExcel on the webのEnterpriseでのみ使用できます (コンシューマー アカウントと E1 アカウントはサポートされていません)。</span><span class="sxs-lookup"><span data-stu-id="19c3e-117">Office Scripts is available only in the Excel on the web for Enterprise E3+ licenses (Consumer and E1 accounts are not supported).</span></span> <span data-ttu-id="19c3e-118">Officeスクリプトは、Excel Mac Windowsサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="19c3e-118">Office Scripts is not yet supported in Excel on Windows and Mac.</span></span>

### <a name="scripts-and-editor"></a><span data-ttu-id="19c3e-119">スクリプトとエディター</span><span class="sxs-lookup"><span data-stu-id="19c3e-119">Scripts and editor</span></span>

<span data-ttu-id="19c3e-120">コード エディターは、Excel on the web (オンライン バージョン) に組み込ばれています。</span><span class="sxs-lookup"><span data-stu-id="19c3e-120">The code editor is built right into Excel on the web (online version).</span></span> <span data-ttu-id="19c3e-121">[編集] や [サブVisual Studio Code] のようなエディターを使用した場合、この編集エクスペリエンスは非常に似ています。</span><span class="sxs-lookup"><span data-stu-id="19c3e-121">If you have used editors like Visual Studio Code or Sublime, this editing experience will be quite similar.</span></span>
<span data-ttu-id="19c3e-122">このエディターで使用Visual Studio Codeのショートカット キーの大部分は、Office編集エクスペリエンスでも機能します。</span><span class="sxs-lookup"><span data-stu-id="19c3e-122">Most of the shortcut keys that Visual Studio Code editor uses work in the Office Scripts editing experience as well.</span></span> <span data-ttu-id="19c3e-123">次のショートカット キーの資料をご覧ください。</span><span class="sxs-lookup"><span data-stu-id="19c3e-123">Check out the following shortcut keys handouts.</span></span>

* [<span data-ttu-id="19c3e-124">macOS</span><span class="sxs-lookup"><span data-stu-id="19c3e-124">macOS</span></span>](https://code.visualstudio.com/shortcuts/keyboard-shortcuts-macos.pdf)
* [<span data-ttu-id="19c3e-125">Windows</span><span class="sxs-lookup"><span data-stu-id="19c3e-125">Windows</span></span>](https://code.visualstudio.com/shortcuts/keyboard-shortcuts-windows.pdf)

#### <a name="key-things-to-note"></a><span data-ttu-id="19c3e-126">重要な注意点</span><span class="sxs-lookup"><span data-stu-id="19c3e-126">Key things to note</span></span>

* <span data-ttu-id="19c3e-127">Officeスクリプトは、サイト、サイト、およびチーム OneDrive for BusinessにSharePointファイルでのみ使用できます。</span><span class="sxs-lookup"><span data-stu-id="19c3e-127">Office Scripts is only available for files stored in OneDrive for Business, SharePoint sites, and Team sites.</span></span>
* <span data-ttu-id="19c3e-128">エディターには、スクリプトの拡張機能は表示されます。</span><span class="sxs-lookup"><span data-stu-id="19c3e-128">The editor doesn't show the script's extension.</span></span> <span data-ttu-id="19c3e-129">実際には、これらは TypeScript ファイルですが、カスタム拡張機能と一緒に格納されます `.osts` 。</span><span class="sxs-lookup"><span data-stu-id="19c3e-129">In reality, these are TypeScript files but they are stored with a custom extension called `.osts`.</span></span>
* <span data-ttu-id="19c3e-130">スクリプトは、ユーザー独自のフォルダーにOneDrive for Businessされます `My Files/Documents/OfficeScripts` 。</span><span class="sxs-lookup"><span data-stu-id="19c3e-130">The scripts are stored in your own OneDrive for Business folder `My Files/Documents/OfficeScripts`.</span></span> <span data-ttu-id="19c3e-131">このフォルダーを管理する必要はもうない。</span><span class="sxs-lookup"><span data-stu-id="19c3e-131">You won't need to manage this folder.</span></span> <span data-ttu-id="19c3e-132">エディターが表示/編集エクスペリエンスを管理する場合は、この側面を無視できます。</span><span class="sxs-lookup"><span data-stu-id="19c3e-132">For your part, you can ignore this aspect as the editor manages the viewing/editing experience.</span></span>
* <span data-ttu-id="19c3e-133">スクリプトは、ファイルの一部としてExcelされません。</span><span class="sxs-lookup"><span data-stu-id="19c3e-133">Scripts are not stored as part of Excel files.</span></span> <span data-ttu-id="19c3e-134">これらは個別に格納されます。</span><span class="sxs-lookup"><span data-stu-id="19c3e-134">They are stored separately.</span></span>
* <span data-ttu-id="19c3e-135">スクリプトをファイルファイルと共有Excel、実際にはスクリプトをファイルにリンクし、添付しないという意味です。</span><span class="sxs-lookup"><span data-stu-id="19c3e-135">You can share the script with an Excel file which in effect means you are linking the script with the file, not attaching it.</span></span> <span data-ttu-id="19c3e-136">Excel ファイルにアクセスできるユーザーは、スクリプトの表示、実行、またはコピーの作成を行えます。</span><span class="sxs-lookup"><span data-stu-id="19c3e-136">Whoever has access to the Excel file will also be able to **view**, **run**, or **make a copy** of the script.</span></span> <span data-ttu-id="19c3e-137">これは VBA マクロと比較して重要な違いです。</span><span class="sxs-lookup"><span data-stu-id="19c3e-137">This is a key difference compared to VBA macros.</span></span>
* <span data-ttu-id="19c3e-138">スクリプトを共有しない限り、他の誰も自分のライブラリに存在するスクリプトにアクセスできません。</span><span class="sxs-lookup"><span data-stu-id="19c3e-138">Unless you share your scripts, no one else can access it as it resides in your own library.</span></span>
* <span data-ttu-id="19c3e-139">スクリプトは、ローカル ディスクまたはカスタム クラウドの場所からリンクできません。</span><span class="sxs-lookup"><span data-stu-id="19c3e-139">Scripts can't be linked from a local disk or custom cloud locations.</span></span> <span data-ttu-id="19c3e-140">Officeスクリプトは、事前に定義された場所 (上記の OneDrive フォルダー) または共有スクリプト上にあるスクリプトのみを認識して実行します。</span><span class="sxs-lookup"><span data-stu-id="19c3e-140">Office Scripts only recognizes and runs a script that is on predefined location (your OneDrive folder mentioned above) or shared scripts.</span></span>
* <span data-ttu-id="19c3e-141">編集中、ファイルはブラウザーに一時的に保存されますが、Excel ウィンドウを閉じる前にスクリプトを保存して、OneDrive場所に保存する必要があります。</span><span class="sxs-lookup"><span data-stu-id="19c3e-141">During editing, files are temporarily saved in the browser but you'll have to save the script before closing the Excel window to save it to the OneDrive location.</span></span> <span data-ttu-id="19c3e-142">編集後にファイルを保存することを忘れないでください。</span><span class="sxs-lookup"><span data-stu-id="19c3e-142">Don't forget to save the file after edits.</span></span>

## <a name="gentle-introduction-to-scripting"></a><span data-ttu-id="19c3e-143">スクリプトの優しい概要</span><span class="sxs-lookup"><span data-stu-id="19c3e-143">Gentle introduction to scripting</span></span>

<span data-ttu-id="19c3e-144">Officeスクリプトは、TypeScript 言語で記述されたスタンドアロン スクリプトで、選択したブックに対して何らかのオートメーションを実行するExcelです。</span><span class="sxs-lookup"><span data-stu-id="19c3e-144">Office Scripts are standalone scripts written in the TypeScript language that contain instructions to perform some automation against the selected Excel workbook.</span></span> <span data-ttu-id="19c3e-145">すべてのオートメーション命令はスクリプト内に自己格納され、スクリプトは他のスクリプトを呼び出したり呼び出したりできません。</span><span class="sxs-lookup"><span data-stu-id="19c3e-145">All automation instructions are self-contained within a script and scripts can't invoke or call other scripts.</span></span> <span data-ttu-id="19c3e-146">すべてのスクリプトはスタンドアロン ファイルに格納され、ユーザーのデータベース フォルダーにOneDriveされます。</span><span class="sxs-lookup"><span data-stu-id="19c3e-146">All scripts are stored in standalone files and stored on the user's OneDrive folder.</span></span> <span data-ttu-id="19c3e-147">新しいスクリプトを記録したり、記録されたスクリプトを編集したり、新しいスクリプトを最初から書き込むなど、すべて組み込みのエディター インターフェイス内で実行できます。</span><span class="sxs-lookup"><span data-stu-id="19c3e-147">You can record a new script, edit a recorded script, or write a whole new script from scratch, all within a built-in editor interface.</span></span> <span data-ttu-id="19c3e-148">スクリプトの最Officeは、ユーザーからのセットアップが不要な場合です。</span><span class="sxs-lookup"><span data-stu-id="19c3e-148">The best part of Office Scripts is that they don't need any further setup from users.</span></span> <span data-ttu-id="19c3e-149">外部ライブラリ、Web ページ、UI 要素、セットアップなどはありません。すべての環境セットアップは、Officeスクリプトによって処理され、簡単な API インターフェイスを介してオートメーションに簡単かつ迅速にアクセスできます。</span><span class="sxs-lookup"><span data-stu-id="19c3e-149">No external libraries, web pages, or UI elements, setup, etc. All the environment setup is handled by Office Scripts and it allows easy and fast access to automation through a simple API interface.</span></span>

<span data-ttu-id="19c3e-150">スクリプトを編集して移動する方法を理解するために役立つ基本的な概念には、次のようなものがあります。</span><span class="sxs-lookup"><span data-stu-id="19c3e-150">Some of the basic concepts helpful to understand how to edit and navigate around scripts include:</span></span>

* <span data-ttu-id="19c3e-151">基本的な TypeScript 言語の構文</span><span class="sxs-lookup"><span data-stu-id="19c3e-151">Basic TypeScript language syntax</span></span>
* <span data-ttu-id="19c3e-152">関数と `main` 引数の理解</span><span class="sxs-lookup"><span data-stu-id="19c3e-152">Understanding of `main` function and arguments</span></span>
* <span data-ttu-id="19c3e-153">オブジェクトと階層、メソッド、プロパティ</span><span class="sxs-lookup"><span data-stu-id="19c3e-153">Objects and hierarchy, methods, properties</span></span>
* <span data-ttu-id="19c3e-154">コレクション (配列): ナビゲーションと操作</span><span class="sxs-lookup"><span data-stu-id="19c3e-154">Collection (array): navigation and operations</span></span>
* <span data-ttu-id="19c3e-155">型の定義</span><span class="sxs-lookup"><span data-stu-id="19c3e-155">Type definitions</span></span>
* <span data-ttu-id="19c3e-156">環境: レコード/編集、実行、結果の確認、共有</span><span class="sxs-lookup"><span data-stu-id="19c3e-156">Environment: record/edit, run, examine results, share</span></span>

<span data-ttu-id="19c3e-157">このビデオとセクションでは、これらの概念について詳しく説明します。</span><span class="sxs-lookup"><span data-stu-id="19c3e-157">This video and section explain some of these concepts in detail.</span></span>

<span data-ttu-id="19c3e-158">[![スクリプトのOffice](../../images/getting-started-v_script.png)](https://youtu.be/8Zsrc1uaiiU "スクリプトの基本")</span><span class="sxs-lookup"><span data-stu-id="19c3e-158">[![Basics of Office Scripts](../../images/getting-started-v_script.png)](https://youtu.be/8Zsrc1uaiiU "Basics of Scripts")</span></span>

### <a name="language-typescript"></a><span data-ttu-id="19c3e-159">言語: TypeScript</span><span class="sxs-lookup"><span data-stu-id="19c3e-159">Language: TypeScript</span></span>

<span data-ttu-id="19c3e-160">[Officeスクリプト](../../index.md)は、静的型定義を追加して JavaScript (世界で最も使用されている言語の 1 つ) 上に構築するオープンソース言語である[TypeScript](https://www.typescriptlang.org/)言語を使用して記述されます。</span><span class="sxs-lookup"><span data-stu-id="19c3e-160">[Office Scripts](../../index.md) is written using the [TypeScript language](https://www.typescriptlang.org/), which is an open-source language that builds on JavaScript (one of the world's most used) by adding static type definitions.</span></span> <span data-ttu-id="19c3e-161">Web サイトが言うように、オブジェクトの図形を記述し、より良いドキュメントを提供し、TypeScript がコードが正しく動作することを検証する方法 `Types` を提供します。</span><span class="sxs-lookup"><span data-stu-id="19c3e-161">As the website says, `Types` provide a way to describe the shape of an object, providing better documentation, and allowing TypeScript to validate that your code is working correctly.</span></span>

<span data-ttu-id="19c3e-162">言語構文自体は、JavaScript を使用して記述され [、TypeScript](https://developer.mozilla.org/docs/Web/JavaScript) の規則を使用してスクリプトで定義された追加のタイピングが含まれます。</span><span class="sxs-lookup"><span data-stu-id="19c3e-162">The language syntax itself is written using [JavaScript](https://developer.mozilla.org/docs/Web/JavaScript) with additional typings defined in the script using TypeScript conventions.</span></span> <span data-ttu-id="19c3e-163">ほとんどの場合、JavaScript で記述Officeスクリプトを考える必要があります。</span><span class="sxs-lookup"><span data-stu-id="19c3e-163">For the most part, you can think of Office Scripts as written in JavaScript.</span></span> <span data-ttu-id="19c3e-164">スクリプトの使用を開始するには、JavaScript 言語の基本を理解Office必要です。オートメーションの旅を始めるには、熟練している必要はありません。</span><span class="sxs-lookup"><span data-stu-id="19c3e-164">It is essential that you understand the basics of JavaScript language to begin your Office Scripts journey; though you don't need to be proficient at it to begin your automation journey.</span></span> <span data-ttu-id="19c3e-165">Officeスクリプトのアクション レコーダーを使用すると、コードコメントが含まれているため、スクリプトステートメントを理解できます。</span><span class="sxs-lookup"><span data-stu-id="19c3e-165">With the Office Scripts' action recorder, you can understand the script statements because code comments are included and you can follow along and make small edits.</span></span>

<span data-ttu-id="19c3e-166">Officeスクリプト API は、スクリプトが Excel と対話できる機能で、コーディングの背景があまりないエンド ユーザー向けに設計されています。</span><span class="sxs-lookup"><span data-stu-id="19c3e-166">Office Scripts APIs, which allow the script to interact with Excel, are designed for end-users who may not have much coding background.</span></span> <span data-ttu-id="19c3e-167">API は同期的に呼び出すことができるので、約束やコールバックなどの高度なトピックを知る必要がなされません。</span><span class="sxs-lookup"><span data-stu-id="19c3e-167">APIs can be invoked synchronously and you don't need to know advanced topics such as promises or callbacks.</span></span> <span data-ttu-id="19c3e-168">Officeスクリプト API の設計では、次の機能が提供されます。</span><span class="sxs-lookup"><span data-stu-id="19c3e-168">Office Scripts API design provides:</span></span>

* <span data-ttu-id="19c3e-169">メソッド、getters/setters を持つ単純なオブジェクト モデル。</span><span class="sxs-lookup"><span data-stu-id="19c3e-169">Simple object model with methods, getters/setters.</span></span>
* <span data-ttu-id="19c3e-170">通常の配列として簡単にアクセスできるオブジェクト コレクション。</span><span class="sxs-lookup"><span data-stu-id="19c3e-170">Easy-to-access object collections as regular arrays.</span></span>
* <span data-ttu-id="19c3e-171">単純なエラー処理オプション。</span><span class="sxs-lookup"><span data-stu-id="19c3e-171">Simple error handling options.</span></span>
* <span data-ttu-id="19c3e-172">ユーザーが目の前のシナリオに集中するのを助ける、選択したシナリオのパフォーマンスを最適化しました。</span><span class="sxs-lookup"><span data-stu-id="19c3e-172">Optimized performance for select scenarios helping users to focus on the scenario at hand.</span></span>

### <a name="main-function-the-scripts-starting-point"></a><span data-ttu-id="19c3e-173">`main` 関数: スクリプトの開始点</span><span class="sxs-lookup"><span data-stu-id="19c3e-173">`main` function: The script's starting point</span></span>

<span data-ttu-id="19c3e-174">Officeスクリプトの実行は関数から始 `main` まります。</span><span class="sxs-lookup"><span data-stu-id="19c3e-174">Office Scripts' execution begins at the `main` function.</span></span> <span data-ttu-id="19c3e-175">スクリプトは、型、インターフェイス、変数などの宣言と共に 1 つ以上の関数を含む 1 つのファイルです。スクリプトに従う場合は、スクリプトを実行するときに常にExcel呼び出す関数として関数を `main` `main` 開始します。</span><span class="sxs-lookup"><span data-stu-id="19c3e-175">A script is a single file containing one or many functions along with declarations of types, interfaces, variables, etc. To follow along with the script, begin with the `main` function as Excel always first invokes the `main` function when you execute any script.</span></span> <span data-ttu-id="19c3e-176">関数には常に、スクリプトが実行されている現在のブックを識別する変数名である、という名前の引数 (またはパラメーター) が少なくとも `main` `workbook` 1 つ含まれます。</span><span class="sxs-lookup"><span data-stu-id="19c3e-176">The `main` function will always have at least one argument (or parameter) named `workbook`, which is just a variable name identifying the current workbook against which the script is running.</span></span> <span data-ttu-id="19c3e-177">(オフライン) の実行で使用する追加Power Automate定義できます。</span><span class="sxs-lookup"><span data-stu-id="19c3e-177">You can define additional arguments for usage with Power Automate (offline) execution.</span></span>

* `function main(workbook: ExcelScript.Workbook)`

<span data-ttu-id="19c3e-178">スクリプトを小さな関数に整理して、コードの再利用性、明快さなどを支援できます。その他の関数は、メイン関数の内側または外側にできますが、常に同じファイルに含まれています。</span><span class="sxs-lookup"><span data-stu-id="19c3e-178">A script can be organized into smaller functions to aid with code reusability, clarity, etc. Other functions can be inside or outside of the main function but always in the same file.</span></span> <span data-ttu-id="19c3e-179">スクリプトは自己格納型であり、同じファイルで定義されている関数のみを使用できます。</span><span class="sxs-lookup"><span data-stu-id="19c3e-179">A script is self-contained and can only use functions defined in the same file.</span></span> <span data-ttu-id="19c3e-180">スクリプトは、別のスクリプトを呼び出Officeできません。</span><span class="sxs-lookup"><span data-stu-id="19c3e-180">Scripts cannot invoke or call another Office Script.</span></span>

<span data-ttu-id="19c3e-181">したがって、要約すると次の作業が行います。</span><span class="sxs-lookup"><span data-stu-id="19c3e-181">So, in summary:</span></span>

* <span data-ttu-id="19c3e-182">関数 `main` は、任意のスクリプトのエントリ ポイントです。</span><span class="sxs-lookup"><span data-stu-id="19c3e-182">The `main` function is the entry point for any script.</span></span> <span data-ttu-id="19c3e-183">関数が実行されると、Excelアプリケーションは、ブックを最初のパラメーターとして指定して、このメイン関数を呼び出します。</span><span class="sxs-lookup"><span data-stu-id="19c3e-183">When the function is executed, the Excel application invokes this main function by providing the workbook as its first parameter.</span></span>
* <span data-ttu-id="19c3e-184">最初の引数とその型宣言は、表示 `workbook` された状態で保持することが重要です。</span><span class="sxs-lookup"><span data-stu-id="19c3e-184">It's important to keep the first argument `workbook` and its type declaration as it appears.</span></span> <span data-ttu-id="19c3e-185">関数に新しい引数を追加できますが (次のセクションを参照)、最初の引数は変更 `main` しません。</span><span class="sxs-lookup"><span data-stu-id="19c3e-185">You can add new arguments to the `main` function (see the next section) but do keep the first argument as is.</span></span>

:::image type="content" source="../../images/getting-started-main-introduction.png" alt-text="主な関数は、スクリプトのエントリ ポイントです。":::

#### <a name="send-or-receive-data-from-other-apps"></a><span data-ttu-id="19c3e-187">他のアプリからのデータの送受信</span><span class="sxs-lookup"><span data-stu-id="19c3e-187">Send or receive data from other apps</span></span>

<span data-ttu-id="19c3e-188">組織の他Excelにスクリプトを実行することで、組織の他[の部分に](https://flow.microsoft.com)Power Automate。</span><span class="sxs-lookup"><span data-stu-id="19c3e-188">You can connect Excel to other parts of your organization by running scripts in [Power Automate](https://flow.microsoft.com).</span></span> <span data-ttu-id="19c3e-189">詳細については、「スクリプト フロー[でのOfficeスクリプトの実行Power Automateします](../../develop/power-automate-integration.md)。</span><span class="sxs-lookup"><span data-stu-id="19c3e-189">Learn more about [running Office Scripts in Power Automate flows](../../develop/power-automate-integration.md).</span></span>

<span data-ttu-id="19c3e-190">データを受信または送信する方法は、Excelを介 `main` して行います。</span><span class="sxs-lookup"><span data-stu-id="19c3e-190">The way to receive or send data from and to Excel is through the `main` function.</span></span> <span data-ttu-id="19c3e-191">これは、受信データと送信データをスクリプトで記述および使用できる情報ゲートウェイと考えてください。</span><span class="sxs-lookup"><span data-stu-id="19c3e-191">Think of it as the information gateway that allows incoming and outgoing data to be described and used in the script.</span></span> <span data-ttu-id="19c3e-192">データ型を使用して、スクリプトの外部からデータを受け取り、TypeScript で認識されるデータ (、など) を返したり、スクリプトで定義したインターフェイスの形式のオブジェクトを `string` `string` `number` `boolean` 返したりできます。</span><span class="sxs-lookup"><span data-stu-id="19c3e-192">You can receive data from outside the script using the `string` data type and return any TypeScript-recognized data such as `string`, `number`, `boolean`, or any objects in the form of interfaces you define in the script.</span></span>

:::image type="content" source="../../images/getting-started-data-in-out.png" alt-text="スクリプトの入力と出力":::

#### <a name="use-functions-to-organize-and-reuse-code"></a><span data-ttu-id="19c3e-194">関数を使用してコードを整理および再利用する</span><span class="sxs-lookup"><span data-stu-id="19c3e-194">Use functions to organize and reuse code</span></span>

<span data-ttu-id="19c3e-195">関数を使用して、スクリプト内でコードを整理および再利用できます。</span><span class="sxs-lookup"><span data-stu-id="19c3e-195">You can use functions to organize and reuse code within your script.</span></span>

:::image type="content" source="../../images/getting-started-use-functions.png" alt-text="スクリプトでの関数の使用":::

### <a name="objects-hierarchy-methods-properties-collections"></a><span data-ttu-id="19c3e-197">オブジェクト、階層、メソッド、プロパティ、コレクション</span><span class="sxs-lookup"><span data-stu-id="19c3e-197">Objects, hierarchy, methods, properties, collections</span></span>

<span data-ttu-id="19c3e-198">すべてのオブジェクトExcelオブジェクト モデルは、タイプのブック オブジェクトから始まるオブジェクトの階層構造で定義されます `ExcelScript.Workbook` 。</span><span class="sxs-lookup"><span data-stu-id="19c3e-198">All of Excel's object model is defined in a hierarchical structure of objects, beginning with the workbook object of type `ExcelScript.Workbook`.</span></span> <span data-ttu-id="19c3e-199">オブジェクトには、メソッド、プロパティ、そのオブジェクト内の他のオブジェクトを含めることができます。</span><span class="sxs-lookup"><span data-stu-id="19c3e-199">An object can contain methods, properties, and other objects within it.</span></span> <span data-ttu-id="19c3e-200">オブジェクトは、メソッドを使用して互いにリンクされます。</span><span class="sxs-lookup"><span data-stu-id="19c3e-200">Objects are linked to each other using the methods.</span></span> <span data-ttu-id="19c3e-201">オブジェクトのメソッドは、別のオブジェクトまたはオブジェクトのコレクションを返す場合があります。</span><span class="sxs-lookup"><span data-stu-id="19c3e-201">An object's method can return another object or collection of objects.</span></span> <span data-ttu-id="19c3e-202">コード エディターのコード エディターのIntelliSense (コード補完) 機能を使用すると、オブジェクト階層を探索できます。</span><span class="sxs-lookup"><span data-stu-id="19c3e-202">Using the code editor's IntelliSense (code completion) feature is a great way to explore the object hierarchy.</span></span> <span data-ttu-id="19c3e-203">公式のリファレンス ドキュメント サイト [を使用](/javascript/api/office-scripts/overview) して、オブジェクト間の関係をフォローすることもできます。</span><span class="sxs-lookup"><span data-stu-id="19c3e-203">You can also use the [official reference documentation site](/javascript/api/office-scripts/overview) to follow along with the relationships among objects.</span></span>

<span data-ttu-id="19c3e-204">オブジェクト [は](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Object) プロパティのコレクションであり、プロパティは名前 (またはキー) と値の関連付けです。</span><span class="sxs-lookup"><span data-stu-id="19c3e-204">An [object](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Object) is a collection of properties, and a property is an association between a name (or key) and a value.</span></span> <span data-ttu-id="19c3e-205">プロパティの値には関数を指定できます。その場合、プロパティはメソッドと呼ばれる場合があります。</span><span class="sxs-lookup"><span data-stu-id="19c3e-205">A property's value can be a function, in which case the property is known as a method.</span></span> <span data-ttu-id="19c3e-206">Office Scripts オブジェクト モデルの場合、オブジェクトは、グラフ、ハイパーリンク、ピボット テーブルなど、ユーザーが操作する Excel ファイル内の物を表します。また、ワークシートの保護属性などのオブジェクトの動作を表す場合があります。</span><span class="sxs-lookup"><span data-stu-id="19c3e-206">In the case of the Office Scripts object model, an object represents a thing in the Excel file that users interact with such as a chart, hyperlink, pivot-table, etc. It can also represent the behavior of an object such as the protection attributes of a worksheet.</span></span>

<span data-ttu-id="19c3e-207">TypeScript オブジェクトとプロパティとメソッドのトピックは非常に深いです。</span><span class="sxs-lookup"><span data-stu-id="19c3e-207">The topic of TypeScript objects and properties vs methods is quite deep.</span></span> <span data-ttu-id="19c3e-208">スクリプトを使い始めて生産性を高めるには、次の基本的なことを覚えておく必要があります。</span><span class="sxs-lookup"><span data-stu-id="19c3e-208">In order to get started with the script and be productive, you can remember a few basic things:</span></span>

* <span data-ttu-id="19c3e-209">オブジェクトとプロパティの両方にアクセスするには、(ドット) 表記を使用し、オブジェクトは左側に、プロパティまたはメソッドは `.` `.` 右側に表示されます。</span><span class="sxs-lookup"><span data-stu-id="19c3e-209">Both objects and properties are accessed using `.` (dot) notation, with the object on the left side of the `.` and the property or method on the right side.</span></span> <span data-ttu-id="19c3e-210">例: `hyperlink.address` , `range.getAddress()` .</span><span class="sxs-lookup"><span data-stu-id="19c3e-210">Examples: `hyperlink.address`, `range.getAddress()`.</span></span>
* <span data-ttu-id="19c3e-211">プロパティは、実際にはスカラー (文字列、ブール値、数値) です。</span><span class="sxs-lookup"><span data-stu-id="19c3e-211">Properties are scalar in nature (strings, booleans, numbers).</span></span> <span data-ttu-id="19c3e-212">たとえば、ブックの名前、ワークシートの位置、テーブルにフッターがあるかどうかの値を指定します。</span><span class="sxs-lookup"><span data-stu-id="19c3e-212">For example, name of a workbook, position of a worksheet, the value of whether the table has a footer or not.</span></span>
* <span data-ttu-id="19c3e-213">メソッドは、オープンクローズかっこを使用して "呼び出された" または "実行" されます。</span><span class="sxs-lookup"><span data-stu-id="19c3e-213">Methods are 'invoked' or 'executed' using the open-close parentheses.</span></span> <span data-ttu-id="19c3e-214">例: `table.delete()`。</span><span class="sxs-lookup"><span data-stu-id="19c3e-214">Example: `table.delete()`.</span></span> <span data-ttu-id="19c3e-215">場合によっては、開いているかっこの間に引数を含めて関数に渡される場合があります `range.setValue('Hello')` 。</span><span class="sxs-lookup"><span data-stu-id="19c3e-215">Sometimes an argument is passed to a function by including them between open-close parentheses: `range.setValue('Hello')`.</span></span> <span data-ttu-id="19c3e-216">多くの引数を関数に渡し (コントラクト/署名で定義)、それらを分離するには、 `,` を使用します。</span><span class="sxs-lookup"><span data-stu-id="19c3e-216">You can pass many arguments to a function (as defined by its contract/signature) and separate them using `,`.</span></span>  <span data-ttu-id="19c3e-217">例: `worksheet.addTable('A1:D6', true)`。</span><span class="sxs-lookup"><span data-stu-id="19c3e-217">For example: `worksheet.addTable('A1:D6', true)`.</span></span> <span data-ttu-id="19c3e-218">文字列、数値、ブール値、その他のオブジェクトなど、メソッドで必要に応じて任意の型の引数を渡す (たとえば、スクリプト内の他の場所に作成されたオブジェクト)。 `worksheet.addTable(targetRange, true)` `targetRange`</span><span class="sxs-lookup"><span data-stu-id="19c3e-218">You can pass arguments of any type as required by the method such as strings, number, boolean, or even other objects, for example, `worksheet.addTable(targetRange, true)`, where `targetRange` is an object created elsewhere in the script.</span></span>
* <span data-ttu-id="19c3e-219">メソッドは、スカラー プロパティ (名前、アドレスなど) や別のオブジェクト (範囲、グラフ) などのオブジェクトを返したり、何も返しません (メソッドの場合など `delete` )。</span><span class="sxs-lookup"><span data-stu-id="19c3e-219">Methods can return a thing such as a scalar property (name, address, etc.) or another object (range, chart), or not return anything at all (such as the case with `delete` methods).</span></span> <span data-ttu-id="19c3e-220">変数を宣言するか、既存の変数に割り当てると、メソッドが返す値を受け取る。</span><span class="sxs-lookup"><span data-stu-id="19c3e-220">You receive what the method returns by declaring a variable or assigning to an existing variable.</span></span> <span data-ttu-id="19c3e-221">次のようなステートメントの左側に表示されます `const table = worksheet.addTable('A1:D6', true)` 。</span><span class="sxs-lookup"><span data-stu-id="19c3e-221">You can see that on the left hand side of statement such as `const table = worksheet.addTable('A1:D6', true)`.</span></span>
* <span data-ttu-id="19c3e-222">ほとんどの場合、Office Scripts オブジェクト モデルは、オブジェクト モデルのさまざまな部分をリンクするメソッドを持つExcel構成されています。</span><span class="sxs-lookup"><span data-stu-id="19c3e-222">For the most part, the Office Scripts object model consists of objects with methods that link various parts of the Excel object model.</span></span> <span data-ttu-id="19c3e-223">ごくまれに、スカラー値またはオブジェクト値のプロパティに出くることはめったに起こりません。</span><span class="sxs-lookup"><span data-stu-id="19c3e-223">Very rarely you'll come across properties that are of scalar or object values.</span></span>
* <span data-ttu-id="19c3e-224">[Officeスクリプト] では、Excel オブジェクト モデル メソッドには、開いているかっこを含む必要があります。</span><span class="sxs-lookup"><span data-stu-id="19c3e-224">In Office Scripts, an Excel object model method has to contain open-close parentheses.</span></span> <span data-ttu-id="19c3e-225">メソッドを指定せずにメソッドを使用する (変数へのメソッドの割り当てなど) は許可されません。</span><span class="sxs-lookup"><span data-stu-id="19c3e-225">Using methods without them is not allowed (such as assigning a method to a variable).</span></span>

<span data-ttu-id="19c3e-226">オブジェクトのいくつかのメソッドを見 `workbook` てみよ。</span><span class="sxs-lookup"><span data-stu-id="19c3e-226">Let's look at a few methods on the `workbook` object.</span></span>

```TypeScript
function main(workbook: ExcelScript.Workbook) {
    // Return a boolean (true or false) setting of whether the workbook is set to auto-save or not. 
    const autoSave = workbook.getAutoSave(); 
    // Get workbook name.
    const name = workbook.getName();
    // Get active cell range object.
    const cell = workbook.getActiveCell();
    // Get table named SALES.
    const cell = workbook.getTable('SALES');
    // Get all slicer objects.
    const slicers = workbook.getSlicers();
}
```

<span data-ttu-id="19c3e-227">この例では次のようになっています。</span><span class="sxs-lookup"><span data-stu-id="19c3e-227">In this example:</span></span>

* <span data-ttu-id="19c3e-228">スカラー プロパティ (文字列、数値、ブール型) などのオブジェクトの `workbook` `getAutoSave()` `getName()` メソッド。</span><span class="sxs-lookup"><span data-stu-id="19c3e-228">The methods of the `workbook` object such as `getAutoSave()` and `getName()` return a scalar property (string, number, boolean).</span></span>
* <span data-ttu-id="19c3e-229">別のオブジェクトを `getActiveCell()` 返すなどのメソッド。</span><span class="sxs-lookup"><span data-stu-id="19c3e-229">Methods such as `getActiveCell()` return another object.</span></span>
* <span data-ttu-id="19c3e-230">メソッド `getTable()` は引数 (この場合はテーブル名) を受け取り、ブック内の特定のテーブルを返します。</span><span class="sxs-lookup"><span data-stu-id="19c3e-230">The `getTable()` method accepts an argument (table name in this case) and returns a specific table in the workbook.</span></span>
* <span data-ttu-id="19c3e-231">このメソッドは、ブック内のすべてのスライサー オブジェクトの配列 (多くの場所をコレクションと呼ばれます `getSlicers()` ) を返します。</span><span class="sxs-lookup"><span data-stu-id="19c3e-231">The `getSlicers()` method returns an array (referred to in many places as a collection) of all slicer objects within the workbook.</span></span>

<span data-ttu-id="19c3e-232">これらのメソッドのすべてがプレフィックスを持っています。これは、Office Scripts オブジェクト モデルでメソッドが何かを返すという規則にすら使用 `get` されます。</span><span class="sxs-lookup"><span data-stu-id="19c3e-232">You'll notice that all of these methods have a `get` prefix, which is just a convention used in the Office Scripts object model to convey that the method is returning something.</span></span> <span data-ttu-id="19c3e-233">これらは一般的に 'getters' とも呼ばれます。</span><span class="sxs-lookup"><span data-stu-id="19c3e-233">They are also commonly referred to as 'getters'.</span></span>

<span data-ttu-id="19c3e-234">次の例では、他に 2 種類のメソッドが表示されます。</span><span class="sxs-lookup"><span data-stu-id="19c3e-234">There are two other types of methods that we'll now see in the next example:</span></span>

```TypeScript
function main(workbook: ExcelScript.Workbook) {
    // Get a worksheet named 'Sheet1.
    const sheet = workbook.getWorksheet('Sheet1'); 
    // Set name to SALES.
    sheet.setName('SALES');
    // Position the worksheet at the beginning.
    sheet.setPosition(0);
}
```

<span data-ttu-id="19c3e-235">この例では次のようになっています。</span><span class="sxs-lookup"><span data-stu-id="19c3e-235">In this example:</span></span>

* <span data-ttu-id="19c3e-236">メソッド `setName()` は、ワークシートに新しい名前を設定します。</span><span class="sxs-lookup"><span data-stu-id="19c3e-236">The `setName()` method sets a new name to the worksheet.</span></span> <span data-ttu-id="19c3e-237">`setPosition()` 位置を最初のセルに設定します。</span><span class="sxs-lookup"><span data-stu-id="19c3e-237">`setPosition()` sets the position to the first cell.</span></span>
* <span data-ttu-id="19c3e-238">このようなメソッドは、Excelのプロパティまたは動作を設定して、ファイルを変更します。</span><span class="sxs-lookup"><span data-stu-id="19c3e-238">Such methods modify the Excel file by setting a property or behavior of the workbook.</span></span> <span data-ttu-id="19c3e-239">これらのメソッドは 'setters' と呼ばれる。</span><span class="sxs-lookup"><span data-stu-id="19c3e-239">These methods are called 'setters'.</span></span>
* <span data-ttu-id="19c3e-240">通常、'setters' にはコンパニオン 'getter' が含め、どちらもメソッド `worksheet.getPosition` `worksheet.setPosition` です。</span><span class="sxs-lookup"><span data-stu-id="19c3e-240">Typically 'setters' have a companion 'getter', for example, `worksheet.getPosition` and `worksheet.setPosition`, both of which are methods.</span></span>

#### <a name="undefined-and-null-primitive-types"></a><span data-ttu-id="19c3e-241">`undefined` プリミティブ `null` 型</span><span class="sxs-lookup"><span data-stu-id="19c3e-241">`undefined` and `null` primitive types</span></span>

<span data-ttu-id="19c3e-242">以下に注意する必要がある 2 つのプリミティブ データ型を示します。</span><span class="sxs-lookup"><span data-stu-id="19c3e-242">The following are two primitive data types that you must be aware of:</span></span>

1. <span data-ttu-id="19c3e-243">この値 [`null`](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/null) は、オブジェクト値が意図的に存在しなかっている場合を表します。</span><span class="sxs-lookup"><span data-stu-id="19c3e-243">The value [`null`](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/null) represents the intentional absence of any object value.</span></span> <span data-ttu-id="19c3e-244">これは JavaScript のプリミティブ値の 1 つであり、変数に値が含めないかどうかを示すために使用されます。</span><span class="sxs-lookup"><span data-stu-id="19c3e-244">It is one of JavaScript's primitive values and is used to indicate that a variable has no value.</span></span>
1. <span data-ttu-id="19c3e-245">値が割り当てられていない変数は型です [`undefined`](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/undefined) 。</span><span class="sxs-lookup"><span data-stu-id="19c3e-245">A variable that has not been assigned a value is of type [`undefined`](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/undefined).</span></span> <span data-ttu-id="19c3e-246">評価対象の変数に割り当てられた値が含られていない場合は、メソッドまたはステートメント `undefined` を返す場合があります。</span><span class="sxs-lookup"><span data-stu-id="19c3e-246">A method or statement can also return `undefined` if the variable that's being evaluated doesn't have an assigned value.</span></span>

<span data-ttu-id="19c3e-247">これら 2 つの種類は、エラー処理の一部としてトリミングされ、適切に処理されないと、かなり頭痛の種になる可能性があります。</span><span class="sxs-lookup"><span data-stu-id="19c3e-247">These two types crop up as part of error handling and can cause quite a bit of headache if not handled properly.</span></span> <span data-ttu-id="19c3e-248">幸いなことに、TypeScript/JavaScript は、変数が型または `undefined` `null` .</span><span class="sxs-lookup"><span data-stu-id="19c3e-248">Fortunately, TypeScript/JavaScript offers a way to check if a variable is of type `undefined` or `null`.</span></span> <span data-ttu-id="19c3e-249">これらのチェックの一部については、エラー処理など、後のセクションで説明します。</span><span class="sxs-lookup"><span data-stu-id="19c3e-249">We will talk about some of those checks in later sections, including error handling.</span></span>

#### <a name="method-chaining"></a><span data-ttu-id="19c3e-250">メソッドチェーン</span><span class="sxs-lookup"><span data-stu-id="19c3e-250">Method chaining</span></span>

<span data-ttu-id="19c3e-251">ドット表記を使用すると、メソッドから返されるオブジェクトを接続してコードを短縮できます。</span><span class="sxs-lookup"><span data-stu-id="19c3e-251">You can use dot notation to connect objects being returned from a method to shorten your code.</span></span> <span data-ttu-id="19c3e-252">この手法を使用すると、コードの読み取りと管理が容易な場合があります。</span><span class="sxs-lookup"><span data-stu-id="19c3e-252">Sometimes this technique makes the code easy to read and manage.</span></span> <span data-ttu-id="19c3e-253">ただし、注意する必要がある点は少ない。</span><span class="sxs-lookup"><span data-stu-id="19c3e-253">However, there are few things to be aware of.</span></span> <span data-ttu-id="19c3e-254">次の例を見てみよ。</span><span class="sxs-lookup"><span data-stu-id="19c3e-254">Let's look at the following examples.</span></span>

<span data-ttu-id="19c3e-255">次のコードは、アクティブ セルと次のセルを取得し、値を設定します。</span><span class="sxs-lookup"><span data-stu-id="19c3e-255">The following code gets the active cell and the next cell, then sets the value.</span></span> <span data-ttu-id="19c3e-256">これは、このコードがすべての時間で成功するためにチェーンを使用する良い候補です。</span><span class="sxs-lookup"><span data-stu-id="19c3e-256">This is a good candidate to use chaining as this code will succeed all the time.</span></span>

```TypeScript
function main(workbook: ExcelScript.Workbook) {
    workbook.getActiveCell().getOffsetRange(0,1).setValue('Next cell');
}
```

<span data-ttu-id="19c3e-257">ただし、次のコード **(SALES** という名前のテーブルを取得し、そのバンド列スタイルをオンにする) に問題があります。</span><span class="sxs-lookup"><span data-stu-id="19c3e-257">However, the following code (which gets a table named **SALES** and turns on its banded column style) has an issue.</span></span>

```TypeScript
function main(workbook: ExcelScript.Workbook) {
  workbook.getTable('SALES').setShowBandedColumns(true);
}
```

<span data-ttu-id="19c3e-258">SALES テーブル **が** 存在しない場合は、</span><span class="sxs-lookup"><span data-stu-id="19c3e-258">What if the **SALES** table doesn't exist?</span></span> <span data-ttu-id="19c3e-259">(SALES などのテーブルが存在しないことを示す JavaScript 型) を返すので、スクリプトはエラー (次に示す) で `getTable('SALES')` `undefined` 失敗 **します**。</span><span class="sxs-lookup"><span data-stu-id="19c3e-259">The script will fail with an error (shown next) because `getTable('SALES')` returns `undefined` (which is a JavaScript type indicating that there is no table such as **SALES**).</span></span> <span data-ttu-id="19c3e-260">on メソッドの呼び出しは意味がありません。つまり、スクリプト `setShowBandedColumns` `undefined` `undefined.setShowBandedColumns(true)` はエラーで終了します。</span><span class="sxs-lookup"><span data-stu-id="19c3e-260">Calling the `setShowBandedColumns` method on `undefined` makes no sense, that is, `undefined.setShowBandedColumns(true)`, and hence the script ends in an error.</span></span>

```text
Line 2: Cannot read property 'setShowBandedColumns' of undefined
```

<span data-ttu-id="19c3e-261">この条件を処理[](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Operators/Optional_chaining)するには、接続されたオブジェクトを介して値にアクセスする方法を提供するオプションのチェーン演算子を使用できます。参照またはメソッドが存在する可能性がある場合、または `undefined` (JavaScript の割り当てられていないオブジェクトまたは存在しないオブジェクトまたは結果を示す JavaScript の方法です)。 `null`</span><span class="sxs-lookup"><span data-stu-id="19c3e-261">You could use the [optional chaining operator](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Operators/Optional_chaining) that provides a way to simplify accessing values through connected objects when it's possible that a reference or method may be `undefined` or `null` (which is JavaScript's way of indicating an unassigned or nonexistent object or result) to handle this condition.</span></span>

```TypeScript
function main(workbook: ExcelScript.Workbook) {
    // This line will not fail as the setShowBandedColumns method is executed only if the SALES table is present.
    workbook.getTable('SALES')?.setShowBandedColumns(true); 
}
```

<span data-ttu-id="19c3e-262">メソッドによって返される存在しないオブジェクトの条件や型を処理する場合は、メソッドから戻り値を割り当て、それを個別に `undefined` 処理する方が良いです。</span><span class="sxs-lookup"><span data-stu-id="19c3e-262">If you wish to handle nonexistent object conditions or `undefined` type being returned by a method, then it is better to assign the return value from the method and handle that separately.</span></span>

```TypeScript
function main(workbook: ExcelScript.Workbook) {
    const salesTable = workbook.getTable('SALES');
    if (salesTable) {
        salesTable.setShowBandedColumns(true);
    } else { 
        // Handle this condition.
    }
}
```

#### <a name="get-object-reference"></a><span data-ttu-id="19c3e-263">オブジェクト参照の取得</span><span class="sxs-lookup"><span data-stu-id="19c3e-263">Get object reference</span></span>

<span data-ttu-id="19c3e-264">オブジェクト `workbook` は関数内でユーザーに与 `main` えられる。</span><span class="sxs-lookup"><span data-stu-id="19c3e-264">The `workbook` object is given to you in the `main` function.</span></span> <span data-ttu-id="19c3e-265">オブジェクトの使用を開始し `workbook` 、そのメソッドに直接アクセスできます。</span><span class="sxs-lookup"><span data-stu-id="19c3e-265">You can begin to use the `workbook` object and access its methods directly.</span></span>

```TypeScript
function main(workbook: ExcelScript.Workbook) {
    // Get workbook name.
    const name = workbook.getName();
    // Display name to console.
    console.log(name);
}
```

<span data-ttu-id="19c3e-266">ブック内の他のすべてのオブジェクトを使用する場合は、オブジェクトから始まり、探しているオブジェクトに移動するまで階層を `workbook` 下に移動します。</span><span class="sxs-lookup"><span data-stu-id="19c3e-266">For using all other objects within the workbook, begin with `workbook` object and go down the hierarchy until you get to the object you are looking for.</span></span> <span data-ttu-id="19c3e-267">オブジェクト参照を取得するには、メソッドを使用してオブジェクトをフェッチするか、次に示すようにオブジェクトのコレクション `get` を取得します。</span><span class="sxs-lookup"><span data-stu-id="19c3e-267">You can get the object reference by fetching the object using its `get` method or by retrieving the collection of objects as shown below:</span></span>

```TypeScript
function main(workbook: ExcelScript.Workbook) {
    // Get the active worksheet.
    const sheet = workbook.getActiveWorksheet();
    // Fetch using an ID or key.
    const sheet = workbook.getWorksheet('SomeSheetName');
    // Invoke methods on the object.
    sheet.setPosition(0); 
    
    // Get collection of methods.
    const tables = sheet.getTables();
    console.log('Total tables in this sheet: ' + tables.length);
}
```

#### <a name="check-if-an-object-exists-then-delete-and-add"></a><span data-ttu-id="19c3e-268">オブジェクトが存在するかどうかを確認し、削除して追加する</span><span class="sxs-lookup"><span data-stu-id="19c3e-268">Check if an object exists, then delete, and add</span></span>

<span data-ttu-id="19c3e-269">定義済みの名前でオブジェクトを作成する場合は、常に存在する類似のオブジェクトを削除してから追加する方が良いです。</span><span class="sxs-lookup"><span data-stu-id="19c3e-269">For creating an object, say with a predefined name, it is always better to remove a similar object that may exist and then add it.</span></span> <span data-ttu-id="19c3e-270">これを行うには、次のパターンを使用します。</span><span class="sxs-lookup"><span data-stu-id="19c3e-270">You can do that using the following pattern.</span></span>

```TypeScript
function main(workbook: ExcelScript.Workbook) {
  // Name of the worksheet to be added. 
  let name = "Index";
  // Check if the worksheet already exists. If not, add the worksheet.
  let sheet = workbook.getWorksheet('Index');
  if (sheet) {
    console.log(`Worksheet by the name ${name} already exists. Deleting it.`);
    // Call the delete method on the object to remove it. 
    sheet.delete();
  } 
    // Add a blank worksheet. 
  console.log(`Adding the worksheet named  ${name}.`)
  const indexSheet = workbook.addWorksheet("Index");
}

```

<span data-ttu-id="19c3e-271">または、存在する可能性があるオブジェクトまたは存在しない可能性があるオブジェクトを削除するには、次のパターンを使用します。</span><span class="sxs-lookup"><span data-stu-id="19c3e-271">Alternatively, for deleting an object that may or may not exist, use the following pattern.</span></span>

```TypeScript
    // The ? preceding delete() will ensure that the API is only invoked if the object exists. 
    workbook.getWorksheet('Index')?.delete(); 
```

#### <a name="note-about-adding-an-object"></a><span data-ttu-id="19c3e-272">オブジェクトの追加に関する注意</span><span class="sxs-lookup"><span data-stu-id="19c3e-272">Note about adding an object</span></span>

<span data-ttu-id="19c3e-273">スライサー、ピボット テーブル、ワークシートなどのオブジェクトを作成、挿入、または追加するには、対応するメソッド **add_Object_します。**</span><span class="sxs-lookup"><span data-stu-id="19c3e-273">To create, insert, or add an object such as a slicer, pivot table, worksheet, etc., use the corresponding **add_Object_** method.</span></span> <span data-ttu-id="19c3e-274">このようなメソッドは、親オブジェクトで使用できます。</span><span class="sxs-lookup"><span data-stu-id="19c3e-274">Such a method is available on its parent object.</span></span> <span data-ttu-id="19c3e-275">たとえば、メソッドは `addChart()` オブジェクトで使用 `worksheet` できます。</span><span class="sxs-lookup"><span data-stu-id="19c3e-275">For example, the `addChart()` method is available on `worksheet` object.</span></span> <span data-ttu-id="19c3e-276">add_Object_ **メソッド** は、作成するオブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="19c3e-276">The **add_Object_** method returns the object it creates.</span></span> <span data-ttu-id="19c3e-277">返された値を受け取り、後でスクリプトで使用します。</span><span class="sxs-lookup"><span data-stu-id="19c3e-277">Receive the returned value and use it later in your script.</span></span>

```TypeScript
function main(workbook: ExcelScript.Workbook) {
  // Add object and get a reference to it. 
  const indexSheet = workbook.addWorksheet("Index");
  // Use it elsewhere in the script 
  console.log(indexSheet.getPosition());
}

```

<span data-ttu-id="19c3e-278">または、存在する可能性があるオブジェクトまたは存在しない可能性があるオブジェクトを削除するには、次のパターンを使用します。</span><span class="sxs-lookup"><span data-stu-id="19c3e-278">Alternatively, for deleting an object that may or may not exist, use this pattern:</span></span>

```TypeScript
    workbook.getWorksheet('Index')?.delete(); // The ? preceding delete() will ensure that the API is only invoked if the object exists. 
```

#### <a name="collections"></a><span data-ttu-id="19c3e-279">コレクション</span><span class="sxs-lookup"><span data-stu-id="19c3e-279">Collections</span></span>

<span data-ttu-id="19c3e-280">コレクションは、テーブル、グラフ、列などのオブジェクトで、配列として取得し、処理のために反復処理できます。</span><span class="sxs-lookup"><span data-stu-id="19c3e-280">Collections are objects such as tables, charts, columns, etc. that can be retrieved as an array and iterated over for processing.</span></span> <span data-ttu-id="19c3e-281">対応するメソッドを使用してコレクションを取得し、次のような多くの TypeScript 配列トラバーサル手法のいずれかを使用して、ループ内のデータを `get` 処理できます。</span><span class="sxs-lookup"><span data-stu-id="19c3e-281">You can retrieve a collection using the corresponding `get` method and process the data in a loop using one of many TypeScript array traversal techniques such as:</span></span>

* [<span data-ttu-id="19c3e-282">`for` または `while`</span><span class="sxs-lookup"><span data-stu-id="19c3e-282">`for` or `while`</span></span>](https://developer.mozilla.org/docs/Web/JavaScript/Guide/Loops_and_iteration)
* [`for..of`](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Statements/for...of)
* [`forEach`](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Array/forEach)

* [<span data-ttu-id="19c3e-283">配列の言語の基本</span><span class="sxs-lookup"><span data-stu-id="19c3e-283">Language basics of arrays</span></span>](https://developer.mozilla.org//docs/Learn/JavaScript/First_steps/Arrays)

<span data-ttu-id="19c3e-284">このスクリプトは、スクリプト API でサポートされているコレクションを使用Office示します。</span><span class="sxs-lookup"><span data-stu-id="19c3e-284">This script demonstrates how to use collections supported in Office Scripts APIs.</span></span> <span data-ttu-id="19c3e-285">ファイル内の各ワークシート タブにランダムな色を設定します。</span><span class="sxs-lookup"><span data-stu-id="19c3e-285">It colors each worksheet tab in the file with a random color.</span></span>

```TypeScript
function main(workbook: ExcelScript.Workbook) {

  // Get all sheets as a collection.
  const sheets = workbook.getWorksheets();
  const names = sheets.map ((sheet) => sheet.getName());
  console.log(names);
  console.log(`Total worksheets inside of this workbook: ${sheets.length}`);
  // Get information from specific sheets within the collection.
  console.log(`First sheet name is: ${names[0]}`);
  if (sheets.length > 1) {
    console.log(`Last sheet's Id is: ${sheets[sheets.length -1].getId()}`);
  }
  // Color each worksheet with random color.
  for (const sheet of sheets) {
    sheet.setTabColor(`#${Math.random().toString(16).substr(-6)}`);
  }
}
```

## <a name="type-declarations"></a><span data-ttu-id="19c3e-286">型宣言</span><span class="sxs-lookup"><span data-stu-id="19c3e-286">Type declarations</span></span>

<span data-ttu-id="19c3e-287">型宣言は、ユーザーが扱う変数の種類を理解するのに役立ちます。</span><span class="sxs-lookup"><span data-stu-id="19c3e-287">Type declarations help users understand the type of variable they are dealing with.</span></span> <span data-ttu-id="19c3e-288">メソッドの自動補完に役立ち、開発時間の品質チェックを支援します。</span><span class="sxs-lookup"><span data-stu-id="19c3e-288">It helps with auto-completion of methods and assists in development time quality checks.</span></span>

<span data-ttu-id="19c3e-289">スクリプト内の型宣言は、関数宣言、変数宣言、定義の定義など、さまざまなIntelliSense検索できます。</span><span class="sxs-lookup"><span data-stu-id="19c3e-289">You can find type declarations in the script in various places including function declaration, variable declaration, IntelliSense definitions, etc.</span></span>

<span data-ttu-id="19c3e-290">例:</span><span class="sxs-lookup"><span data-stu-id="19c3e-290">Examples:</span></span>

* `function main(workbook: ExcelScript.Workbook)`
* `let myRange: ExcelScript.Range;`
* `function getMaxAmount(range: ExcelScript.Range): number`

<span data-ttu-id="19c3e-291">通常は異なる色で明確に表示されるので、コード エディターで型を簡単に識別できます。</span><span class="sxs-lookup"><span data-stu-id="19c3e-291">You can identify the types easily in the code editor as it usually appears distinctly in a different color.</span></span> <span data-ttu-id="19c3e-292">通常、コロン `:` は型宣言の前に表示されます。</span><span class="sxs-lookup"><span data-stu-id="19c3e-292">A colon `:` usually precedes the type declaration.</span></span>  

<span data-ttu-id="19c3e-293">TypeScript では、追加のコードを記述せずに大きな力を得る可能性がある型推論を使用することで、書き込み型を省略できます。</span><span class="sxs-lookup"><span data-stu-id="19c3e-293">Writing types can be optional in TypeScript because type inference allows you to get a lot of power without writing additional code.</span></span> <span data-ttu-id="19c3e-294">ほとんどの場合、TypeScript 言語は変数の種類を確認するのが優れた方法です。</span><span class="sxs-lookup"><span data-stu-id="19c3e-294">For the most part, the TypeScript language is good at inferring the types of variables.</span></span> <span data-ttu-id="19c3e-295">ただし、特定の場合Officeスクリプトでは、言語が型を明確に識別できない場合は、型宣言を明示的に定義する必要があります。</span><span class="sxs-lookup"><span data-stu-id="19c3e-295">However, in certain cases, Office Scripts require the type declarations to be explicitly defined if the language is unable to clearly identify the type.</span></span> <span data-ttu-id="19c3e-296">また、スクリプトでは明示的 `any` または暗黙的Officeされません。</span><span class="sxs-lookup"><span data-stu-id="19c3e-296">Also, explicit or implicit `any` is not allowed in Office Script.</span></span> <span data-ttu-id="19c3e-297">その詳細については、後で説明します。</span><span class="sxs-lookup"><span data-stu-id="19c3e-297">More on that later.</span></span>

### <a name="excelscript-types"></a><span data-ttu-id="19c3e-298">`ExcelScript` 型</span><span class="sxs-lookup"><span data-stu-id="19c3e-298">`ExcelScript` types</span></span>

<span data-ttu-id="19c3e-299">[Officeスクリプト] では、次の種類を使用します。</span><span class="sxs-lookup"><span data-stu-id="19c3e-299">In Office Scripts, you will use the following kinds of types.</span></span>

* <span data-ttu-id="19c3e-300">、、、、など、ネイティブ `number` `string` `object` `boolean` 言語 `null` の種類。</span><span class="sxs-lookup"><span data-stu-id="19c3e-300">Native language types such as `number`, `string`, `object`, `boolean`, `null`, etc.</span></span>
* <span data-ttu-id="19c3e-301">ExcelAPI の種類。</span><span class="sxs-lookup"><span data-stu-id="19c3e-301">Excel API types.</span></span> <span data-ttu-id="19c3e-302">これらはで始まります `ExcelScript` 。</span><span class="sxs-lookup"><span data-stu-id="19c3e-302">They begin with `ExcelScript`.</span></span> <span data-ttu-id="19c3e-303">たとえば `ExcelScript.Range` `ExcelScript.Table` 、、、など。</span><span class="sxs-lookup"><span data-stu-id="19c3e-303">For example, `ExcelScript.Range`, `ExcelScript.Table`, etc.</span></span>
* <span data-ttu-id="19c3e-304">ステートメントを使用してスクリプトで定義したカスタム インターフェイス `interface` 。</span><span class="sxs-lookup"><span data-stu-id="19c3e-304">Any custom interfaces you may have defined in the script using `interface` statements.</span></span>

<span data-ttu-id="19c3e-305">次に、これらの各グループの例を参照してください。</span><span class="sxs-lookup"><span data-stu-id="19c3e-305">See examples of each of these groups next.</span></span>

<span data-ttu-id="19c3e-306">**_ネイティブ言語の種類_**</span><span class="sxs-lookup"><span data-stu-id="19c3e-306">**_Native language types_**</span></span>

<span data-ttu-id="19c3e-307">次の例では、場所 、、および `string` `number` 使用されている `boolean` 場所に注意してください。</span><span class="sxs-lookup"><span data-stu-id="19c3e-307">In the following example, notice places where `string`, `number`, and `boolean` have been used.</span></span> <span data-ttu-id="19c3e-308">これらは、ネイティブ **の TypeScript** 言語の種類です。</span><span class="sxs-lookup"><span data-stu-id="19c3e-308">These are native **TypeScript** language types.</span></span>

```TypeScript
function main(workbook: ExcelScript.Workbook)
{
  const table = workbook.getActiveWorksheet().getTables()[0];
  const sales = table.getColumnByName('Sales').getRange().getValues();
  console.log(sales);
  // Add 100 to each value.
  const revisedSales = salesAs1DArray.map(data => data as number + 100);
  // Add a column.
  table.addColumn(-1, revisedSales);  
}
/**
 * Extract a column from 2D array and return result.
 */
function extractColumn(data: (string | number | boolean)[][], index: number): (string | number | boolean)[] {

  const column = data.map((row) => {
    return row[index];
  })
  return column;
}
/**
 * Convert a flat array into a 2D array that can be used as range column.
 */
function convertColumnTo2D(data: (string | number | boolean)[]): (string | number | boolean)[][] {

  const columnAs2D = data.map((row) => {
    return [row];
  })
  return columnAs2D;
}
```

<span data-ttu-id="19c3e-309">**_ExcelScript の種類_**</span><span class="sxs-lookup"><span data-stu-id="19c3e-309">**_ExcelScript types_**</span></span>

<span data-ttu-id="19c3e-310">次の例では、ヘルパー関数は 2 つの引数を受け取ります。</span><span class="sxs-lookup"><span data-stu-id="19c3e-310">In the following example, a helper function takes two arguments.</span></span> <span data-ttu-id="19c3e-311">最初の変数は `sheet` 型の変数 `ExcelScript.Worksheet` です。</span><span class="sxs-lookup"><span data-stu-id="19c3e-311">The first one is the `sheet` variable which is of type `ExcelScript.Worksheet` type.</span></span>

```TypeScript
function main(workbook: ExcelScript.Workbook) {
    const sheet = workbook.getWorksheet('Sheet5');
    const data = ['2016', 'Bikes', 'Seats', '1500', .05];
    addRow(sheet, data);
    return;
}

function addRow(sheet: ExcelScript.Worksheet, data: (string | number | boolean)[]): void {

    const usedRange = sheet.getUsedRange();
    let startCell: ExcelScript.Range;
    // If the sheet is empty, then use A1 as starting cell for update.
    if (usedRange) { 
      startCell = usedRange.getLastRow().getCell(0, 0).getOffsetRange(1, 0);
    } else {
      startCell = sheet.getRange('A1');
    }
    console.log(startCell.getAddress());
    const targetRange = startCell.getResizedRange(0, data.length - 1);      
    targetRange.setValues([data]);
    return;
}
```

<span data-ttu-id="19c3e-312">**_カスタム型_**</span><span class="sxs-lookup"><span data-stu-id="19c3e-312">**_Custom types_**</span></span>

<span data-ttu-id="19c3e-313">カスタム インターフェイスは `ReportImages` 、イメージを別のフロー アクションに戻す場合に使用します。</span><span class="sxs-lookup"><span data-stu-id="19c3e-313">The custom interface `ReportImages` is used to return images to another flow action.</span></span> <span data-ttu-id="19c3e-314">関数 `main` 宣言には、 `: ReportImages` その型のオブジェクトが返されるという TypeScript を指示する命令が含まれています。</span><span class="sxs-lookup"><span data-stu-id="19c3e-314">The `main` function declaration includes `: ReportImages` instruction to tell TypeScript that an object of that type is being returned.</span></span>

```TypeScript
function main(workbook: ExcelScript.Workbook): ReportImages {
  let chart = workbook.getWorksheet("Sheet1").getCharts()[0];
  const table = workbook.getWorksheet('InvoiceAmounts').getTables()[0];
  
  const chartImage = chart.getImage();
  const tableImage = table.getRange().getImage();
  return {
    chartImage,
    tableImage
  }
}

interface ReportImages {
  chartImage: string
  tableImage: string
}
```

### <a name="type-assertion-overriding-the-type"></a><span data-ttu-id="19c3e-315">型アサーション (型のオーバーライド)</span><span class="sxs-lookup"><span data-stu-id="19c3e-315">Type assertion (overriding the type)</span></span>

<span data-ttu-id="19c3e-316">TypeScript のドキュメント [には、「TypeScript](https://www.typescriptlang.org/docs/handbook/basic-types.html#type-assertions) よりも値の詳細が分かっている状況に終わる場合があります。</span><span class="sxs-lookup"><span data-stu-id="19c3e-316">As the TypeScript [documentation](https://www.typescriptlang.org/docs/handbook/basic-types.html#type-assertions) states, "Sometimes you'll end up in a situation where you'll know more about a value than TypeScript does.</span></span> <span data-ttu-id="19c3e-317">通常、これは、エンティティの種類が現在の型よりも具体的である可能性があることを知っている場合に発生します。</span><span class="sxs-lookup"><span data-stu-id="19c3e-317">Usually, this will happen when you know the type of some entity could be more specific than its current type.</span></span> <span data-ttu-id="19c3e-318">型アサーションは、コンパイラに 「信頼して、自分が何をしているのかを知っている」ことを伝える方法です。</span><span class="sxs-lookup"><span data-stu-id="19c3e-318">Type assertions are a way to tell the compiler “trust me, I know what I'm doing.”</span></span> <span data-ttu-id="19c3e-319">型アサーションは、他の言語でキャストされる型に似ていますが、特別なチェックやデータの再構築は実行します。</span><span class="sxs-lookup"><span data-stu-id="19c3e-319">A type assertion is like a type cast in other languages, but it performs no special checking or restructuring of data.</span></span> <span data-ttu-id="19c3e-320">実行時に影響を与え、コンパイラによって純粋に使用されます。</span><span class="sxs-lookup"><span data-stu-id="19c3e-320">It has no runtime impact and is used purely by the compiler."</span></span>

<span data-ttu-id="19c3e-321">次のコードに示すように、キーワードを使用するか、角かっこを使用して `as` 型をアサートできます。</span><span class="sxs-lookup"><span data-stu-id="19c3e-321">You can assert the type using the `as` keyword or using angle brackets as shown in following code.</span></span>

```TypeScript
function main(workbook: ExcelScript.Workbook) {
  let data = workbook.getActiveCell().getValue();
  // Since the add10 function only accepts number, assert data's type as number, otherwise the script cannot be run.
  const answer1 = add10(data as number);
  const answer2 = add10(<number> data);
}

function add10(data: number) { 
  return data + 10;
}
```

#### <a name="any-type-in-the-script"></a><span data-ttu-id="19c3e-322">スクリプト内の 'any' 型</span><span class="sxs-lookup"><span data-stu-id="19c3e-322">'any' type in the script</span></span>

<span data-ttu-id="19c3e-323">[TypeScript Web サイトの状態](https://www.typescriptlang.org/docs/handbook/basic-types.html#any):</span><span class="sxs-lookup"><span data-stu-id="19c3e-323">The [TypeScript website states](https://www.typescriptlang.org/docs/handbook/basic-types.html#any):</span></span>

  <span data-ttu-id="19c3e-324">一部の状況では、すべての型情報が使用できるとは言え、宣言に不適切な労力がかかる場合があります。</span><span class="sxs-lookup"><span data-stu-id="19c3e-324">In some situations, not all type information is available or its declaration would take an inappropriate amount of effort.</span></span> <span data-ttu-id="19c3e-325">これらは、TypeScript またはサードパーティ ライブラリなしで記述されたコードの値に対して発生する可能性があります。</span><span class="sxs-lookup"><span data-stu-id="19c3e-325">These may occur for values from code that has been written without TypeScript or a 3rd party library.</span></span> <span data-ttu-id="19c3e-326">このような場合は、型チェックをオプトアウトする必要があります。</span><span class="sxs-lookup"><span data-stu-id="19c3e-326">In these cases, we might want to opt-out of type checking.</span></span> <span data-ttu-id="19c3e-327">これを行うには、これらの値に次の種類のラベルを付 `any` します。</span><span class="sxs-lookup"><span data-stu-id="19c3e-327">To do so, we label these values with the `any` type:</span></span>

  ```TypeScript
  declare function getValue(key: string): any;
  // OK, return value of 'getValue' is not checked
  const str: string = getValue("myString");
  ```

<span data-ttu-id="19c3e-328">**明示的 `any` は許可されません**</span><span class="sxs-lookup"><span data-stu-id="19c3e-328">**Explicit `any` is NOT allowed**</span></span>

```TypeScript
// This is not allowed
let someVariable: any; 
```

<span data-ttu-id="19c3e-329">この `any` 型は、スクリプトが API を処理Office方法にExcelします。</span><span class="sxs-lookup"><span data-stu-id="19c3e-329">The `any` type presents challenges to the way Office Scripts processes the Excel APIs.</span></span> <span data-ttu-id="19c3e-330">変数が API に送信され、処理Excel問題が発生します。</span><span class="sxs-lookup"><span data-stu-id="19c3e-330">It causes issues when the variables are sent to Excel APIs for processing.</span></span> <span data-ttu-id="19c3e-331">スクリプトで使用される変数の種類を知ることは、スクリプトの処理に不可欠であるため、型を持つ変数の明示的な定義 `any` は禁止されています。</span><span class="sxs-lookup"><span data-stu-id="19c3e-331">Knowing the type of variables used in the script is essential to the processing of script and hence explicit definition of any variable with `any` type is prohibited.</span></span> <span data-ttu-id="19c3e-332">スクリプトで型が宣言された変数がある場合は、コンパイル時エラー (スクリプトを実行する前のエラー) `any` が表示されます。</span><span class="sxs-lookup"><span data-stu-id="19c3e-332">You will receive a compile-time error (error prior to running the script) if there is any variable with `any` type declared in the script.</span></span> <span data-ttu-id="19c3e-333">エディターにもエラーが表示されます。</span><span class="sxs-lookup"><span data-stu-id="19c3e-333">You will see an error in the editor as well.</span></span>

:::image type="content" source="../../images/getting-started-eanyi.png" alt-text="明示的な 'any' エラー":::

:::image type="content" source="../../images/getting-started-expany.png" alt-text="Output に表示される明示的な 'any' エラー":::

<span data-ttu-id="19c3e-336">前の図に表示されたコードでは、 `[5, 16] Explicit Any is not allowed` 行 5 列 16 が型を宣言します `any` 。</span><span class="sxs-lookup"><span data-stu-id="19c3e-336">In the code displayed in the previous image, `[5, 16] Explicit Any is not allowed` indicates that line 5 column 16 declares the `any` type.</span></span> <span data-ttu-id="19c3e-337">これにより、エラーを含むコード行を見つけるのに役立ちます。</span><span class="sxs-lookup"><span data-stu-id="19c3e-337">This helps you locate the line of code that contains the error.</span></span>

<span data-ttu-id="19c3e-338">この問題を回避するには、常に変数の型を宣言します。</span><span class="sxs-lookup"><span data-stu-id="19c3e-338">To get around this issue, always declare the type of the variable.</span></span>

<span data-ttu-id="19c3e-339">変数の種類が不明な場合は、TypeScript の 1 つのクールなトリックを使用すると、共用体の型 [を定義できます](https://www.typescriptlang.org/docs/handbook/unions-and-intersections.html)。</span><span class="sxs-lookup"><span data-stu-id="19c3e-339">If you are uncertain about the type of a variable, one cool trick in TypeScript allows you to define [union types](https://www.typescriptlang.org/docs/handbook/unions-and-intersections.html).</span></span> <span data-ttu-id="19c3e-340">これは、変数が範囲の値を保持する場合に使用できます。これは、多くの型を使用できます。</span><span class="sxs-lookup"><span data-stu-id="19c3e-340">This can be used for variables to hold a range values, which can be of many types.</span></span>

```TypeScript
// Define value as a union type rather than 'any' type.
let value: (string | number | boolean);
value = someValue_from_another_source;
//...
someRange.setValue(value);
```

### <a name="type-inference"></a><span data-ttu-id="19c3e-341">型の推論</span><span class="sxs-lookup"><span data-stu-id="19c3e-341">Type inference</span></span>

<span data-ttu-id="19c3e-342">TypeScript では、明示的な型注釈[](https://www.typescriptlang.org/docs/handbook/type-inference.html)がない場合に型の情報を提供するために型推論を使用する場所がいくつかあります。</span><span class="sxs-lookup"><span data-stu-id="19c3e-342">In TypeScript, there are several places where [type inference](https://www.typescriptlang.org/docs/handbook/type-inference.html) is used to provide type information when there is no explicit type annotation.</span></span> <span data-ttu-id="19c3e-343">たとえば、x 変数の型は、次のコードの数値と推測されます。</span><span class="sxs-lookup"><span data-stu-id="19c3e-343">For example, the type of the x variable is inferred to be a number in the following code.</span></span>

```TypeScript
let x = 3;
//  ^ = let x: number
```

<span data-ttu-id="19c3e-344">この種の推論は、変数とメンバーを初期化し、パラメーターの既定値を設定し、関数の戻り値の型を決定するときに行います。</span><span class="sxs-lookup"><span data-stu-id="19c3e-344">This kind of inference takes place when initializing variables and members, setting parameter default values, and determining function return types.</span></span>

### <a name="no-implicit-any-rule"></a><span data-ttu-id="19c3e-345">暗黙的なしルール</span><span class="sxs-lookup"><span data-stu-id="19c3e-345">no-implicit-any rule</span></span>

<span data-ttu-id="19c3e-346">スクリプトでは、明示的または暗黙的に宣言するために使用される変数の種類が必要です。</span><span class="sxs-lookup"><span data-stu-id="19c3e-346">A script requires the types of the variables used to be explicitly or implicitly declared.</span></span> <span data-ttu-id="19c3e-347">TypeScript コンパイラが変数の種類を特定できない場合 (型が明示的に宣言されていないか、型の推論ができないため)、コンパイル時間エラーが発生します (スクリプトを実行する前にエラーが発生します)。</span><span class="sxs-lookup"><span data-stu-id="19c3e-347">If the TypeScript compiler is unable to determine the type of a variable (either because type is not declared explicitly or type inference is not possible), then you will receive a compilation time error (error prior to running the script).</span></span> <span data-ttu-id="19c3e-348">エディターにもエラーが表示されます。</span><span class="sxs-lookup"><span data-stu-id="19c3e-348">You will see an error in the editor as well.</span></span>

:::image type="content" source="../../images/getting-started-iany.png" alt-text="エディターに表示される暗黙的な 'any' エラー":::

<span data-ttu-id="19c3e-350">変数は型なしで宣言され、TypeScript は宣言時に型を特定できないので、次のスクリプトではコンパイル時間エラーが発生します。</span><span class="sxs-lookup"><span data-stu-id="19c3e-350">The following scripts have compilation time errors because variables are declared without types and TypeScript cannot determine the type at the time of declaration.</span></span>

```TypeScript
function main(workbook: ExcelScript.Workbook) {
    // The variable 'value' gets 'any' type
    // because no type is declared.
    let value; 
    // Even when a number type is assigned,
    // the type of 'value' remains any.
    value = 10; 
    // The following statement fails because
    // Office Scripts can't send an argument
    // of type 'any' to Excel for processing.
    workbook.getActiveCell().setValue(value);
    return;
}
```

```TypeScript
function main(workbook: ExcelScript.Workbook) {
    // The variable 'cell' gets 'any' type
    // because no type is defined.
    let cell; 
    cell = workbook.getActiveCell().getValue();
    // Office Scripts can't assign Range type object
    // to a variable of 'any' type.
    console.log(cell.getValue());
    return;
}
```

<span data-ttu-id="19c3e-351">このエラーを回避するには、代わりに次のパターンを使用します。</span><span class="sxs-lookup"><span data-stu-id="19c3e-351">To avoid this error, use the following patterns instead.</span></span> <span data-ttu-id="19c3e-352">それぞれの場合、変数とその型は同時に宣言されます。</span><span class="sxs-lookup"><span data-stu-id="19c3e-352">In each case, the variable and its type are declared at the same time.</span></span>

```TypeScript
function main(workbook: ExcelScript.Workbook) {
    const value: number = 10; 
    workbook.getActiveCell().setValue(value);
    return;
}
```

```TypeScript
function main(workbook: ExcelScript.Workbook) {
    const cell: ExcelScript.Range = workbook.getActiveCell().getValue();
    console.log(cell.getValue()); 
    return;
}
```

## <a name="error-handling"></a><span data-ttu-id="19c3e-353">エラー処理</span><span class="sxs-lookup"><span data-stu-id="19c3e-353">Error handling</span></span>

<span data-ttu-id="19c3e-354">Officeスクリプト エラーは、次のいずれかのカテゴリに分類できます。</span><span class="sxs-lookup"><span data-stu-id="19c3e-354">Office Scripts error can be classified into one of the following categories.</span></span>

1. <span data-ttu-id="19c3e-355">エディターに表示されるコンパイル時の警告</span><span class="sxs-lookup"><span data-stu-id="19c3e-355">Compile-time warning shown in the editor</span></span>
1. <span data-ttu-id="19c3e-356">実行時に表示されますが、実行が開始される前に発生するコンパイル時エラー</span><span class="sxs-lookup"><span data-stu-id="19c3e-356">Compile-time error that appears when you run but occurs before execution begins</span></span>
1. <span data-ttu-id="19c3e-357">ランタイム エラー</span><span class="sxs-lookup"><span data-stu-id="19c3e-357">Runtime error</span></span>

<span data-ttu-id="19c3e-358">エディターの警告は、エディターの波状の赤い下線を使用して識別できます。</span><span class="sxs-lookup"><span data-stu-id="19c3e-358">Editor warnings can be identified using the wavy red underlines in the editor:</span></span>

:::image type="content" source="../../images/getting-started-eanyi.png" alt-text="エディターに表示されるコンパイル時の警告":::

<span data-ttu-id="19c3e-360">オレンジ色の警告の下線と灰色の情報メッセージが表示される場合があります。</span><span class="sxs-lookup"><span data-stu-id="19c3e-360">At times, you may also see orange warning underlines and grey informational messages.</span></span> <span data-ttu-id="19c3e-361">エラーは発生しませんが、密接に調べる必要があります。</span><span class="sxs-lookup"><span data-stu-id="19c3e-361">They should be examined closely though they are not going to cause errors.</span></span>

<span data-ttu-id="19c3e-362">両方のエラー メッセージが同一に見えるので、コンパイル時エラーと実行時エラーを区別することはできません。</span><span class="sxs-lookup"><span data-stu-id="19c3e-362">It isn't possible to distinguish between compile-time and runtime errors as both error messages look identical.</span></span> <span data-ttu-id="19c3e-363">これらはどちらも、実際にスクリプトを実行するときに発生します。</span><span class="sxs-lookup"><span data-stu-id="19c3e-363">They both occur when you actually execute the script.</span></span> <span data-ttu-id="19c3e-364">次の図は、コンパイル時エラーと実行時エラーの例を示しています。</span><span class="sxs-lookup"><span data-stu-id="19c3e-364">The following images show examples of a compile-time error and a runtime error.</span></span>

:::image type="content" source="../../images/getting-started-expany.png" alt-text="コンパイル時エラーの例":::

:::image type="content" source="../../images/getting-started-error-basic.png" alt-text="実行時エラーの例":::

<span data-ttu-id="19c3e-367">どちらの場合も、エラーが発生した行番号が表示されます。</span><span class="sxs-lookup"><span data-stu-id="19c3e-367">In both cases, you will see the line number where the error occurred.</span></span> <span data-ttu-id="19c3e-368">その後、コードを確認し、問題を解決し、もう一度実行できます。</span><span class="sxs-lookup"><span data-stu-id="19c3e-368">You can then examine the code, fix the issue, and run again.</span></span>

<span data-ttu-id="19c3e-369">ランタイム エラーを回避するためのベスト プラクティスを次に示します。</span><span class="sxs-lookup"><span data-stu-id="19c3e-369">Following are a few best practices to avoid runtime errors.</span></span>

### <a name="check-for-object-existence-before-deletion"></a><span data-ttu-id="19c3e-370">削除前にオブジェクトの存在を確認する</span><span class="sxs-lookup"><span data-stu-id="19c3e-370">Check for object existence before deletion</span></span>

<span data-ttu-id="19c3e-371">または、存在する可能性があるオブジェクトまたは存在しない可能性があるオブジェクトを削除するには、次のパターンを使用します。</span><span class="sxs-lookup"><span data-stu-id="19c3e-371">Alternatively, for deleting an object that may or may not exist, use this pattern:</span></span>

```TypeScript
// The ? ensures that the delete() API is only invoked if the object exists.
workbook.getWorksheet('Index')?.delete();

// Alternative:
const indexSheet = workbook.getWorksheet('Index');
if (indexSheet) {
    indexSheet.delete();
}
```

### <a name="do-pre-checks-at-the-beginning-of-the-script"></a><span data-ttu-id="19c3e-372">スクリプトの先頭で事前チェックを実行する</span><span class="sxs-lookup"><span data-stu-id="19c3e-372">Do pre-checks at the beginning of the script</span></span>

<span data-ttu-id="19c3e-373">ベスト プラクティスとして、スクリプトを実行する前に、すべての入力が Excel ファイルに存在する必要があります。</span><span class="sxs-lookup"><span data-stu-id="19c3e-373">As a best practice, always ensure that all your inputs are present in the Excel file prior to running your script.</span></span> <span data-ttu-id="19c3e-374">ブックに存在するオブジェクトについて、特定の前提を設定している可能性があります。</span><span class="sxs-lookup"><span data-stu-id="19c3e-374">You may have made certain assumptions about objects being present in the workbook.</span></span> <span data-ttu-id="19c3e-375">これらのオブジェクトが存在しない場合、オブジェクトまたはデータの読み取り時にスクリプトにエラーが発生する可能性があります。</span><span class="sxs-lookup"><span data-stu-id="19c3e-375">If those objects don't exist, your script may encounter an error when you read the object or its data.</span></span> <span data-ttu-id="19c3e-376">更新または処理の一部が既に完了した後に処理とエラーが途中で開始されるのではなく、スクリプトの開始時にすべての事前チェックを実行する方が良いです。</span><span class="sxs-lookup"><span data-stu-id="19c3e-376">Rather than beginning the processing and erroring in the middle after part of the updates or processing has already finished, it is better to do all pre-checks at the start of the script.</span></span>

<span data-ttu-id="19c3e-377">たとえば、次のスクリプトでは、Table1 と Table2 という名前の 2 つのテーブルが存在する必要があります。</span><span class="sxs-lookup"><span data-stu-id="19c3e-377">For example, the following script requires two tables named Table1 and Table2 to be present.</span></span> <span data-ttu-id="19c3e-378">したがって、スクリプトは存在を確認し、ステートメントと、存在しない場合は適切な `return` メッセージで終わります。</span><span class="sxs-lookup"><span data-stu-id="19c3e-378">Hence the script checks for their presence and ends with the `return` statement and an appropriate message if they are not present.</span></span>

```TypeScript
function main(workbook: ExcelScript.Workbook) {

  // Tables that should be in the workbook for the script to work:
  const TargetTableName = 'Table1';
  const SourceTableName = 'Table2';

  // Get the table objects.
  let targetTable = workbook.getTable(TargetTableName);
  let sourceTable = workbook.getTable(SourceTableName);

  if (!targetTable || !sourceTable) {
    console.log(`Required tables missing - Check that both source (${TargetTableName}) and target (${SourceTableName}) tables are present before running the script.`);
    return;
  }

  // Continue....
}
```

<span data-ttu-id="19c3e-379">入力データの存在を確認する検証が別の関数で行っている場合は、関数からステートメントを発行してスクリプトを `return` 終了することが重要 `main` です。</span><span class="sxs-lookup"><span data-stu-id="19c3e-379">If the verification to ensure the presence of input data is happening in a separate function, it's important to end the script by issuing the `return` statement from the `main` function.</span></span>

<span data-ttu-id="19c3e-380">次の例では、関数 `main` は事前チェック `inputPresent` を実行するために関数を呼び出します。</span><span class="sxs-lookup"><span data-stu-id="19c3e-380">In the following example, the `main` function calls the `inputPresent` function to do the pre-checks.</span></span> <span data-ttu-id="19c3e-381">`inputPresent` 必要なすべての入力が存在するかどうかを示すブール値 ( `true` `false` または ) を返します。</span><span class="sxs-lookup"><span data-stu-id="19c3e-381">`inputPresent` returns a boolean (`true` or `false`) indicating whether all required inputs are present or not.</span></span> <span data-ttu-id="19c3e-382">その後、スクリプトを直ちに終了するには、関数がステートメント (つまり、関数内 `main` `return` から) を `main` 発行する必要があります。</span><span class="sxs-lookup"><span data-stu-id="19c3e-382">It's then the responsibility of the `main` function to issue the `return` statement (that is, from within the `main` function) to end the script immediately.</span></span>

```TypeScript
function main(workbook: ExcelScript.Workbook) {

  // Get the table objects.
  if (!inputPresent(workbook)) {
    return;
  }

  // Continue....
}

function inputPresent( workbook: ExcelScript.Workbook): boolean {

  // Tables that should be in the workbook for the script to work:
  const TargetTableName = 'Table1';
  const SourceTableName = 'Table2';

  // Get the table objects.
  let targetTable = workbook.getTable(TargetTableName);
  let sourceTable = workbook.getTable(SourceTableName);

  if (!targetTable || !sourceTable) {
    console.log(`Required tables missing - Check that both source (${TargetTableName}) and target (${SourceTableName}) tables are present before running the script.`);
    return false;
  }
  return true;
}
```

### <a name="when-to-abort-throw-the-script"></a><span data-ttu-id="19c3e-383">スクリプトを中止する場合 ( `throw` )</span><span class="sxs-lookup"><span data-stu-id="19c3e-383">When to abort (`throw`) the script</span></span>  

<span data-ttu-id="19c3e-384">ほとんどの場合、スクリプトから ( ) を中止 `throw` する必要があります。</span><span class="sxs-lookup"><span data-stu-id="19c3e-384">For the most part, you don't need to abort (`throw`) from your script.</span></span> <span data-ttu-id="19c3e-385">これは、スクリプトが通常、問題のためにスクリプトの実行に失敗したとユーザーに通知する理由です。</span><span class="sxs-lookup"><span data-stu-id="19c3e-385">This is because the script's usually informs the user that the script failed to run due to an issue.</span></span> <span data-ttu-id="19c3e-386">ほとんどの場合、エラー メッセージと関数からのステートメントでスクリプトを終了しても `return` 十分 `main` です。</span><span class="sxs-lookup"><span data-stu-id="19c3e-386">In most case, it's sufficient to end the script with an error message and a `return` statement from the `main` function.</span></span>

<span data-ttu-id="19c3e-387">ただし、スクリプトがスクリプトの一部として実行されている場合Power Automate条件が満たされない場合は、フローを中止できます。</span><span class="sxs-lookup"><span data-stu-id="19c3e-387">However, if your script is running as part of Power Automate, you may want to abort the flow if certain conditions are not met.</span></span> <span data-ttu-id="19c3e-388">したがって、エラーが発生した場合ではなく、スクリプトを中止して以降のコード ステートメントが実行されないステートメントを発行 `return` `throw` することが重要です。</span><span class="sxs-lookup"><span data-stu-id="19c3e-388">It's therefore important to not `return` upon an error but rather issue a `throw` statement to abort the script so that any subsequent code statements don't run.</span></span>

```TypeScript
function main(workbook: ExcelScript.Workbook) {

  // Tables that should be in the workbook for the script to work:
  const TargetTableName = 'Table1';
  const SourceTableName = 'Table2';

  // Get the table objects.
  let targetTable = workbook.getTable(TargetTableName);
  let sourceTable = workbook.getTable(SourceTableName);

  if (!targetTable || !sourceTable) {
    // Abort script.
    throw `Required tables missing - Check that both source (${TargetTableName}) and target (${SourceTableName}) tables are present before running the script.`;
  }
  
```

<span data-ttu-id="19c3e-389">次のセクションで説明したように、もう 1 つのシナリオは、複数の関数 (どの呼び出しを呼び出すなど) が関係し、エラーの伝達が難しい `main` `functionX` `functionY` 場合です。</span><span class="sxs-lookup"><span data-stu-id="19c3e-389">As mentioned in the following section, another scenario is when you have several functions involved (`main` calls `functionX` which calls `functionY`, etc.) which makes it hard to propagate the error.</span></span> <span data-ttu-id="19c3e-390">メッセージを含む入れ子になった関数から中止/スローする方が、エラー メッセージを表示してエラーを返すよりも `main` `main` 簡単です。</span><span class="sxs-lookup"><span data-stu-id="19c3e-390">Aborting/throwing from the nested function with a message may be easier than returning an error all the way up to `main` and returning from `main` with an error message.</span></span>

### <a name="when-to-use-trycatch-throw-exception"></a><span data-ttu-id="19c3e-391">try.を使用する場合。catch (スロー例外)</span><span class="sxs-lookup"><span data-stu-id="19c3e-391">When to use try..catch (throw exception)</span></span>

<span data-ttu-id="19c3e-392">この [`try..catch`](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Statements/try...catch) 手法は、API 呼び出しが失敗した場合に検出し、スクリプト内のエラーを処理する方法です。</span><span class="sxs-lookup"><span data-stu-id="19c3e-392">The [`try..catch`](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Statements/try...catch) technique is a way to detect if an API call failed and handle that error in your script.</span></span> <span data-ttu-id="19c3e-393">API の戻り値を確認して、正常に完了したと確認することが重要な場合があります。</span><span class="sxs-lookup"><span data-stu-id="19c3e-393">It may be important to check the return value of an API to verify that it was completed successfully.</span></span>

<span data-ttu-id="19c3e-394">次のスニペット例を考えてみましょう。</span><span class="sxs-lookup"><span data-stu-id="19c3e-394">Consider the following example snippet.</span></span>

```TypeScript
function main(workbook: ExcelScript.Workbook) {

  // Somewhere in the script, perform a large data update.
  range.setValues(someLargeValues);

}
```

<span data-ttu-id="19c3e-395">呼 `setValues()` び出しが失敗し、スクリプトが失敗する可能性があります。</span><span class="sxs-lookup"><span data-stu-id="19c3e-395">The `setValues()` call may fail and result in the script failure.</span></span> <span data-ttu-id="19c3e-396">コードでこの条件を処理し、エラー メッセージをカスタマイズしたり、更新プログラムを小さな単位に分割したりすることができます。その場合は、API がエラーを返し、そのエラーを解釈または処理することが重要です。</span><span class="sxs-lookup"><span data-stu-id="19c3e-396">You may wish to handle this condition in your code and perhaps customize the error message or break up the update into smaller units, etc. In that case, it's important to know that the API returned an error and interpret or handle that error.</span></span>

```TypeScript
try {
    range.setValues(someLargeValues);
} catch (error) {
    console.log(`The script failed to update the values at location ____. Please inspect and run again.`);
    console.log(error);
    return; // End script (assuming this is in main function).
}

// OR...

try {
    range.setValues(someLargeValues);
} catch (error) {
    console.log(`The script failed to update the values at location ____. Trying a different approach`);
    handleUpdatesInSmallerChunks(someLargeValues);
}

// Continue...
}
```

<span data-ttu-id="19c3e-397">もう 1 つのシナリオは、main 関数が別の関数を呼び出し、次に別の関数を呼び出し (など)、気にする API 呼び出しが下位関数でダウンする場合です。</span><span class="sxs-lookup"><span data-stu-id="19c3e-397">Another scenario is when main function calls another function, which in turn calls another function (and so on..), and the API call that you care about happens down in the bottom function.</span></span> <span data-ttu-id="19c3e-398">エラーを最大まで伝達すると、実行可能または便利 `main` ではない可能性があります。</span><span class="sxs-lookup"><span data-stu-id="19c3e-398">Propagating the error all the way up to `main` may not be feasible or convenient.</span></span> <span data-ttu-id="19c3e-399">その場合、最下位関数にエラーをスローする方が便利です。</span><span class="sxs-lookup"><span data-stu-id="19c3e-399">In that case, throwing an error in the bottom function will be most convenient.</span></span>

```TypeScript

function main(workbook: ExcelScript.Workbook) {
    ...
    updateRangeInChunks(sheet.getRange("B1"), data);
    ...
}

function updateRangeInChunks(
    ...
    updateNextChunk(startCell, values, rowsPerChunk, totalRowsUpdated);
    ...
}

function updateTargetRange(
      targetCell: ExcelScript.Range,
      values: (string | boolean | number)[][]
    ) {
    const targetRange = targetCell.getResizedRange(values.length - 1, values[0].length - 1);
    console.log(`Updating the range: ${targetRange.getAddress()}`);
    try {
      targetRange.setValues(values);
    } catch (e) {
      throw `Error while updating the whole range: ${JSON.stringify(e)}`;
    }
    return;
}
```

<span data-ttu-id="19c3e-400">*警告*: ループ `try..catch` 内で使用すると、スクリプトの速度が低下します。</span><span class="sxs-lookup"><span data-stu-id="19c3e-400">*Warning*: Using `try..catch` inside of a loop will slow down your script.</span></span> <span data-ttu-id="19c3e-401">ループの内側または周囲でこれを使用しないようにします。</span><span class="sxs-lookup"><span data-stu-id="19c3e-401">Avoid using this inside of or around loops.</span></span>

## <a name="basic-performance-considerations"></a><span data-ttu-id="19c3e-402">基本的なパフォーマンスに関する考慮事項</span><span class="sxs-lookup"><span data-stu-id="19c3e-402">Basic performance considerations</span></span>

### <a name="avoid-slow-operations-in-the-loop"></a><span data-ttu-id="19c3e-403">ループ内の低速な操作を回避する</span><span class="sxs-lookup"><span data-stu-id="19c3e-403">Avoid slow operations in the loop</span></span>

<span data-ttu-id="19c3e-404">ループ ステートメント (、など) の内部/周囲で実行すると、パフォーマンス `for` `for..of` `map` `forEach` が低下する可能性があります。</span><span class="sxs-lookup"><span data-stu-id="19c3e-404">Certain operations when done inside/around the loop statements such as `for`, `for..of`, `map`, `forEach`, etc. can lead to slow performance.</span></span> <span data-ttu-id="19c3e-405">次の API カテゴリは使用しないようにします。</span><span class="sxs-lookup"><span data-stu-id="19c3e-405">Avoid the following API categories.</span></span>

* <span data-ttu-id="19c3e-406">`get*` API</span><span class="sxs-lookup"><span data-stu-id="19c3e-406">`get*` APIs</span></span>

<span data-ttu-id="19c3e-407">ループの内部で読み取るのではなく、ループの外部で必要なすべてのデータを読み取ります。</span><span class="sxs-lookup"><span data-stu-id="19c3e-407">Read all the data you need outside of the loop rather than reading it inside of the loop.</span></span> <span data-ttu-id="19c3e-408">時には、ループの内側を読み取るのを避けるのは難しい場合があります。このような場合は、ループ数が大きすぎず、または大量のデータ構造をループしないようにバッチで管理してください。</span><span class="sxs-lookup"><span data-stu-id="19c3e-408">At times, it is hard to avoid reading inside of loops; in such a case, make sure your loop counts are not too large or manage them in batches to avoid having to loop through a large data structure.</span></span>

<span data-ttu-id="19c3e-409">**注**: 扱う範囲/データが非常に大きい場合 (>100K セルなど)、読み取り/書き込みを複数のチャンクに分割するなどの高度な手法を使用する必要があります。</span><span class="sxs-lookup"><span data-stu-id="19c3e-409">**Note**: If the range/data you are dealing with is quite large (say >100K cells), you may need to use advanced techniques like breaking up your read/writes into multiple chunks.</span></span> <span data-ttu-id="19c3e-410">次のビデオは、実際には中規模のデータセットアップ用です。</span><span class="sxs-lookup"><span data-stu-id="19c3e-410">The following video is really for a small-mid sized data setup.</span></span> <span data-ttu-id="19c3e-411">大規模なデータセットについては、「高度なデータ書き [込みシナリオ」を参照してください](write-large-dataset.md)。</span><span class="sxs-lookup"><span data-stu-id="19c3e-411">For a large dataset, refer to [advanced data write scenario](write-large-dataset.md).</span></span>

<span data-ttu-id="19c3e-412">[![読み取り/書き込みの最適化ヒントを提供するビデオ](../../images/getting-started-v_perf.jpg)](https://youtu.be/lsR_GvVW3Pg "読み取り/書き込みの最適化ヒントを示すビデオ")</span><span class="sxs-lookup"><span data-stu-id="19c3e-412">[![Video providing a read-and-write optimization tip](../../images/getting-started-v_perf.jpg)](https://youtu.be/lsR_GvVW3Pg "Video showing read-and-write optimization tip")</span></span>

* <span data-ttu-id="19c3e-413">`console.log` ステートメント (次の例を参照)</span><span class="sxs-lookup"><span data-stu-id="19c3e-413">`console.log` statement (see the following example)</span></span>

```TypeScript
// Color each cell with random color.
for (let row = 0; row < rows; row++) {
    for (let col = 0; col < cols; col++) {
        range
            .getCell(row, col)
            .getFormat()
            .getFill()
            .setColor(`#${Math.random().toString(16).substr(-6)}`);
        /* Avoid such console.log inside loop */
        // console.log("Updating" + range.getCell(row, col).getAddress());
    }
}
```

* <span data-ttu-id="19c3e-414">`try {} catch ()` ステートメント</span><span class="sxs-lookup"><span data-stu-id="19c3e-414">`try {} catch ()` statement</span></span>

<span data-ttu-id="19c3e-415">例外処理ループを `for` 避ける。</span><span class="sxs-lookup"><span data-stu-id="19c3e-415">Avoid exception handling `for` loops.</span></span> <span data-ttu-id="19c3e-416">内部ループと外部ループの両方。</span><span class="sxs-lookup"><span data-stu-id="19c3e-416">Both inside and outside loops.</span></span>

## <a name="note-to-vba-developers"></a><span data-ttu-id="19c3e-417">VBA 開発者への注意</span><span class="sxs-lookup"><span data-stu-id="19c3e-417">Note to VBA developers</span></span>

<span data-ttu-id="19c3e-418">TypeScript 言語は、VBA の構文と名前付け規則の両方と異なります。</span><span class="sxs-lookup"><span data-stu-id="19c3e-418">The TypeScript language differs from VBA both syntactically as well as in naming conventions.</span></span>

<span data-ttu-id="19c3e-419">次の同等のスニペットを確認してください。</span><span class="sxs-lookup"><span data-stu-id="19c3e-419">Check out the following equivalent snippets.</span></span>

```vba
Worksheets("Sheet1").Range("A1:G37").Clear
```

```TypeScript
workbook.getWorksheet('Sheet1').getRange('A1:G37').clear(ExcelScript.ClearApplyTo.all);
```

<span data-ttu-id="19c3e-420">TypeScript について呼び出す必要があります。</span><span class="sxs-lookup"><span data-stu-id="19c3e-420">A few things to call out about TypeScript:</span></span>

* <span data-ttu-id="19c3e-421">すべてのメソッドを実行するには、開いているかっこが必要な場合があります。</span><span class="sxs-lookup"><span data-stu-id="19c3e-421">You may notice that all methods need to have open-close parentheses to execute.</span></span> <span data-ttu-id="19c3e-422">引数は同じ方法で渡されますが、一部の引数は実行に必要な場合があります (必須と省略可能)。</span><span class="sxs-lookup"><span data-stu-id="19c3e-422">Arguments are passed identically but some arguments may be required for execution (that is, required vs optional).</span></span>
* <span data-ttu-id="19c3e-423">名前付け規則は、PascalCase 規則の代わりに camelCase に従います。</span><span class="sxs-lookup"><span data-stu-id="19c3e-423">The naming convention follows camelCase instead of PascalCase convention.</span></span>
* <span data-ttu-id="19c3e-424">メソッドは通常、オブジェクト メンバーの読み取りまたは書き込みを行っているかどうかを示 `get` `set` すプレフィックスを持っています。</span><span class="sxs-lookup"><span data-stu-id="19c3e-424">Methods usually have `get` or `set` prefixes indicating whether it is reading or writing object members.</span></span>
* <span data-ttu-id="19c3e-425">コード ブロックは、オープンクローズ中かっこで定義され、識別されます `{` `}` 。</span><span class="sxs-lookup"><span data-stu-id="19c3e-425">The code blocks are defined and identified by open-close curly braces: `{` `}`.</span></span> <span data-ttu-id="19c3e-426">条件、ステートメント、ループ、関数定義などにブロック `if` `while` `for` が必要です。</span><span class="sxs-lookup"><span data-stu-id="19c3e-426">Blocks are required for `if` conditions, `while` statements, `for` loops, function definitions, etc.</span></span>
* <span data-ttu-id="19c3e-427">関数は他の関数を呼び出し、関数内で関数を定義できます。</span><span class="sxs-lookup"><span data-stu-id="19c3e-427">Functions can call other functions and you can even define functions within a function.</span></span>

<span data-ttu-id="19c3e-428">全体的に、TypeScript は異なる言語であり、その間に類似点が少ない。</span><span class="sxs-lookup"><span data-stu-id="19c3e-428">Overall, TypeScript is a different language and there are few similarities between them.</span></span> <span data-ttu-id="19c3e-429">ただし、Officeスクリプト API 自体は、VBA API と同様の用語とデータ モデル (オブジェクト モデル) 階層を使用します。</span><span class="sxs-lookup"><span data-stu-id="19c3e-429">However, the Office Scripts API themselves use similar terminology and data-model (object model) hierarchy as VBA APIs and that should help you navigate around.</span></span>
