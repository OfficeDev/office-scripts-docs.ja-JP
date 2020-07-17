---
title: パワー自動化を使用して Office スクリプトを実行する
description: Power 自動ワークフローを使用して、web 上の Excel で Office スクリプトを取得する方法について説明します。
ms.date: 07/10/2020
localization_priority: Normal
ms.openlocfilehash: bd8fea08b7a9303ad2ceace787de6457a33fb979
ms.sourcegitcommit: ebd1079c7e2695ac0e7e4c616f2439975e196875
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 07/17/2020
ms.locfileid: "45160447"
---
# <a name="run-office-scripts-with-power-automate"></a><span data-ttu-id="0e413-103">パワー自動化を使用して Office スクリプトを実行する</span><span class="sxs-lookup"><span data-stu-id="0e413-103">Run Office Scripts with Power Automate</span></span>

<span data-ttu-id="0e413-104">[Power オートメーション](https://flow.microsoft.com)を使用すると、より大きな自動化されたワークフローに Office スクリプトを追加することができます。</span><span class="sxs-lookup"><span data-stu-id="0e413-104">[Power Automate](https://flow.microsoft.com) lets you add Office Scripts to a larger, automated workflow.</span></span> <span data-ttu-id="0e413-105">Power オートメーションでは、ワークシートのテーブルに電子メールの内容を追加したり、ブックのコメントに基づいてプロジェクト管理ツールでアクションを作成したりするなどの操作を実行できます。</span><span class="sxs-lookup"><span data-stu-id="0e413-105">You can use Power Automate do things like add the contents of an email to a worksheet's table or create actions in your project management tools based on workbook comments.</span></span>

## <a name="getting-started"></a><span data-ttu-id="0e413-106">はじめに</span><span class="sxs-lookup"><span data-stu-id="0e413-106">Getting started</span></span>

<span data-ttu-id="0e413-107">電力を自動自動化することが初めての場合は、「 [Power オートメーションの使用を開始](/power-automate/getting-started)する」を参照することをお勧めします。</span><span class="sxs-lookup"><span data-stu-id="0e413-107">If you are new to Power Automate, we recommend visiting [Get started with Power Automate](/power-automate/getting-started).</span></span> <span data-ttu-id="0e413-108">ここでは、使用可能な自動化のすべての機能について詳しく知ることができます。</span><span class="sxs-lookup"><span data-stu-id="0e413-108">There, you can learn more about all the automation possibilities available to you.</span></span> <span data-ttu-id="0e413-109">ここでは、Office スクリプトが電力自動化とどのように機能するか、および Excel の操作を改善する方法に重点を置いてドキュメントを作成します。</span><span class="sxs-lookup"><span data-stu-id="0e413-109">The documents here focus on how Office Scripts work with Power Automate and how that can help improve your Excel experience.</span></span>

<span data-ttu-id="0e413-110">Power オートメーションと Office のスクリプトの組み合わせを開始するには、チュートリアルの次の手順を実行し[て、Power 自動化を使用したスクリプトの使用を開始](../tutorials/excel-power-automate-manual.md)します。</span><span class="sxs-lookup"><span data-stu-id="0e413-110">To begin combining Power Automate and Office Scripts, follow the tutorial [Start using scripts with Power Automate](../tutorials/excel-power-automate-manual.md).</span></span> <span data-ttu-id="0e413-111">これにより、簡単なスクリプトを呼び出すフローを作成する方法を学習できます。</span><span class="sxs-lookup"><span data-stu-id="0e413-111">This will teach you how to create a flow that calls a simple script.</span></span> <span data-ttu-id="0e413-112">そのチュートリアルを完了した後、[自動実行電源自動化フローチュートリアルで [スクリプトにデータを渡す](../tutorials/excel-power-automate-trigger.md)] を参照してください。 Office スクリプトを power オートメーションフローに接続する方法について詳しくは、こちらを参照してください。</span><span class="sxs-lookup"><span data-stu-id="0e413-112">After you've completed that tutorial and the [Pass data to scripts in an automatically-run Power Automate flow](../tutorials/excel-power-automate-trigger.md) tutorial, return here for detailed information about connecting Office Scripts to Power Automate flows.</span></span>

## <a name="excel-online-business-connector"></a><span data-ttu-id="0e413-113">Excel Online (Business) コネクタ</span><span class="sxs-lookup"><span data-stu-id="0e413-113">Excel Online (Business) connector</span></span>

<span data-ttu-id="0e413-114">[コネクタ](/connectors/connectors)は、電力の自動化とアプリケーションの間のブリッジです。</span><span class="sxs-lookup"><span data-stu-id="0e413-114">[Connectors](/connectors/connectors) are the bridges between Power Automate and applications.</span></span> <span data-ttu-id="0e413-115">[Excel Online (Business) コネクタ](/connectors/excelonlinebusiness)を使用すると、excel ブックへのアクセスがフローに付与されます。</span><span class="sxs-lookup"><span data-stu-id="0e413-115">The [Excel Online (Business) connector](/connectors/excelonlinebusiness) gives your flows access to Excel workbooks.</span></span> <span data-ttu-id="0e413-116">"スクリプトを実行する" アクションを使用すると、選択したブックからアクセス可能な Office スクリプトを呼び出すことができます。</span><span class="sxs-lookup"><span data-stu-id="0e413-116">The "Run script" action lets you call any Office Script accessible through the selected workbook.</span></span> <span data-ttu-id="0e413-117">フローを使用してスクリプトを実行できるだけでなく、スクリプト内のフローを使用して、ブックとの間でデータをやり取りすることもできます。</span><span class="sxs-lookup"><span data-stu-id="0e413-117">Not only can you run scripts through a flow, you can pass data to and from the workbook with the flow through the scripts.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="0e413-118">"スクリプトを実行する" アクションを実行すると、Excel コネクタを使用するユーザーに、ブックとそのデータに対して重要なアクセス権が与えられます。</span><span class="sxs-lookup"><span data-stu-id="0e413-118">The "Run script" action gives people who use the Excel connector significant access to your workbook and its data.</span></span> <span data-ttu-id="0e413-119">また、外部の[呼び出しからの外部呼び出し](external-calls.md)について説明するように、外部 API を呼び出すスクリプトにはセキュリティリスクがあります。</span><span class="sxs-lookup"><span data-stu-id="0e413-119">Additionally, there are security risks with scripts that make external API calls, as explained in [External calls from Power Automate](external-calls.md).</span></span> <span data-ttu-id="0e413-120">管理者が非常に機密性の高いデータの公開を懸念している場合は、Excel Online コネクタをオフにするか、 [Office スクリプト管理者コントロール](https://support.microsoft.com/office/19d3c51a-6ca2-40ab-978d-60fa49554dcf)を使用して office スクリプトへのアクセスを制限することができます。</span><span class="sxs-lookup"><span data-stu-id="0e413-120">If your admin is concerned with the exposure of highly sensitive data, they can either turn off the Excel Online connector or restrict access to Office Scripts through the [Office Scripts administrator controls](https://support.microsoft.com/office/19d3c51a-6ca2-40ab-978d-60fa49554dcf).</span></span>

## <a name="data-transfer-in-flows-for-scripts"></a><span data-ttu-id="0e413-121">スクリプトのフローでのデータ転送</span><span class="sxs-lookup"><span data-stu-id="0e413-121">Data transfer in flows for scripts</span></span>

<span data-ttu-id="0e413-122">電源自動化を使用すると、フローの手順間でデータを渡すことができます。</span><span class="sxs-lookup"><span data-stu-id="0e413-122">Power Automate lets you pass pieces of data between steps of your flow.</span></span> <span data-ttu-id="0e413-123">必要な種類の情報を受け入れるようにスクリプトを構成して、フローに必要なブックから任意のものを返すことができます。</span><span class="sxs-lookup"><span data-stu-id="0e413-123">Scripts can be configured to accept whatever types of information you need and return anything from your workbook that you want in your flow.</span></span> <span data-ttu-id="0e413-124">スクリプトへの入力は、関数にパラメーターを追加することによって指定され `main` ます (に加えて `workbook: ExcelScript.Workbook` )。</span><span class="sxs-lookup"><span data-stu-id="0e413-124">Input for your script is specified by adding parameters to the `main` function (in addition to `workbook: ExcelScript.Workbook`).</span></span> <span data-ttu-id="0e413-125">スクリプトからの出力は、戻り値の型をに追加することによって宣言され `main` ます。</span><span class="sxs-lookup"><span data-stu-id="0e413-125">Output from the script is declared by adding a return type to `main`.</span></span>

> [!NOTE]
> <span data-ttu-id="0e413-126">フローに "実行スクリプト" ブロックを作成すると、受け入れられるパラメーターと返される型が設定されます。</span><span class="sxs-lookup"><span data-stu-id="0e413-126">When you create a "Run Script" block in your flow, the accepted parameters and returned types are populated.</span></span> <span data-ttu-id="0e413-127">スクリプトのパラメーターまたは戻り値の型を変更する場合は、フローの "Run script" ブロックをやり直す必要があります。</span><span class="sxs-lookup"><span data-stu-id="0e413-127">If you change the parameters or return types of your script, you'll need to redo the "Run script" block of your flow.</span></span> <span data-ttu-id="0e413-128">これにより、データが正しく解析されるようになります。</span><span class="sxs-lookup"><span data-stu-id="0e413-128">This ensures the data is being parsed correctly.</span></span>

<span data-ttu-id="0e413-129">次のセクションでは、電力の自動化に使用されるスクリプトの入力と出力の詳細について説明します。</span><span class="sxs-lookup"><span data-stu-id="0e413-129">The following sections cover the details of input and output for scripts used in Power Automate.</span></span> <span data-ttu-id="0e413-130">このトピックを学習するための実践的なアプローチを希望される場合は、「自動[実行パワー自動フローのチュートリアルで、スクリプトにデータを渡す」](../tutorials/excel-power-automate-trigger.md)をお試しください。または、[自動タスクリマインダー](../resources/scenarios/task-reminders.md)サンプルシナリオを参照してください。</span><span class="sxs-lookup"><span data-stu-id="0e413-130">If you'd like a hands-on approach to learning this topic, try out the [Pass data to scripts in an automatically-run Power Automate flow](../tutorials/excel-power-automate-trigger.md) tutorial or explore the [Automated task reminders](../resources/scenarios/task-reminders.md) sample scenario.</span></span>

### <a name="main-parameters-passing-data-to-a-script"></a><span data-ttu-id="0e413-131">`main`パラメーター: スクリプトにデータを渡す</span><span class="sxs-lookup"><span data-stu-id="0e413-131">`main` Parameters: Passing data to a script</span></span>

<span data-ttu-id="0e413-132">すべてのスクリプトの入力は、関数の追加パラメーターとして指定され `main` ます。</span><span class="sxs-lookup"><span data-stu-id="0e413-132">All script input is specified as additional parameters for the `main` function.</span></span> <span data-ttu-id="0e413-133">たとえば、入力として名前を表すを受け入れるスクリプトが必要な場合は、 `string` `main` 署名をに変更し `function main(workbook: ExcelScript.Workbook, name: string)` ます。</span><span class="sxs-lookup"><span data-stu-id="0e413-133">For example, if you wanted a script to accept a `string` that represents a name as input, you would change the `main` signature to `function main(workbook: ExcelScript.Workbook, name: string)`.</span></span>

<span data-ttu-id="0e413-134">Power 自動化でフローを構成するときは、スクリプトの入力を静的な値、[式](/power-automate/use-expressions-in-conditions)、または動的コンテンツとして指定できます。</span><span class="sxs-lookup"><span data-stu-id="0e413-134">When you're configuring a flow in Power Automate, you can specify script input as static values, [expressions](/power-automate/use-expressions-in-conditions), or dynamic content.</span></span> <span data-ttu-id="0e413-135">個々のサービスのコネクタの詳細については、「[電源自動化コネクタ](/connectors/)」のドキュメントを参照してください。</span><span class="sxs-lookup"><span data-stu-id="0e413-135">Details on an individual service's connector can be found in the [Power Automate Connector documentation](/connectors/).</span></span>

<span data-ttu-id="0e413-136">入力パラメーターをスクリプトの関数に追加するときは `main` 、次の制限と制限事項を考慮してください。</span><span class="sxs-lookup"><span data-stu-id="0e413-136">When adding input parameters to a script's `main` function, consider the following allowances and restrictions.</span></span>

1. <span data-ttu-id="0e413-137">最初のパラメーターの型はでなければなりません `ExcelScript.Workbook` 。</span><span class="sxs-lookup"><span data-stu-id="0e413-137">The first parameter must be of type `ExcelScript.Workbook`.</span></span> <span data-ttu-id="0e413-138">そのパラメーター名は重要ではありません。</span><span class="sxs-lookup"><span data-stu-id="0e413-138">Its parameter name does not matter.</span></span>

2. <span data-ttu-id="0e413-139">すべてのパラメーターには、型を指定する必要があります。</span><span class="sxs-lookup"><span data-stu-id="0e413-139">Every parameter must have a type.</span></span>

3. <span data-ttu-id="0e413-140">基本的な型、、、、、、 `string` `number` `boolean` `any` `unknown` `object` 、 `undefined` がサポートされています。</span><span class="sxs-lookup"><span data-stu-id="0e413-140">The basic types `string`, `number`, `boolean`, `any`, `unknown`, `object`, and `undefined` are supported.</span></span>

4. <span data-ttu-id="0e413-141">前にリストされていた基本的な種類の配列がサポートされています。</span><span class="sxs-lookup"><span data-stu-id="0e413-141">Arrays of the previously listed basic types are supported.</span></span>

5. <span data-ttu-id="0e413-142">入れ子になった配列は、パラメーターとしてサポートされます (戻り値の型としてではありません)。</span><span class="sxs-lookup"><span data-stu-id="0e413-142">Nested arrays are supported as parameters (but not as return types).</span></span>

6. <span data-ttu-id="0e413-143">共用体型は、1つの型 ( `string` 、 `number` 、または) に属するリテラルの和集合である場合に使用でき `boolean` ます。</span><span class="sxs-lookup"><span data-stu-id="0e413-143">Union types are allowed if they are a union of literals belonging to a single type (`string`, `number`, or `boolean`).</span></span> <span data-ttu-id="0e413-144">サポートされている型と未定義の共用体もサポートされています。</span><span class="sxs-lookup"><span data-stu-id="0e413-144">Unions of a supported type with undefined are also supported.</span></span>

7. <span data-ttu-id="0e413-145">オブジェクトの種類は、型 `string` 、 `number` 、、 `boolean` サポートされている配列、またはその他のサポートされているオブジェクトのプロパティが含まれている場合に許可されます。</span><span class="sxs-lookup"><span data-stu-id="0e413-145">Object types are allowed if they contain properties of type `string`, `number`, `boolean`, supported arrays, or other supported objects.</span></span> <span data-ttu-id="0e413-146">次の例は、パラメータタイプとしてサポートされているネストされたオブジェクトを示しています。</span><span class="sxs-lookup"><span data-stu-id="0e413-146">The following example shows nested objects that are supported as parameter types:</span></span>

    ```TypeScript
    // Office Scripts can return an Employee object because Position only contains strings and numbers.
    interface Employee {
        name: string;
        job: Position;
    }

    interface Position {
        id: number;
        title: string;
    }
    ```

8. <span data-ttu-id="0e413-147">オブジェクトのインターフェイスまたはクラス定義は、スクリプトで定義されている必要があります。</span><span class="sxs-lookup"><span data-stu-id="0e413-147">Objects must have their interface or class definition defined in the script.</span></span> <span data-ttu-id="0e413-148">また、次の例に示すように、オブジェクトを匿名でインラインで定義することもできます。</span><span class="sxs-lookup"><span data-stu-id="0e413-148">An object can also be defined anonymously inline, as in the following example:</span></span>

    ```TypeScript
    function main(workbook: ExcelScript.Workbook): {name: string, email: string}
    ```

9. <span data-ttu-id="0e413-149">省略可能なパラメーターを指定できます。オプションの修飾子 (たとえば、) を使用することもでき `?` `function main(workbook: ExcelScript.Workbook, Name?: string)` ます。</span><span class="sxs-lookup"><span data-stu-id="0e413-149">Optional parameters are allowed and can be denoted as such by using the optional modifier `?` (for example, `function main(workbook: ExcelScript.Workbook, Name?: string)`).</span></span>

10. <span data-ttu-id="0e413-150">既定のパラメーター値を使用できます (例 `async function main(workbook: ExcelScript.Workbook, Name: string = 'Jane Doe')` :</span><span class="sxs-lookup"><span data-stu-id="0e413-150">Default parameter values are allowed (for example `async function main(workbook: ExcelScript.Workbook, Name: string = 'Jane Doe')`.</span></span>

### <a name="returning-data-from-a-script"></a><span data-ttu-id="0e413-151">スクリプトからデータを返す</span><span class="sxs-lookup"><span data-stu-id="0e413-151">Returning data from a script</span></span>

<span data-ttu-id="0e413-152">スクリプトは、Power オートメーションフローで動的コンテンツとして使用するブックからのデータを返すことができます。</span><span class="sxs-lookup"><span data-stu-id="0e413-152">Scripts can return data from the workbook to be used as dynamic content in a Power Automate flow.</span></span> <span data-ttu-id="0e413-153">入力パラメーターと同様に、Power オートメーションでは戻り値の型にいくつかの制限が課されます。</span><span class="sxs-lookup"><span data-stu-id="0e413-153">As with input parameters, Power Automate places some restrictions on the return type.</span></span>

1. <span data-ttu-id="0e413-154">基本的な型、、、、、 `string` `number` がサポートされてい `boolean` `void` `undefined` ます。</span><span class="sxs-lookup"><span data-stu-id="0e413-154">The basic types `string`, `number`, `boolean`, `void`, and `undefined` are supported.</span></span>

2. <span data-ttu-id="0e413-155">戻り値の型として使用される共用体型は、スクリプトパラメーターとして使用する場合と同じ制限に従います。</span><span class="sxs-lookup"><span data-stu-id="0e413-155">Union types used as return types follow the same restrictions as they do when used as script parameters.</span></span>

3. <span data-ttu-id="0e413-156">配列型は `string` 、型、、またはのいずれかである場合に使用でき `number` `boolean` ます。</span><span class="sxs-lookup"><span data-stu-id="0e413-156">Array types are allowed if they are of type `string`, `number`, or `boolean`.</span></span> <span data-ttu-id="0e413-157">また、型がサポートされている共用体型またはサポートされているリテラル型の場合にも使用できます。</span><span class="sxs-lookup"><span data-stu-id="0e413-157">They are also allowed if the type is a supported union or supported literal type.</span></span>

4. <span data-ttu-id="0e413-158">戻り値の型として使用されるオブジェクトの種類は、スクリプトパラメーターとして使用する場合と同じ制限に従います。</span><span class="sxs-lookup"><span data-stu-id="0e413-158">Object types used as return types follow the same restrictions as they do when used as script parameters.</span></span>

5. <span data-ttu-id="0e413-159">暗黙的な入力はサポートされていますが、定義された型と同じルールに従う必要があります。</span><span class="sxs-lookup"><span data-stu-id="0e413-159">Implicit typing is supported, though it must follow the same rules as a defined type.</span></span>

## <a name="avoid-using-relative-references"></a><span data-ttu-id="0e413-160">相対参照の使用を避ける</span><span class="sxs-lookup"><span data-stu-id="0e413-160">Avoid using relative references</span></span>

<span data-ttu-id="0e413-161">Power オートメーションは、ユーザーの代わりに、選択した Excel ブックでスクリプトを実行します。</span><span class="sxs-lookup"><span data-stu-id="0e413-161">Power Automate runs your script in the chosen Excel workbook on your behalf.</span></span> <span data-ttu-id="0e413-162">これが発生すると、ブックが閉じられる場合があります。</span><span class="sxs-lookup"><span data-stu-id="0e413-162">The workbook might be closed when this happens.</span></span> <span data-ttu-id="0e413-163">など、ユーザーの現在の状態に依存する API は、 `Workbook.getActiveWorksheet` 電力の自動処理によって実行されると失敗します。</span><span class="sxs-lookup"><span data-stu-id="0e413-163">Any API that relies on the user's current state, such as `Workbook.getActiveWorksheet`, will fail when run through Power Automate.</span></span> <span data-ttu-id="0e413-164">スクリプトを設計するときは、必ずワークシートおよび範囲の絶対参照を使用してください。</span><span class="sxs-lookup"><span data-stu-id="0e413-164">When designing your scripts, be sure to use absolute references for worksheets and ranges.</span></span>

<span data-ttu-id="0e413-165">次のメソッドは、Power オートメーションフローでスクリプトから呼び出されたときにエラーをスローして失敗します。</span><span class="sxs-lookup"><span data-stu-id="0e413-165">The following methods will throw an error and fail when called from a script in a Power Automate flow.</span></span>

| <span data-ttu-id="0e413-166">クラス</span><span class="sxs-lookup"><span data-stu-id="0e413-166">Class</span></span> | <span data-ttu-id="0e413-167">メソッド</span><span class="sxs-lookup"><span data-stu-id="0e413-167">Method</span></span> |
|--|--|
| [<span data-ttu-id="0e413-168">グラフ</span><span class="sxs-lookup"><span data-stu-id="0e413-168">Chart</span></span>](/javascript/api/office-scripts/excelscript/excelscript.chart) | `activate` |
| [<span data-ttu-id="0e413-169">Range</span><span class="sxs-lookup"><span data-stu-id="0e413-169">Range</span></span>](/javascript/api/office-scripts/excelscript/excelscript.range) | `select` |
| [<span data-ttu-id="0e413-170">ブック</span><span class="sxs-lookup"><span data-stu-id="0e413-170">Workbook</span></span>](/javascript/api/office-scripts/excelscript/excelscript.workbook) | `getActiveCell` |
| [<span data-ttu-id="0e413-171">ブック</span><span class="sxs-lookup"><span data-stu-id="0e413-171">Workbook</span></span>](/javascript/api/office-scripts/excelscript/excelscript.workbook) | `getActiveChart` |
| [<span data-ttu-id="0e413-172">ブック</span><span class="sxs-lookup"><span data-stu-id="0e413-172">Workbook</span></span>](/javascript/api/office-scripts/excelscript/excelscript.workbook) | `getActiveSlicer` |
| [<span data-ttu-id="0e413-173">ブック</span><span class="sxs-lookup"><span data-stu-id="0e413-173">Workbook</span></span>](/javascript/api/office-scripts/excelscript/excelscript.workbook) | `getActiveWorksheet` |
| [<span data-ttu-id="0e413-174">ブック</span><span class="sxs-lookup"><span data-stu-id="0e413-174">Workbook</span></span>](/javascript/api/office-scripts/excelscript/excelscript.workbook) | `getSelectedRange` |
| [<span data-ttu-id="0e413-175">ブック</span><span class="sxs-lookup"><span data-stu-id="0e413-175">Workbook</span></span>](/javascript/api/office-scripts/excelscript/excelscript.workbook) | `getSelectedRanges` |
| [<span data-ttu-id="0e413-176">ワークシート</span><span class="sxs-lookup"><span data-stu-id="0e413-176">Worksheet</span></span>](/javascript/api/office-scripts/excelscript/excelscript.workbook) | `activate` |

## <a name="example"></a><span data-ttu-id="0e413-177">例</span><span class="sxs-lookup"><span data-stu-id="0e413-177">Example</span></span>

<span data-ttu-id="0e413-178">次のスクリーンショットは、 [GitHub](https://github.com/)の問題がユーザーに割り当てられたときにトリガーされる電源自動化フローを示しています。</span><span class="sxs-lookup"><span data-stu-id="0e413-178">The following screenshot shows a Power Automate flow that's triggered whenever a [GitHub](https://github.com/) issue is assigned to you.</span></span> <span data-ttu-id="0e413-179">このフローは、Excel ブックのテーブルに問題を追加するスクリプトを実行します。</span><span class="sxs-lookup"><span data-stu-id="0e413-179">The flow runs a script that adds the issue to a table in an Excel workbook.</span></span> <span data-ttu-id="0e413-180">そのテーブルに5つ以上の問題がある場合、フローはメール事前通知を送信します。</span><span class="sxs-lookup"><span data-stu-id="0e413-180">If there are five or more issues in that table, the flow sends an email reminder.</span></span>

![電源自動化フローエディターに示されている例のフロー。](../images/power-automate-parameter-return-sample.png)

<span data-ttu-id="0e413-182">`main`スクリプトの関数は、[案件 ID] と [issue title] を入力パラメーターとして指定し、スクリプトは issue テーブル内の行数を返します。</span><span class="sxs-lookup"><span data-stu-id="0e413-182">The `main` function of the script specifies the issue ID and issue title as input parameters, and the script returns the number of rows in the issue table.</span></span>

```TypeScript
function main(
  workbook: ExcelScript.Workbook,
  issueId: string,
  issueTitle: string): number {
  // Get the "GitHub" worksheet.
  let worksheet = workbook.getWorksheet("GitHub");

  // Get the first table in this worksheet, which contains the table of GitHub issues.
  let issueTable = worksheet.getTables()[0];

  // Add the issue ID and issue title as a row.
  issueTable.addRow(-1, [issueId, issueTitle]);

  // Return the number of rows in the table, which represents how many issues are assigned to this user.
  return issueTable.getRangeBetweenHeaderAndTotal().getRowCount();
}
```

## <a name="see-also"></a><span data-ttu-id="0e413-183">関連項目</span><span class="sxs-lookup"><span data-stu-id="0e413-183">See also</span></span>

- [<span data-ttu-id="0e413-184">Power オートメーションを使用して web 上の Excel で Office スクリプトを実行する</span><span class="sxs-lookup"><span data-stu-id="0e413-184">Run Office Scripts in Excel on the web with Power Automate</span></span>](../tutorials/excel-power-automate-manual.md)
- [<span data-ttu-id="0e413-185">自動実行電源自動化フローで、スクリプトにデータを渡す</span><span class="sxs-lookup"><span data-stu-id="0e413-185">Pass data to scripts in an automatically-run Power Automate flow</span></span>](../tutorials/excel-power-automate-trigger.md)
- [<span data-ttu-id="0e413-186">Excel on the web での Office スクリプトのスクリプトの基本事項</span><span class="sxs-lookup"><span data-stu-id="0e413-186">Scripting fundamentals for Office Scripts in Excel on the web</span></span>](scripting-fundamentals.md)
- [<span data-ttu-id="0e413-187">Power Automate の使用を開始する</span><span class="sxs-lookup"><span data-stu-id="0e413-187">Get started with Power Automate</span></span>](/power-automate/getting-started)
- [<span data-ttu-id="0e413-188">Excel Online (ビジネス向け) コネクタのリファレンスドキュメント</span><span class="sxs-lookup"><span data-stu-id="0e413-188">Excel Online (Business) connector reference documentation</span></span>](/connectors/excelonlinebusiness/)
