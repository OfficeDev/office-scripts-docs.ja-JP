---
title: Power Automate Officeスクリプトを実行する
description: Power Automate ワークフロー Office操作する Web 上の Excel スクリプトを取得する方法。
ms.date: 12/16/2020
localization_priority: Normal
ms.openlocfilehash: 1ca9aa14efe7cf2c91100a32fbc9a69054012f06
ms.sourcegitcommit: 45ffe3dbd2c834b78592ad35928cf8096f5e80bc
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 04/14/2021
ms.locfileid: "51755071"
---
# <a name="run-office-scripts-with-power-automate"></a><span data-ttu-id="f8cef-103">Power Automate Officeスクリプトを実行する</span><span class="sxs-lookup"><span data-stu-id="f8cef-103">Run Office Scripts with Power Automate</span></span>

<span data-ttu-id="f8cef-104">[Power Automate を](https://flow.microsoft.com) 使用すると、Officeスクリプトを大規模で自動化されたワークフローに追加できます。</span><span class="sxs-lookup"><span data-stu-id="f8cef-104">[Power Automate](https://flow.microsoft.com) lets you add Office Scripts to a larger, automated workflow.</span></span> <span data-ttu-id="f8cef-105">Power Automate を使用すると、ワークシートのテーブルに電子メールの内容を追加したり、ブックのコメントに基づいてプロジェクト管理ツールでアクションを作成したりすることができます。</span><span class="sxs-lookup"><span data-stu-id="f8cef-105">You can use Power Automate do things like add the contents of an email to a worksheet's table or create actions in your project management tools based on workbook comments.</span></span>

## <a name="getting-started"></a><span data-ttu-id="f8cef-106">はじめに</span><span class="sxs-lookup"><span data-stu-id="f8cef-106">Getting started</span></span>

<span data-ttu-id="f8cef-107">Power Automate を使用する場合は、「Power Automate の使用を開始 [する」にアクセスすることをお勧めします](/power-automate/getting-started)。</span><span class="sxs-lookup"><span data-stu-id="f8cef-107">If you are new to Power Automate, we recommend visiting [Get started with Power Automate](/power-automate/getting-started).</span></span> <span data-ttu-id="f8cef-108">そこで、利用可能なすべてのオートメーションの可能性について詳しくは、ご覧ください。</span><span class="sxs-lookup"><span data-stu-id="f8cef-108">There, you can learn more about all the automation possibilities available to you.</span></span> <span data-ttu-id="f8cef-109">このドキュメントでは、Power Automate Officeスクリプトがどのように動作し、Excel エクスペリエンスを向上させるのに役立つのかについて重点的に取り上っています。</span><span class="sxs-lookup"><span data-stu-id="f8cef-109">The documents here focus on how Office Scripts work with Power Automate and how that can help improve your Excel experience.</span></span>

<span data-ttu-id="f8cef-110">Power Automate スクリプトと Officeスクリプトの組み合Office、Power Automate を使用したスクリプトの [使用を開始するチュートリアルに従います](../tutorials/excel-power-automate-manual.md)。</span><span class="sxs-lookup"><span data-stu-id="f8cef-110">To begin combining Power Automate and Office Scripts, follow the tutorial [Start using scripts with Power Automate](../tutorials/excel-power-automate-manual.md).</span></span> <span data-ttu-id="f8cef-111">これにより、単純なスクリプトを呼び出すフローを作成する方法を説明します。</span><span class="sxs-lookup"><span data-stu-id="f8cef-111">This will teach you how to create a flow that calls a simple script.</span></span> <span data-ttu-id="f8cef-112">このチュートリアルと、自動的に実行される [Power Automate](../tutorials/excel-power-automate-trigger.md) フロー チュートリアルのスクリプトへのデータの渡しを完了したら、Office スクリプトを Power Automate フローに接続する方法の詳細については、ここに戻します。</span><span class="sxs-lookup"><span data-stu-id="f8cef-112">After you've completed that tutorial and the [Pass data to scripts in an automatically-run Power Automate flow](../tutorials/excel-power-automate-trigger.md) tutorial, return here for detailed information about connecting Office Scripts to Power Automate flows.</span></span>

## <a name="excel-online-business-connector"></a><span data-ttu-id="f8cef-113">Excel Online (Business) コネクタ</span><span class="sxs-lookup"><span data-stu-id="f8cef-113">Excel Online (Business) connector</span></span>

<span data-ttu-id="f8cef-114">[コネクタは](/connectors/connectors) 、Power Automate とアプリケーションの間のブリッジです。</span><span class="sxs-lookup"><span data-stu-id="f8cef-114">[Connectors](/connectors/connectors) are the bridges between Power Automate and applications.</span></span> <span data-ttu-id="f8cef-115">[Excel Online (Business) コネクタを使用すると](/connectors/excelonlinebusiness)、フローから Excel ブックにアクセスできます。</span><span class="sxs-lookup"><span data-stu-id="f8cef-115">The [Excel Online (Business) connector](/connectors/excelonlinebusiness) gives your flows access to Excel workbooks.</span></span> <span data-ttu-id="f8cef-116">"スクリプトの実行" アクションでは、選択したブックからアクセスOfficeスクリプトを呼び出します。</span><span class="sxs-lookup"><span data-stu-id="f8cef-116">The "Run script" action lets you call any Office Script accessible through the selected workbook.</span></span> <span data-ttu-id="f8cef-117">また、フローによってデータを提供できるようスクリプトにパラメーターを入力したり、フロー内の後の手順に関する情報をスクリプトから返したりすることもできます。</span><span class="sxs-lookup"><span data-stu-id="f8cef-117">You can also give your scripts input parameters so data can be provided by the flow, or have your script return information for later steps in the flow.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="f8cef-118">"スクリプトの実行" アクションにより、Excel コネクタを使用するユーザーはブックとそのデータに重要なアクセス権を与えます。</span><span class="sxs-lookup"><span data-stu-id="f8cef-118">The "Run script" action gives people who use the Excel connector significant access to your workbook and its data.</span></span> <span data-ttu-id="f8cef-119">さらに [、「Power Automate](external-calls.md)からの外部呼び出し」で説明したように、外部 API 呼び出しを行うスクリプトにはセキュリティリスクがあります。</span><span class="sxs-lookup"><span data-stu-id="f8cef-119">Additionally, there are security risks with scripts that make external API calls, as explained in [External calls from Power Automate](external-calls.md).</span></span> <span data-ttu-id="f8cef-120">管理者が機密性の高いデータの露出に関心がある場合は、Excel Online コネクタをオフにするか、Office スクリプト管理者コントロールを使用して Office スクリプトへのアクセス [を制限できます](/microsoft-365/admin/manage/manage-office-scripts-settings)。</span><span class="sxs-lookup"><span data-stu-id="f8cef-120">If your admin is concerned with the exposure of highly sensitive data, they can either turn off the Excel Online connector or restrict access to Office Scripts through the [Office Scripts administrator controls](/microsoft-365/admin/manage/manage-office-scripts-settings).</span></span>

## <a name="data-transfer-in-flows-for-scripts"></a><span data-ttu-id="f8cef-121">スクリプトのフローでのデータ転送</span><span class="sxs-lookup"><span data-stu-id="f8cef-121">Data transfer in flows for scripts</span></span>

<span data-ttu-id="f8cef-122">Power Automate を使用すると、フローのステップ間でデータを渡します。</span><span class="sxs-lookup"><span data-stu-id="f8cef-122">Power Automate lets you pass pieces of data between steps of your flow.</span></span> <span data-ttu-id="f8cef-123">必要な情報の種類を受け入れ、フローで必要な情報をブックから返すスクリプトを構成できます。</span><span class="sxs-lookup"><span data-stu-id="f8cef-123">Scripts can be configured to accept whatever types of information you need and return anything from your workbook that you want in your flow.</span></span> <span data-ttu-id="f8cef-124">スクリプトの入力は、(に加えて) 関数にパラメーター `main` を追加することで指定されます `workbook: ExcelScript.Workbook` 。</span><span class="sxs-lookup"><span data-stu-id="f8cef-124">Input for your script is specified by adding parameters to the `main` function (in addition to `workbook: ExcelScript.Workbook`).</span></span> <span data-ttu-id="f8cef-125">スクリプトからの出力は、 に戻り値の型を追加することで宣言されます `main` 。</span><span class="sxs-lookup"><span data-stu-id="f8cef-125">Output from the script is declared by adding a return type to `main`.</span></span>

> [!NOTE]
> <span data-ttu-id="f8cef-126">フローで "スクリプトの実行" ブロックを作成すると、受け入れられるパラメーターと返される型が設定されます。</span><span class="sxs-lookup"><span data-stu-id="f8cef-126">When you create a "Run Script" block in your flow, the accepted parameters and returned types are populated.</span></span> <span data-ttu-id="f8cef-127">スクリプトのパラメーターまたは戻り値の種類を変更する場合は、フローの "スクリプトの実行" ブロックをやり直す必要があります。</span><span class="sxs-lookup"><span data-stu-id="f8cef-127">If you change the parameters or return types of your script, you'll need to redo the "Run script" block of your flow.</span></span> <span data-ttu-id="f8cef-128">これにより、データが正しく解析されます。</span><span class="sxs-lookup"><span data-stu-id="f8cef-128">This ensures the data is being parsed correctly.</span></span>

<span data-ttu-id="f8cef-129">次のセクションでは、Power Automate で使用されるスクリプトの入力と出力の詳細について説明します。</span><span class="sxs-lookup"><span data-stu-id="f8cef-129">The following sections cover the details of input and output for scripts used in Power Automate.</span></span> <span data-ttu-id="f8cef-130">このトピックの学習に関する実践的なアプローチが必要な場合は、自動実行の[Power Automate](../tutorials/excel-power-automate-trigger.md)フローチュートリアルでスクリプトにデータを渡すチュートリアル[](../resources/scenarios/task-reminders.md)を試してみるか、自動タスク リマインダーのサンプル シナリオを参照してください。</span><span class="sxs-lookup"><span data-stu-id="f8cef-130">If you'd like a hands-on approach to learning this topic, try out the [Pass data to scripts in an automatically-run Power Automate flow](../tutorials/excel-power-automate-trigger.md) tutorial or explore the [Automated task reminders](../resources/scenarios/task-reminders.md) sample scenario.</span></span>

### <a name="main-parameters-passing-data-to-a-script"></a><span data-ttu-id="f8cef-131">`main` パラメーター: スクリプトにデータを渡す</span><span class="sxs-lookup"><span data-stu-id="f8cef-131">`main` Parameters: Passing data to a script</span></span>

<span data-ttu-id="f8cef-132">すべてのスクリプト入力は、関数の追加パラメーターとして指定 `main` されます。</span><span class="sxs-lookup"><span data-stu-id="f8cef-132">All script input is specified as additional parameters for the `main` function.</span></span> <span data-ttu-id="f8cef-133">たとえば、名前を入力として表すスクリプトを受け入れる場合は、署名 `string` を `main` に変更します `function main(workbook: ExcelScript.Workbook, name: string)` 。</span><span class="sxs-lookup"><span data-stu-id="f8cef-133">For example, if you wanted a script to accept a `string` that represents a name as input, you would change the `main` signature to `function main(workbook: ExcelScript.Workbook, name: string)`.</span></span>

<span data-ttu-id="f8cef-134">Power Automate でフローを構成する場合は、スクリプト入力を静的な値、式、または動的 [コンテンツとして指定](/power-automate/use-expressions-in-conditions)できます。</span><span class="sxs-lookup"><span data-stu-id="f8cef-134">When you're configuring a flow in Power Automate, you can specify script input as static values, [expressions](/power-automate/use-expressions-in-conditions), or dynamic content.</span></span> <span data-ttu-id="f8cef-135">個々のサービスのコネクタの詳細については [、「Power Automate Connector」のドキュメントを参照してください](/connectors/)。</span><span class="sxs-lookup"><span data-stu-id="f8cef-135">Details on an individual service's connector can be found in the [Power Automate Connector documentation](/connectors/).</span></span>

<span data-ttu-id="f8cef-136">スクリプトの関数に入力パラメーターを追加する場合は、次の許容値と `main` 制限を考慮してください。</span><span class="sxs-lookup"><span data-stu-id="f8cef-136">When adding input parameters to a script's `main` function, consider the following allowances and restrictions.</span></span>

1. <span data-ttu-id="f8cef-137">最初のパラメーターは型である必要があります `ExcelScript.Workbook` 。</span><span class="sxs-lookup"><span data-stu-id="f8cef-137">The first parameter must be of type `ExcelScript.Workbook`.</span></span> <span data-ttu-id="f8cef-138">パラメーター名は関係ありません。</span><span class="sxs-lookup"><span data-stu-id="f8cef-138">Its parameter name does not matter.</span></span>

2. <span data-ttu-id="f8cef-139">すべてのパラメーターには、型 (または `string` など) が必要 `number` です。</span><span class="sxs-lookup"><span data-stu-id="f8cef-139">Every parameter must have a type (such as `string` or `number`).</span></span>

3. <span data-ttu-id="f8cef-140">基本的な型 `string` `number` `boolean` `any` 、、、、 `unknown` `object` `undefined` がサポートされています。</span><span class="sxs-lookup"><span data-stu-id="f8cef-140">The basic types `string`, `number`, `boolean`, `any`, `unknown`, `object`, and `undefined` are supported.</span></span>

4. <span data-ttu-id="f8cef-141">前に示した基本型の配列がサポートされています。</span><span class="sxs-lookup"><span data-stu-id="f8cef-141">Arrays of the previously listed basic types are supported.</span></span>

5. <span data-ttu-id="f8cef-142">入れ子になった配列はパラメーターとしてサポートされます (ただし、戻り値の型としてサポートされません)。</span><span class="sxs-lookup"><span data-stu-id="f8cef-142">Nested arrays are supported as parameters (but not as return types).</span></span>

6. <span data-ttu-id="f8cef-143">共用体の型は、単一の型 (など) に属するリテラルの共用体である場合に使用できます `"Left" | "Right"` 。</span><span class="sxs-lookup"><span data-stu-id="f8cef-143">Union types are allowed if they are a union of literals belonging to a single type (such as `"Left" | "Right"`).</span></span> <span data-ttu-id="f8cef-144">未定義のサポートされている型の共用体もサポートされています (など `string | undefined` )。</span><span class="sxs-lookup"><span data-stu-id="f8cef-144">Unions of a supported type with undefined are also supported (such as `string | undefined`).</span></span>

7. <span data-ttu-id="f8cef-145">オブジェクト型は、型、、サポートされている配列、または他のサポートされているオブジェクトのプロパティが含まれている `string` `number` `boolean` 場合に使用できます。</span><span class="sxs-lookup"><span data-stu-id="f8cef-145">Object types are allowed if they contain properties of type `string`, `number`, `boolean`, supported arrays, or other supported objects.</span></span> <span data-ttu-id="f8cef-146">次の例は、パラメーターの種類としてサポートされている入れ子になったオブジェクトを示しています。</span><span class="sxs-lookup"><span data-stu-id="f8cef-146">The following example shows nested objects that are supported as parameter types:</span></span>

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

8. <span data-ttu-id="f8cef-147">オブジェクトには、スクリプトで定義されているインターフェイスまたはクラス定義が必要です。</span><span class="sxs-lookup"><span data-stu-id="f8cef-147">Objects must have their interface or class definition defined in the script.</span></span> <span data-ttu-id="f8cef-148">次の例のように、オブジェクトをインラインで匿名で定義できます。</span><span class="sxs-lookup"><span data-stu-id="f8cef-148">An object can also be defined anonymously inline, as in the following example:</span></span>

    ```TypeScript
    function main(workbook: ExcelScript.Workbook): {name: string, email: string}
    ```

9. <span data-ttu-id="f8cef-149">省略可能なパラメーターは許可され、省略可能な修飾子 (たとえば) を使用して `?` 指定できます `function main(workbook: ExcelScript.Workbook, Name?: string)` 。</span><span class="sxs-lookup"><span data-stu-id="f8cef-149">Optional parameters are allowed and can be denoted as such by using the optional modifier `?` (for example, `function main(workbook: ExcelScript.Workbook, Name?: string)`).</span></span>

10. <span data-ttu-id="f8cef-150">既定のパラメーター値を使用できます (たとえば `async function main(workbook: ExcelScript.Workbook, Name: string = 'Jane Doe')` .</span><span class="sxs-lookup"><span data-stu-id="f8cef-150">Default parameter values are allowed (for example `async function main(workbook: ExcelScript.Workbook, Name: string = 'Jane Doe')`.</span></span>

### <a name="returning-data-from-a-script"></a><span data-ttu-id="f8cef-151">スクリプトからデータを返す</span><span class="sxs-lookup"><span data-stu-id="f8cef-151">Returning data from a script</span></span>

<span data-ttu-id="f8cef-152">スクリプトは、Power Automate フローで動的コンテンツとして使用するブックからデータを返します。</span><span class="sxs-lookup"><span data-stu-id="f8cef-152">Scripts can return data from the workbook to be used as dynamic content in a Power Automate flow.</span></span> <span data-ttu-id="f8cef-153">入力パラメーターと同様に、Power Automate は戻り値の種類にいくつかの制限を設定します。</span><span class="sxs-lookup"><span data-stu-id="f8cef-153">As with input parameters, Power Automate places some restrictions on the return type.</span></span>

1. <span data-ttu-id="f8cef-154">基本の型 `string` `number` `boolean` `void` 、、、 `undefined` がサポートされています。</span><span class="sxs-lookup"><span data-stu-id="f8cef-154">The basic types `string`, `number`, `boolean`, `void`, and `undefined` are supported.</span></span>

2. <span data-ttu-id="f8cef-155">戻り値の型として使用される Union 型は、スクリプト パラメーターとして使用する場合と同じ制限に従います。</span><span class="sxs-lookup"><span data-stu-id="f8cef-155">Union types used as return types follow the same restrictions as they do when used as script parameters.</span></span>

3. <span data-ttu-id="f8cef-156">配列型は、型 、、または `string` の場合 `number` に使用できます `boolean` 。</span><span class="sxs-lookup"><span data-stu-id="f8cef-156">Array types are allowed if they are of type `string`, `number`, or `boolean`.</span></span> <span data-ttu-id="f8cef-157">また、この型がサポートされている共用体またはサポートされているリテラル型である場合にも使用できます。</span><span class="sxs-lookup"><span data-stu-id="f8cef-157">They are also allowed if the type is a supported union or supported literal type.</span></span>

4. <span data-ttu-id="f8cef-158">戻り値の型として使用されるオブジェクトの種類は、スクリプト パラメーターとして使用する場合と同じ制限に従います。</span><span class="sxs-lookup"><span data-stu-id="f8cef-158">Object types used as return types follow the same restrictions as they do when used as script parameters.</span></span>

5. <span data-ttu-id="f8cef-159">暗黙的な型指定はサポートされています。定義された型と同じルールに従う必要があります。</span><span class="sxs-lookup"><span data-stu-id="f8cef-159">Implicit typing is supported, though it must follow the same rules as a defined type.</span></span>

## <a name="example"></a><span data-ttu-id="f8cef-160">例</span><span class="sxs-lookup"><span data-stu-id="f8cef-160">Example</span></span>

<span data-ttu-id="f8cef-161">次のスクリーンショットは、GitHub の問題が割り当てられるたびにトリガーされる [Power Automate](https://github.com/) フローを示しています。</span><span class="sxs-lookup"><span data-stu-id="f8cef-161">The following screenshot shows a Power Automate flow that's triggered whenever a [GitHub](https://github.com/) issue is assigned to you.</span></span> <span data-ttu-id="f8cef-162">フローは、Excel ブック内のテーブルに問題を追加するスクリプトを実行します。</span><span class="sxs-lookup"><span data-stu-id="f8cef-162">The flow runs a script that adds the issue to a table in an Excel workbook.</span></span> <span data-ttu-id="f8cef-163">そのテーブルに 5 つ以上の問題がある場合、フローはメールリマインダーを送信します。</span><span class="sxs-lookup"><span data-stu-id="f8cef-163">If there are five or more issues in that table, the flow sends an email reminder.</span></span>

:::image type="content" source="../images/power-automate-parameter-return-sample.png" alt-text="例のフローを示す Power Automate フロー エディター。":::

<span data-ttu-id="f8cef-165">スクリプトの関数は、問題 ID と発行タイトルを入力パラメーターとして指定し、スクリプトは問題テーブル内の行数 `main` を返します。</span><span class="sxs-lookup"><span data-stu-id="f8cef-165">The `main` function of the script specifies the issue ID and issue title as input parameters, and the script returns the number of rows in the issue table.</span></span>

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

## <a name="see-also"></a><span data-ttu-id="f8cef-166">関連項目</span><span class="sxs-lookup"><span data-stu-id="f8cef-166">See also</span></span>

- [<span data-ttu-id="f8cef-167">Power Automate Office Excel で Web 上のスクリプトを実行する</span><span class="sxs-lookup"><span data-stu-id="f8cef-167">Run Office Scripts in Excel on the web with Power Automate</span></span>](../tutorials/excel-power-automate-manual.md)
- [<span data-ttu-id="f8cef-168">自動で実行される Power Automate フロー内で、データをスクリプトに渡す</span><span class="sxs-lookup"><span data-stu-id="f8cef-168">Pass data to scripts in an automatically-run Power Automate flow</span></span>](../tutorials/excel-power-automate-trigger.md)
- [<span data-ttu-id="f8cef-169">自動で実行される Power Automate フローにスクリプトからデータを返す</span><span class="sxs-lookup"><span data-stu-id="f8cef-169">Return data from a script to an automatically-run Power Automate flow</span></span>](../tutorials/excel-power-automate-returns.md)
- [<span data-ttu-id="f8cef-170">Power Automate with Office スクリプト</span><span class="sxs-lookup"><span data-stu-id="f8cef-170">Troubleshooting information for Power Automate with Office Scripts</span></span>](../testing/power-automate-troubleshooting.md)
- [<span data-ttu-id="f8cef-171">Power Automate の使用を開始する</span><span class="sxs-lookup"><span data-stu-id="f8cef-171">Get started with Power Automate</span></span>](/power-automate/getting-started)
- [<span data-ttu-id="f8cef-172">Excel Online (Business) コネクタリファレンス ドキュメント</span><span class="sxs-lookup"><span data-stu-id="f8cef-172">Excel Online (Business) connector reference documentation</span></span>](/connectors/excelonlinebusiness/)
