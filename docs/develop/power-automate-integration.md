---
title: Power Automate を使用した Office スクリプトの実行
description: Power Automate ワークフローを使用して Excel on the web の Office スクリプトを取得する方法。
ms.date: 05/17/2021
localization_priority: Normal
ms.openlocfilehash: 7562a2b2359cde67a9a47e0640515018fe23ac35
ms.sourcegitcommit: 4687693f02fc90a57ba30c461f35046e02e6f5fb
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 05/19/2021
ms.locfileid: "52545041"
---
# <a name="run-office-scripts-with-power-automate"></a><span data-ttu-id="db8be-103">Power Automate を使用した Office スクリプトの実行</span><span class="sxs-lookup"><span data-stu-id="db8be-103">Run Office Scripts with Power Automate</span></span>

<span data-ttu-id="db8be-104">[Power Automate](https://flow.microsoft.com) を使用すると、Office スクリプトを大規模で自動化されたワークフローに追加できます。</span><span class="sxs-lookup"><span data-stu-id="db8be-104">[Power Automate](https://flow.microsoft.com) lets you add Office Scripts to a larger, automated workflow.</span></span> <span data-ttu-id="db8be-105">Power Automate を使って、メールの内容をワークシートのテーブルに追加したり、ブックのコメントに基づいてプロジェクト管理ツールでアクションを作成したりできます。</span><span class="sxs-lookup"><span data-stu-id="db8be-105">You can use Power Automate do things like add the contents of an email to a worksheet's table or create actions in your project management tools based on workbook comments.</span></span>

## <a name="get-started"></a><span data-ttu-id="db8be-106">作業の開始</span><span class="sxs-lookup"><span data-stu-id="db8be-106">Get started</span></span>

<span data-ttu-id="db8be-107">Power Automate を初めて使用する場合は、「[Power Automate に関する入門情報](/power-automate/getting-started)」にアクセスすることをお勧めします。</span><span class="sxs-lookup"><span data-stu-id="db8be-107">If you are new to Power Automate, we recommend visiting [Get started with Power Automate](/power-automate/getting-started).</span></span> <span data-ttu-id="db8be-108">そちらで、利用可能なすべてのオートメーションの可能性について詳しく学ぶことができます。</span><span class="sxs-lookup"><span data-stu-id="db8be-108">There, you can learn more about all the automation possibilities available to you.</span></span> <span data-ttu-id="db8be-109">このドキュメントでは、Power Automate での Office スクリプトの動作と、それが Excel エクスペリエンスの改善にどのように役立つかに重点が置かれています。</span><span class="sxs-lookup"><span data-stu-id="db8be-109">The documents here focus on how Office Scripts work with Power Automate and how that can help improve your Excel experience.</span></span>

<span data-ttu-id="db8be-110">Power Automate と Office スクリプトの統合を開始するには、チュートリアル「[Power Automate でスクリプトの使用を開始する](../tutorials/excel-power-automate-manual.md)」に従ってください。</span><span class="sxs-lookup"><span data-stu-id="db8be-110">To begin combining Power Automate and Office Scripts, follow the tutorial [Start using scripts with Power Automate](../tutorials/excel-power-automate-manual.md).</span></span> <span data-ttu-id="db8be-111">単純なスクリプトを呼び出すフローの作成方法について学ぶことができます。</span><span class="sxs-lookup"><span data-stu-id="db8be-111">This will teach you how to create a flow that calls a simple script.</span></span> <span data-ttu-id="db8be-112">このチュートリアルと「[自動で実行される Power Automate フロー内で、データをスクリプトに渡す](../tutorials/excel-power-automate-trigger.md)」のチュートリアルが完了したら、こちらに戻り、Office スクリプトを Power Automate フローに接続する方法の詳細をご確認ください。</span><span class="sxs-lookup"><span data-stu-id="db8be-112">After you've completed that tutorial and the [Pass data to scripts in an automatically-run Power Automate flow](../tutorials/excel-power-automate-trigger.md) tutorial, return here for detailed information about connecting Office Scripts to Power Automate flows.</span></span>

## <a name="excel-online-business-connector"></a><span data-ttu-id="db8be-113">Excel Online (Business) コネクタ</span><span class="sxs-lookup"><span data-stu-id="db8be-113">Excel Online (Business) connector</span></span>

<span data-ttu-id="db8be-114">[コネクタ](/connectors/connectors)は、Power Automate とアプリケーション間のブリッジです。</span><span class="sxs-lookup"><span data-stu-id="db8be-114">[Connectors](/connectors/connectors) are the bridges between Power Automate and applications.</span></span> <span data-ttu-id="db8be-115">[Excel Online (Business) コネクタ](/connectors/excelonlinebusiness)を使用すると、フローに Excel ブックへのアクセスが提供されます。</span><span class="sxs-lookup"><span data-stu-id="db8be-115">The [Excel Online (Business) connector](/connectors/excelonlinebusiness) gives your flows access to Excel workbooks.</span></span> <span data-ttu-id="db8be-116">"スクリプトの実行" アクションにより、選択したブックからアクセスできるすべての Office スクリプトを呼び出すことができます。</span><span class="sxs-lookup"><span data-stu-id="db8be-116">The "Run script" action lets you call any Office Script accessible through the selected workbook.</span></span> <span data-ttu-id="db8be-117">また、フローによってデータを提供したり、フローの後の手順用にスクリプトで情報を返したりできるよう、スクリプトに入力パラメーターを指定することもできます。</span><span class="sxs-lookup"><span data-stu-id="db8be-117">You can also give your scripts input parameters so data can be provided by the flow, or have your script return information for later steps in the flow.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="db8be-118">"スクリプトの実行" アクションにより、Excel コネクタを使用するユーザーにブックとそのデータへの重要なアクセス権が付与されます。</span><span class="sxs-lookup"><span data-stu-id="db8be-118">The "Run script" action gives people who use the Excel connector significant access to your workbook and its data.</span></span> <span data-ttu-id="db8be-119">さらに、「[Power Automate からの外部呼び出し](external-calls.md)」で説明されているとおり、外部 API の呼び出しを行うスクリプトにセキュリティ上のリスクがあります。</span><span class="sxs-lookup"><span data-stu-id="db8be-119">Additionally, there are security risks with scripts that make external API calls, as explained in [External calls from Power Automate](external-calls.md).</span></span> <span data-ttu-id="db8be-120">管理者が機密性の高いデータの流出を懸念している場合は、Excel Online コネクタをオフにするか、[Office スクリプト管理者制御](/microsoft-365/admin/manage/manage-office-scripts-settings)で Office スクリプトへのアクセスを制限することができます。</span><span class="sxs-lookup"><span data-stu-id="db8be-120">If your admin is concerned with the exposure of highly sensitive data, they can either turn off the Excel Online connector or restrict access to Office Scripts through the [Office Scripts administrator controls](/microsoft-365/admin/manage/manage-office-scripts-settings).</span></span>

## <a name="data-transfer-in-flows-for-scripts"></a><span data-ttu-id="db8be-121">スクリプトのフローでのデータ転送</span><span class="sxs-lookup"><span data-stu-id="db8be-121">Data transfer in flows for scripts</span></span>

<span data-ttu-id="db8be-122">Power Automate を使用すると、フローのステップ間でデータの一部を渡すことができます。</span><span class="sxs-lookup"><span data-stu-id="db8be-122">Power Automate lets you pass pieces of data between steps of your flow.</span></span> <span data-ttu-id="db8be-123">スクリプトを構成して、必要な種類の情報を受け入れたり、フローに必要なものをブックから返したりすることができます。</span><span class="sxs-lookup"><span data-stu-id="db8be-123">Scripts can be configured to accept whatever types of information you need and return anything from your workbook that you want in your flow.</span></span> <span data-ttu-id="db8be-124">スクリプトの入力は、(`workbook: ExcelScript.Workbook` に加えて) `main` 関数にパラメーターを追加することによって指定されます。</span><span class="sxs-lookup"><span data-stu-id="db8be-124">Input for your script is specified by adding parameters to the `main` function (in addition to `workbook: ExcelScript.Workbook`).</span></span> <span data-ttu-id="db8be-125">スクリプトからの出力は、`main` に戻り値の型を追加することによって宣言されます。</span><span class="sxs-lookup"><span data-stu-id="db8be-125">Output from the script is declared by adding a return type to `main`.</span></span>

> [!NOTE]
> <span data-ttu-id="db8be-126">フローで "スクリプトの実行" ブロックを作成すると、承認されたパラメーターと返された型が入力されます。</span><span class="sxs-lookup"><span data-stu-id="db8be-126">When you create a "Run Script" block in your flow, the accepted parameters and returned types are populated.</span></span> <span data-ttu-id="db8be-127">スクリプトのパラメーターまたは戻り値の型を変更する場合は、フローの "スクリプトの実行" ブロックを再実行する必要があります。</span><span class="sxs-lookup"><span data-stu-id="db8be-127">If you change the parameters or return types of your script, you'll need to redo the "Run script" block of your flow.</span></span> <span data-ttu-id="db8be-128">これにより、データが正しく解析されていることが確認されます。</span><span class="sxs-lookup"><span data-stu-id="db8be-128">This ensures the data is being parsed correctly.</span></span>

<span data-ttu-id="db8be-129">次のセクションでは、Power Automate で使用されるスクリプトの入力と出力の詳細について説明します。</span><span class="sxs-lookup"><span data-stu-id="db8be-129">The following sections cover the details of input and output for scripts used in Power Automate.</span></span> <span data-ttu-id="db8be-130">このトピックについて学ぶための実践的なアプローチが必要な場合は、「[自動で実行される Power Automate フロー内で、データをスクリプトに渡す](../tutorials/excel-power-automate-trigger.md)」チュートリアルを試すか、[タスクの自動アラーム](../resources/scenarios/task-reminders.md)のサンプル シナリオを確認してください。</span><span class="sxs-lookup"><span data-stu-id="db8be-130">If you'd like a hands-on approach to learning this topic, try out the [Pass data to scripts in an automatically-run Power Automate flow](../tutorials/excel-power-automate-trigger.md) tutorial or explore the [Automated task reminders](../resources/scenarios/task-reminders.md) sample scenario.</span></span>

### <a name="main-parameters-pass-data-to-a-script"></a><span data-ttu-id="db8be-131">`main` パラメーター: スクリプトにデータを渡す</span><span class="sxs-lookup"><span data-stu-id="db8be-131">`main` Parameters: Pass data to a script</span></span>

<span data-ttu-id="db8be-132">すべてのスクリプト入力は、`main` 関数の追加パラメーターとして指定されます。</span><span class="sxs-lookup"><span data-stu-id="db8be-132">All script input is specified as additional parameters for the `main` function.</span></span> <span data-ttu-id="db8be-133">たとえば、入力として名前を表す `string` をスクリプトで受け入れるようにする場合は、`main` 署名を `function main(workbook: ExcelScript.Workbook, name: string)` に変更します。</span><span class="sxs-lookup"><span data-stu-id="db8be-133">For example, if you wanted a script to accept a `string` that represents a name as input, you would change the `main` signature to `function main(workbook: ExcelScript.Workbook, name: string)`.</span></span>

<span data-ttu-id="db8be-134">Power Automate でフローを構成する場合、スクリプト入力を静的な値、[式](/power-automate/use-expressions-in-conditions)、または動的なコンテンツとして指定できます。</span><span class="sxs-lookup"><span data-stu-id="db8be-134">When you're configuring a flow in Power Automate, you can specify script input as static values, [expressions](/power-automate/use-expressions-in-conditions), or dynamic content.</span></span> <span data-ttu-id="db8be-135">個々のサービスのコネクタの詳細については、[Power Automate コネクタに関するドキュメント](/connectors/)を参照してください。</span><span class="sxs-lookup"><span data-stu-id="db8be-135">Details on an individual service's connector can be found in the [Power Automate Connector documentation](/connectors/).</span></span>

<span data-ttu-id="db8be-136">スクリプトの `main` 関数に入力パラメーターを追加する場合は、次の上限や制限を検討してください。</span><span class="sxs-lookup"><span data-stu-id="db8be-136">When adding input parameters to a script's `main` function, consider the following allowances and restrictions.</span></span>

1. <span data-ttu-id="db8be-137">最後のパラメーターは `ExcelScript.Workbook` の型にする必要があります。</span><span class="sxs-lookup"><span data-stu-id="db8be-137">The first parameter must be of type `ExcelScript.Workbook`.</span></span> <span data-ttu-id="db8be-138">そのパラメーター名は自由に指定できます。</span><span class="sxs-lookup"><span data-stu-id="db8be-138">Its parameter name does not matter.</span></span>

2. <span data-ttu-id="db8be-139">すべてのパラメーターには、型 (`string` または `number` など) が必要です。</span><span class="sxs-lookup"><span data-stu-id="db8be-139">Every parameter must have a type (such as `string` or `number`).</span></span>

3. <span data-ttu-id="db8be-140">基本的な型 `string` `number` `boolean` `unknown` 、、、、 `object` `undefined` がサポートされています。</span><span class="sxs-lookup"><span data-stu-id="db8be-140">The basic types `string`, `number`, `boolean`, `unknown`, `object`, and `undefined` are supported.</span></span>

4. <span data-ttu-id="db8be-141">以前に一覧表示された基本型の配列がサポートされています。</span><span class="sxs-lookup"><span data-stu-id="db8be-141">Arrays of the previously listed basic types are supported.</span></span>

5. <span data-ttu-id="db8be-142">入れ子にされた配列はパラメーターとしてサポートされます (戻り値の型としてはサポートされません)。</span><span class="sxs-lookup"><span data-stu-id="db8be-142">Nested arrays are supported as parameters (but not as return types).</span></span>

6. <span data-ttu-id="db8be-143">共用体型は、単一の型に属するリテラルの共用体 (`"Left" | "Right"` など) の場合に許可されます。</span><span class="sxs-lookup"><span data-stu-id="db8be-143">Union types are allowed if they are a union of literals belonging to a single type (such as `"Left" | "Right"`).</span></span> <span data-ttu-id="db8be-144">undefined を含むサポートされる型の共用体 (`string | undefined` など) もサポートされます。</span><span class="sxs-lookup"><span data-stu-id="db8be-144">Unions of a supported type with undefined are also supported (such as `string | undefined`).</span></span>

7. <span data-ttu-id="db8be-145">オブジェクト型は、型 `string`、`number`、`boolean`、サポートされている配列、または他のサポートされているオブジェクトのプロパティが含まれる場合に許可されます。</span><span class="sxs-lookup"><span data-stu-id="db8be-145">Object types are allowed if they contain properties of type `string`, `number`, `boolean`, supported arrays, or other supported objects.</span></span> <span data-ttu-id="db8be-146">次の例は、パラメーターの型としてサポートされる入れ子にされたオブジェクトを示しています。</span><span class="sxs-lookup"><span data-stu-id="db8be-146">The following example shows nested objects that are supported as parameter types:</span></span>

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

8. <span data-ttu-id="db8be-147">オブジェクトのインターフェイスまたはクラス定義はスクリプトで定義されている必要があります。</span><span class="sxs-lookup"><span data-stu-id="db8be-147">Objects must have their interface or class definition defined in the script.</span></span> <span data-ttu-id="db8be-148">次の例のように、オブジェクトをインラインで匿名で定義することができます。</span><span class="sxs-lookup"><span data-stu-id="db8be-148">An object can also be defined anonymously inline, as in the following example:</span></span>

    ```TypeScript
    function main(workbook: ExcelScript.Workbook): {name: string, email: string}
    ```

9. <span data-ttu-id="db8be-149">オプション パラメーターは許可されており、オプションの修飾子 `?` を使用してそのようなものとして示すことができます (例: `function main(workbook: ExcelScript.Workbook, Name?: string)`)。</span><span class="sxs-lookup"><span data-stu-id="db8be-149">Optional parameters are allowed and can be denoted as such by using the optional modifier `?` (for example, `function main(workbook: ExcelScript.Workbook, Name?: string)`).</span></span>

10. <span data-ttu-id="db8be-150">既定のパラメーター値は許可されています (例: `async function main(workbook: ExcelScript.Workbook, Name: string = 'Jane Doe')`)。</span><span class="sxs-lookup"><span data-stu-id="db8be-150">Default parameter values are allowed (for example `async function main(workbook: ExcelScript.Workbook, Name: string = 'Jane Doe')`.</span></span>

### <a name="return-data-from-a-script"></a><span data-ttu-id="db8be-151">スクリプトからデータを返す</span><span class="sxs-lookup"><span data-stu-id="db8be-151">Return data from a script</span></span>

<span data-ttu-id="db8be-152">スクリプトではブックからデータを返すことができ、Power Automate フローの動的なコンテンツとして使用することができます。</span><span class="sxs-lookup"><span data-stu-id="db8be-152">Scripts can return data from the workbook to be used as dynamic content in a Power Automate flow.</span></span> <span data-ttu-id="db8be-153">入力パラメーターと同様に、Power Automate では、戻り値の型にいくつかの制限が設定されます。</span><span class="sxs-lookup"><span data-stu-id="db8be-153">As with input parameters, Power Automate places some restrictions on the return type.</span></span>

1. <span data-ttu-id="db8be-154">基本型 `string`、`number`、`boolean`、`void`、`undefined` がサポートされています。</span><span class="sxs-lookup"><span data-stu-id="db8be-154">The basic types `string`, `number`, `boolean`, `void`, and `undefined` are supported.</span></span>

2. <span data-ttu-id="db8be-155">戻り値の型として使用される共用体の型は、スクリプト パラメーターとして使用する場合と同じ制限に従います。</span><span class="sxs-lookup"><span data-stu-id="db8be-155">Union types used as return types follow the same restrictions as they do when used as script parameters.</span></span>

3. <span data-ttu-id="db8be-156">配列型は、`string`、`number`、または `boolean` の型の場合に許可されます。</span><span class="sxs-lookup"><span data-stu-id="db8be-156">Array types are allowed if they are of type `string`, `number`, or `boolean`.</span></span> <span data-ttu-id="db8be-157">型がサポートされている共用体またはサポートされているリテラルの型の場合も許可されます。</span><span class="sxs-lookup"><span data-stu-id="db8be-157">They are also allowed if the type is a supported union or supported literal type.</span></span>

4. <span data-ttu-id="db8be-158">戻り値の型として使用されるオブジェクトの型は、スクリプト パラメーターとして使用する場合と同じ制限に従います。</span><span class="sxs-lookup"><span data-stu-id="db8be-158">Object types used as return types follow the same restrictions as they do when used as script parameters.</span></span>

5. <span data-ttu-id="db8be-159">暗黙的な入力はサポートされていますが、定義された型と同じ規則に従う必要があります。</span><span class="sxs-lookup"><span data-stu-id="db8be-159">Implicit typing is supported, though it must follow the same rules as a defined type.</span></span>

## <a name="example"></a><span data-ttu-id="db8be-160">例</span><span class="sxs-lookup"><span data-stu-id="db8be-160">Example</span></span>

<span data-ttu-id="db8be-161">次のスクリーンショットは、[GitHub](https://github.com/) の問題がお客様に割り当てられるたびにトリガーされる Power Automate フローを示しています。</span><span class="sxs-lookup"><span data-stu-id="db8be-161">The following screenshot shows a Power Automate flow that's triggered whenever a [GitHub](https://github.com/) issue is assigned to you.</span></span> <span data-ttu-id="db8be-162">このフローでは、Excel ブックのテーブルに問題を追加するスクリプトが実行されます。</span><span class="sxs-lookup"><span data-stu-id="db8be-162">The flow runs a script that adds the issue to a table in an Excel workbook.</span></span> <span data-ttu-id="db8be-163">そのテーブルに 5 つ以上の問題がある場合、フローでメール アラームが送信されます。</span><span class="sxs-lookup"><span data-stu-id="db8be-163">If there are five or more issues in that table, the flow sends an email reminder.</span></span>

:::image type="content" source="../images/power-automate-parameter-return-sample.png" alt-text="フロー Power Automate示すフロー エディターの例":::

<span data-ttu-id="db8be-165">スクリプトの `main` 関数では、問題の ID と問題のタイトルが入力パラメーターとして指定され、スクリプトによって問題テーブルの行数が返されます。</span><span class="sxs-lookup"><span data-stu-id="db8be-165">The `main` function of the script specifies the issue ID and issue title as input parameters, and the script returns the number of rows in the issue table.</span></span>

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

## <a name="see-also"></a><span data-ttu-id="db8be-166">関連項目</span><span class="sxs-lookup"><span data-stu-id="db8be-166">See also</span></span>

- [<span data-ttu-id="db8be-167">Power Automate を使用して、Excel on the web で Office スクリプトを実行する</span><span class="sxs-lookup"><span data-stu-id="db8be-167">Run Office Scripts in Excel on the web with Power Automate</span></span>](../tutorials/excel-power-automate-manual.md)
- [<span data-ttu-id="db8be-168">自動で実行される Power Automate フロー内で、データをスクリプトに渡す</span><span class="sxs-lookup"><span data-stu-id="db8be-168">Pass data to scripts in an automatically-run Power Automate flow</span></span>](../tutorials/excel-power-automate-trigger.md)
- [<span data-ttu-id="db8be-169">自動で実行される Power Automate フローにスクリプトからデータを返す</span><span class="sxs-lookup"><span data-stu-id="db8be-169">Return data from a script to an automatically-run Power Automate flow</span></span>](../tutorials/excel-power-automate-returns.md)
- [<span data-ttu-id="db8be-170">Office スクリプトを使用した Power Automate のトラブルシューティング情報</span><span class="sxs-lookup"><span data-stu-id="db8be-170">Troubleshooting information for Power Automate with Office Scripts</span></span>](../testing/power-automate-troubleshooting.md)
- [<span data-ttu-id="db8be-171">Power Automate の使用を開始する</span><span class="sxs-lookup"><span data-stu-id="db8be-171">Get started with Power Automate</span></span>](/power-automate/getting-started)
- [<span data-ttu-id="db8be-172">Excel Online (Business) コネクタ リファレンス ドキュメント</span><span class="sxs-lookup"><span data-stu-id="db8be-172">Excel Online (Business) connector reference documentation</span></span>](/connectors/excelonlinebusiness/)
