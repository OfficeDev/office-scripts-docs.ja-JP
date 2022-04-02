---
title: Officeで実行されているスクリプトのトラブルシューティングPower Automate
description: ヒント、プラットフォーム情報、および既知の問題と、スクリプトとスクリプトのOffice統合Power Automate。
ms.date: 11/01/2021
ms.localizationpriority: medium
ms.openlocfilehash: 2c256c2ddc64fcfc510f24e27662234f44b65ac0
ms.sourcegitcommit: 7023b9e23499806901a5ecf8ebc460b76887cca6
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 03/31/2022
ms.locfileid: "64586032"
---
# <a name="troubleshoot-office-scripts-running-in-power-automate"></a>Officeで実行されているスクリプトのトラブルシューティングPower Automate

Power Automateスクリプトの自動化をOfficeレベルに移動できます。 ただし、Power Automateセッションでスクリプトを代理で実行Excel、いくつかの重要な点に注意してください。

> [!TIP]
> Power Automate で Office スクリプトを使い始める場合は、Office スクリプトと Power Automate を実行してプラットフォームについて説明します[](../develop/power-automate-integration.md)。

## <a name="avoid-relative-references"></a>相対参照を避ける

Power Automate選択したブックでスクリプトをExcel代わりに実行します。 この場合、ブックが閉じられます。 ユーザーの現在の状態 `Workbook.getActiveWorksheet`(など) に依存する API は、アプリケーションの動作が異Power Automate。 これは、API がユーザーのビューまたはカーソルの相対的な位置に基づいており、その参照がユーザー フロー内に存在Power Automateです。

一部の相対参照 API は、エラーをスロー Power Automate。 他のユーザーは、ユーザーの状態を意味する既定の動作を持っています。 スクリプトを設計する場合は、ワークシートと範囲に絶対参照を使用してください。 これにより、ワークシートPower Automate場合でも、フローの一貫性が保たれる可能性があります。

### <a name="script-methods-that-fail-when-run-in-power-automate-flows"></a>スクリプト フローで実行すると失敗するスクリプト メソッドPower Automateします。

次のメソッドは、エラーをスローし、エラー フロー内のスクリプトから呼び出Power Automateします。

| クラス | Method |
|--|--|
| [Chart](/javascript/api/office-scripts/excelscript/excelscript.chart) | `activate` |
| [Range](/javascript/api/office-scripts/excelscript/excelscript.range) | `select` |
| [Workbook](/javascript/api/office-scripts/excelscript/excelscript.workbook) | `getActiveCell` |
| [Workbook](/javascript/api/office-scripts/excelscript/excelscript.workbook) | `getActiveChart` |
| [Workbook](/javascript/api/office-scripts/excelscript/excelscript.workbook) | `getActiveSlicer` |
| [Workbook](/javascript/api/office-scripts/excelscript/excelscript.workbook) | `getSelectedRange` |
| [Workbook](/javascript/api/office-scripts/excelscript/excelscript.workbook) | `getSelectedRanges` |

### <a name="script-methods-with-a-default-behavior-in-power-automate-flows"></a>スクリプト フローの既定の動作を持つスクリプト メソッドPower Automateします。

次のメソッドは、ユーザーの現在の状態の代りとして、既定の動作を使用します。

| クラス | Method | Power Automate動作 |
|--|--|--|
| [Workbook](/javascript/api/office-scripts/excelscript/excelscript.workbook) | `getActiveWorksheet` | ブックの最初のワークシート、またはメソッドによって現在アクティブ化されているワークシートのいずれかを返 `Worksheet.activate` します。 |
| [Worksheet](/javascript/api/office-scripts/excelscript/excelscript.worksheet) | `activate` | の目的でワークシートをアクティブなワークシートとしてマークします `Workbook.getActiveWorksheet`。 |

## <a name="data-refresh-not-supported-in-power-automate"></a>データ更新は、データ更新プログラムではPower Automate

Officeスクリプトは、スクリプトで実行するとデータを更新Power Automate。 フローで呼び出 `PivotTable.refresh` された場合は何もしないなどのメソッド。 さらに、Power Automateブック リンクを使用する数式のデータ更新はトリガーされません。

### <a name="script-methods-that-do-nothing-when-run-in-power-automate-flows"></a>スクリプト フローで実行するときに何もしないスクリプト メソッドPower Automateします。

次のメソッドは、スクリプトを使用して呼び出した場合、スクリプトPower Automate。 それでも正常に返され、エラーはスローしません。

| クラス | Method |
|--|--|
| [PivotTable](/javascript/api/office-scripts/excelscript/excelscript.pivottable) | `refresh` |
| [Workbook](/javascript/api/office-scripts/excelscript/excelscript.workbook) | `refreshAllDataConnections` |
| [ブック](/javascript/api/office-scripts/excelscript/excelscript.workbook) | `refreshAllPivotTables` |
| [ワークシート](/javascript/api/office-scripts/excelscript/excelscript.worksheet) | `refreshAllPivotTables` |

## <a name="select-workbooks-with-the-file-browser-control"></a>ファイル ブラウザー コントロールを使用してブックを選択する

アプリケーション フローの **スクリプトの実行** ステップPower Automate、フローの一部であるブックを選択する必要があります。 ブックの名前を手動で入力する代わりに、ファイル ブラウザーを使用してブックを選択します。

:::image type="content" source="../images/power-automate-file-browser.png" alt-text="[Power Automateファイル ブラウザーの表示] オプションを示すスクリプトの実行アクションです。":::

ブックの動的選択のPower Automateの詳細なコンテキストと回避策の詳細については、Microsoft Power Automate Community のこのスレッド[を参照してください](https://powerusers.microsoft.com/t5/Power-Automate-Ideas/Allow-for-dynamic-quot-file-quot-value-for-excel-quot-get-a-row/idi-p/103091#)。

## <a name="pass-entire-arrays-as-script-parameters"></a>配列全体をスクリプト パラメーターとして渡す

Power Automateを使用すると、ユーザーは配列を変数として、または配列内の 1 つの要素としてコネクタに渡します。 既定では、フロー内に配列を構築する単一の要素を渡します。 配列全体を引数として受け取るスクリプトまたは他のコネクタの場合は、[配列全体を入力する切り替え] ボタンを選択して、配列を 1 つの完全なオブジェクトとして渡す必要があります。 このボタンは、各配列パラメーター入力フィールドの右上隅にあります。

:::image type="content" source="../images/combine-worksheets-flow-3.png" alt-text="コントロール フィールド入力ボックスに配列全体を入力するために切り替えるボタン。":::

## <a name="time-zone-differences"></a>タイム ゾーンの違い

Excelファイルに固有の場所やタイム ゾーンが存在しない場合。 ユーザーがブックを開くたび、そのユーザーのローカル タイム ゾーンを日付の計算に使用します。 Power Automateは常に UTC を使用します。

スクリプトで日付または時刻を使用する場合、スクリプトがローカルでテストされる場合と、スクリプトがローカルで実行される場合と、スクリプトの動作に違いPower Automate。 Power Automateを使用すると、変換、書式設定、調整を行います。 Power Automate [](https://flow.microsoft.com/blog/working-with-dates-and-times/) および Parameters: Pass data to a script でこれらの関数を使用する方法については、「フロー[`main`](../develop/power-automate-integration.md#main-parameters-pass-data-to-a-script)内の日付と時刻の操作」を参照して、スクリプトの時間情報を提供する方法について説明します。

## <a name="script-parameter-fields-or-returned-output-not-appearing-in-power-automate"></a>スクリプト パラメーター フィールドまたは返される出力が、スクリプト パラメーターフィールドに表示Power Automate

スクリプトのパラメーターまたは返されるデータが、データ フロー ビルダーに正確に反映されないPower Automateがあります。

- スクリプト署名 (パラメーターまたは戻り値) は、ビジネス (**online)** コネクタが追加Excel変更されています。
- スクリプト署名は、サポートされていない型を使用します。 パラメーターの下のリストに対して[](../develop/power-automate-integration.md#main-parameters-pass-data-to-a-script)型を[](../develop/power-automate-integration.md#return-data-from-a-script)確認し、「スクリプトを使用してスクリプトを実行する」[Officeを](../develop/power-automate-integration.md)Power Automateします。

スクリプトの署名は、作成時Excel **ビジネス (Online)** コネクタと一緒に格納されます。 古いコネクタを削除し、新しいコネクタを作成して、スクリプトの実行アクションの最新のパラメーターと戻り **値を取得** します。

## <a name="see-also"></a>関連項目

- [スクリプトOfficeトラブルシューティング](troubleshooting.md)
- [Power Automate を使用した Office スクリプトの実行](../develop/power-automate-integration.md)
- [Excel Online (Business) コネクタ リファレンス ドキュメント](/connectors/excelonlinebusiness/)
