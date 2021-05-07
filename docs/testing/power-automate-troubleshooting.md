---
title: Office スクリプトを使用した Power Automate のトラブルシューティング情報
description: ヒント、プラットフォーム情報、および既知の問題と、スクリプトとスクリプトのOffice統合Power Automate。
ms.date: 01/14/2021
localization_priority: Normal
ms.openlocfilehash: bcfedb8db88d74f16e46c604121bceff3c7c7382
ms.sourcegitcommit: f7a7aebfb687f2a35dbed07ed62ff352a114525a
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 05/06/2021
ms.locfileid: "52232649"
---
# <a name="troubleshooting-information-for-power-automate-with-office-scripts"></a>Office スクリプトを使用した Power Automate のトラブルシューティング情報

Power Automateスクリプトオートメーションを次Officeレベルに移動できます。 ただし、Power Automateに独立したセッションでスクリプトを実行Excel、いくつかの重要な点に注意してください。

> [!TIP]
> Power Automate で Office スクリプトを使用する場合は、Office スクリプトを Power Automate で実行[](../develop/power-automate-integration.md)して、プラットフォームについて説明します。

## <a name="avoid-using-relative-references"></a>相対参照の使用を避ける

Power Automate、選択したブックでスクリプトをExcel代わりに実行します。 この場合、ブックが閉じられます。 ユーザーの現在の状態 (など) に依存する API は、ユーザーの動作 `Workbook.getActiveWorksheet` が異Power Automate。 これは、API がユーザーのビューまたはカーソルの相対位置に基づいており、その参照がビュー フロー内に存在Power Automateです。

一部の相対参照 API は、エラーをスロー Power Automate。 他のユーザーは、ユーザーの状態を意味する既定の動作を持っています。 スクリプトを設計する場合は、ワークシートと範囲に絶対参照を使用してください。 これにより、ワークシートPower Automate場合でも、一貫性のあるフローを作成できます。

### <a name="script-methods-that-fail-when-run-power-automate-flows"></a>フローの実行時に失敗するスクリプト メソッドPower Automateします。

次のメソッドは、エラーをスローし、エラー フロー内のスクリプトから呼び出Power Automateします。

| クラス | メソッド |
|--|--|
| [グラフ](/javascript/api/office-scripts/excelscript/excelscript.chart) | `activate` |
| [Range](/javascript/api/office-scripts/excelscript/excelscript.range) | `select` |
| [ブック](/javascript/api/office-scripts/excelscript/excelscript.workbook) | `getActiveCell` |
| [ブック](/javascript/api/office-scripts/excelscript/excelscript.workbook) | `getActiveChart` |
| [ブック](/javascript/api/office-scripts/excelscript/excelscript.workbook) | `getActiveSlicer` |
| [ブック](/javascript/api/office-scripts/excelscript/excelscript.workbook) | `getSelectedRange` |
| [ブック](/javascript/api/office-scripts/excelscript/excelscript.workbook) | `getSelectedRanges` |

### <a name="script-methods-with-a-default-behavior-in-power-automate-flows"></a>スクリプト フローの既定の動作を持つスクリプト メソッドPower Automateします。

次のメソッドは、ユーザーの現在の状態の代りとして、既定の動作を使用します。

| クラス | メソッド | Power Automate動作 |
|--|--|--|
| [ブック](/javascript/api/office-scripts/excelscript/excelscript.workbook) | `getActiveWorksheet` | ブックの最初のワークシート、またはメソッドによって現在アクティブ化されているワークシートのいずれかを返 `Worksheet.activate` します。 |
| [ワークシート](/javascript/api/office-scripts/excelscript/excelscript.worksheet) | `activate` | の目的でワークシートをアクティブなワークシートとしてマークします `Workbook.getActiveWorksheet` 。 |

## <a name="select-workbooks-with-the-file-browser-control"></a>ファイル ブラウザー コントロールを使用してブックを選択する

アプリケーション フローの **スクリプトの実行** ステップをPower Automate、フローの一部であるブックを選択する必要があります。 ブックの名前を手動で入力する代わりに、ファイル ブラウザーを使用してブックを選択します。

:::image type="content" source="../images/power-automate-file-browser.png" alt-text="[Power Automateファイル ブラウザーの表示] オプションを示すスクリプトの実行アクション":::

ブックの動的選択に関するPower Automateの制限と潜在的な回避策に関する詳細なコンテキストについては、Microsoft Power Automate Community のこのスレッドを[参照してください](https://powerusers.microsoft.com/t5/Power-Automate-Ideas/Allow-for-dynamic-quot-file-quot-value-for-excel-quot-get-a-row/idi-p/103091#)。

## <a name="time-zone-differences"></a>タイム ゾーンの違い

Excelファイルに固有の場所やタイム ゾーンが存在しない。 ユーザーがブックを開くたび、そのユーザーのローカル タイム ゾーンを日付の計算に使用します。 Power Automateは常に UTC を使用します。

スクリプトで日付または時刻を使用する場合、スクリプトがローカルでテストされる場合と、スクリプトを実行するときに動作の違いPower Automate。 Power Automateを使用すると、変換、書式設定、調整を行います。 Power Automate[](https://flow.microsoft.com/blog/working-with-dates-and-times/)および[ `main` Parameters で](../develop/power-automate-integration.md#main-parameters-passing-data-to-a-script)これらの関数を使用する方法の手順については、「フロー内の日付と時刻を操作する:スクリプトにデータを渡す」を参照して、スクリプトの時間情報を提供する方法について説明します。

## <a name="see-also"></a>こちらもご覧ください

- [Office スクリプトのトラブルシューティング](troubleshooting.md)
- [Power Automate を使用した Office スクリプトの実行](../develop/power-automate-integration.md)
- [Excel Online (Business) コネクタ リファレンス ドキュメント](/connectors/excelonlinebusiness/)
