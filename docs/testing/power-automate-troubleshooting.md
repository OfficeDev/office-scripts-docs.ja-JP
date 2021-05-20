---
title: Power Automateで実行されているOfficeスクリプトのトラブルシューティング
description: ヒント、プラットフォーム情報、およびOfficeスクリプトとPower Automateの統合に関する既知の問題。
ms.date: 05/17/2021
localization_priority: Normal
ms.openlocfilehash: e26378051c764d97b4e8d748abc85fbe095c7b03
ms.sourcegitcommit: 4687693f02fc90a57ba30c461f35046e02e6f5fb
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 05/19/2021
ms.locfileid: "52545572"
---
# <a name="troubleshoot-office-scripts-running-in-power-automate"></a>Power Automateで実行されているOfficeスクリプトのトラブルシューティング

Power Automateを使用すると、Officeスクリプトの自動化を次のレベルに引き上げます。 ただし、Power Automateは独立したExcel セッションでスクリプトを実行するため、注意すべき重要な点がいくつかあります。

> [!TIP]
> Power Automateでスクリプト Officeを使用し始めたばかりの場合は、Power Automateを使用して[Officeスクリプトを実行](../develop/power-automate-integration.md)してプラットフォームについて学んでください。

## <a name="avoid-relative-references"></a>相対参照を避ける

Power Automateは、選択したExcelブックでスクリプトを実行します。 この場合、ブックは閉じられる可能性があります。 など、ユーザーの現在の状態に依存する API `Workbook.getActiveWorksheet` は、Power Automateで異なる動作をする可能性があります。 これは、API がユーザーのビューまたはカーソルの相対位置に基づいており、その参照がPower Automateフローに存在しないためです。

一部の相対参照 API は、Power Automateでエラーをスローします。 その他のユーザーの状態を意味する既定の動作があります。 スクリプトを設計する場合は、ワークシートと範囲に絶対参照を使用してください。 これにより、ワークシートが並べ替えられた場合でも、Power Automateフローの一貫性が保たれます。

### <a name="script-methods-that-fail-when-run-power-automate-flows"></a>Power Automateフローの実行時に失敗するスクリプト メソッド

次のメソッドは、Power Automate フローのスクリプトから呼び出されるとエラーをスローし、失敗します。

| クラス | メソッド |
|--|--|
| [グラフ](/javascript/api/office-scripts/excelscript/excelscript.chart) | `activate` |
| [Range](/javascript/api/office-scripts/excelscript/excelscript.range) | `select` |
| [ブック](/javascript/api/office-scripts/excelscript/excelscript.workbook) | `getActiveCell` |
| [ブック](/javascript/api/office-scripts/excelscript/excelscript.workbook) | `getActiveChart` |
| [ブック](/javascript/api/office-scripts/excelscript/excelscript.workbook) | `getActiveSlicer` |
| [ブック](/javascript/api/office-scripts/excelscript/excelscript.workbook) | `getSelectedRange` |
| [ブック](/javascript/api/office-scripts/excelscript/excelscript.workbook) | `getSelectedRanges` |

### <a name="script-methods-with-a-default-behavior-in-power-automate-flows"></a>Power Automateフローでの既定の動作を持つスクリプト メソッド

次のメソッドは、既定の動作を使用します。

| クラス | メソッド | Power Automate動作 |
|--|--|--|
| [ブック](/javascript/api/office-scripts/excelscript/excelscript.workbook) | `getActiveWorksheet` | ブックの最初のワークシート、またはメソッドによって現在アクティブになっているワークシートを返 `Worksheet.activate` します。 |
| [ワークシート](/javascript/api/office-scripts/excelscript/excelscript.worksheet) | `activate` | ワークシートを の目的でアクティブなワークシートとしてマーク `Workbook.getActiveWorksheet` します。 |

## <a name="select-workbooks-with-the-file-browser-control"></a>ファイル ブラウザー コントロールを使用してブックを選択する

Power Automate フローの **[スクリプトの実行**] ステップを構築する場合は、フローの一部であるブックを選択する必要があります。 ブックの名前を手動で入力する代わりに、ファイル ブラウザを使用してブックを選択します。

:::image type="content" source="../images/power-automate-file-browser.png" alt-text="[ピッカー ファイル ブラウザーの表示] オプションを表示する [スクリプトの実行] アクションをPower Automate":::

Power Automateの制限に関する詳細なコンテキストと、ブックの動的選択に対する潜在的な回避策については[、Microsoft Power Automate Communityのこのスレッドを](https://powerusers.microsoft.com/t5/Power-Automate-Ideas/Allow-for-dynamic-quot-file-quot-value-for-excel-quot-get-a-row/idi-p/103091#)参照してください。

## <a name="time-zone-differences"></a>タイム ゾーンの違い

Excelファイルには、固有の場所やタイムゾーンがありません。 ユーザーがブックを開くたびに、セッションは日付の計算にユーザーのローカル タイム ゾーンを使用します。 Power Automateは常に UTC を使用します。

スクリプトで日付または時刻を使用する場合、スクリプトをローカルでテストするときと、Power Automateを実行する場合と動作に違いが生じる可能性があります。 Power Automateを使用すると、時間の変換、書式設定、および調整ができます。 Power Automateのこれらの関数を使用する方法については、[フロー内の日付と時刻の操作](https://flow.microsoft.com/blog/working-with-dates-and-times/)を参照してください[ `main` 。](../develop/power-automate-integration.md#main-parameters-pass-data-to-a-script)

## <a name="see-also"></a>関連項目

- [Office スクリプトのトラブルシューティング](troubleshooting.md)
- [Power Automate を使用した Office スクリプトの実行](../develop/power-automate-integration.md)
- [Excel Online (Business) コネクタ リファレンス ドキュメント](/connectors/excelonlinebusiness/)
