---
title: Office スクリプトを使用した Power Automate のトラブルシューティング情報
description: Office Scripts と Power Automate の統合に関するヒント、プラットフォーム情報、既知の問題。
ms.date: 01/14/2021
localization_priority: Normal
ms.openlocfilehash: b0f5b2f542216789f0d96f309cb7d799d201ba0f
ms.sourcegitcommit: e7e019ba36c2f49451ec08c71a1679eb6dba4268
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 01/22/2021
ms.locfileid: "49933267"
---
# <a name="troubleshooting-information-for-power-automate-with-office-scripts"></a>Office スクリプトを使用した Power Automate のトラブルシューティング情報

Power Automate を使用すると、Officeスクリプトの自動化を次のレベルに進めできます。 ただし、Power Automate は独立した Excel セッションでユーザーに代わってスクリプトを実行しますが、注意が必要ないくつかの重要な点があります。

> [!TIP]
> Power Automate で Office スクリプトを使い始め始めたばかりの場合は、Power Automate を使った [Office スクリプト](../develop/power-automate-integration.md) の実行から始め、プラットフォームについて学習してください。

## <a name="avoid-using-relative-references"></a>相対参照の使用を避ける

Power Automate は、選択した Excel ブックでスクリプトをユーザーに代わって実行します。 この場合、ブックが閉じられます。 Power Automate では、ユーザーの現在の状態に依存する API (など) の動作 `Workbook.getActiveWorksheet` が異なる場合があります。 これは、API がユーザーのビューまたはカーソルの相対位置に基づいており、その参照が Power Automate フローに存在しないためです。

一部の相対参照 API は Power Automate でエラーをスローします。 他のユーザーは、ユーザーの状態を意味する既定の動作を持っています。 スクリプトを設計する場合は、ワークシートと範囲の絶対参照を使用してください。 これにより、ワークシートが再配置された場合でも、Power Automate フローの一貫性が維持されます。

### <a name="script-methods-that-fail-when-run-power-automate-flows"></a>Power Automate フローの実行時に失敗するスクリプト メソッド

次のメソッドは、Power Automate フローのスクリプトから呼び出された場合にエラーをスローし、失敗します。

| クラス | Method |
|--|--|
| [Chart](/javascript/api/office-scripts/excelscript/excelscript.chart) | `activate` |
| [Range](/javascript/api/office-scripts/excelscript/excelscript.range) | `select` |
| [Workbook](/javascript/api/office-scripts/excelscript/excelscript.workbook) | `getActiveCell` |
| [Workbook](/javascript/api/office-scripts/excelscript/excelscript.workbook) | `getActiveChart` |
| [Workbook](/javascript/api/office-scripts/excelscript/excelscript.workbook) | `getActiveSlicer` |
| [Workbook](/javascript/api/office-scripts/excelscript/excelscript.workbook) | `getSelectedRange` |
| [Workbook](/javascript/api/office-scripts/excelscript/excelscript.workbook) | `getSelectedRanges` |

### <a name="script-methods-with-a-default-behavior-in-power-automate-flows"></a>Power Automate フローでの既定の動作を持つスクリプト メソッド

次のメソッドは、ユーザーの現在の状態の代用として、既定の動作を使用します。

| クラス | Method | Power Automate の動作 |
|--|--|--|
| [Workbook](/javascript/api/office-scripts/excelscript/excelscript.workbook) | `getActiveWorksheet` | ブック内の最初のワークシート、またはメソッドによって現在アクティブになっているワークシートを返 `Worksheet.activate` します。 |
| [Worksheet](/javascript/api/office-scripts/excelscript/excelscript.worksheet) | `activate` | 次の目的のために、ワークシートをアクティブ ワークシートとしてマークします `Workbook.getActiveWorksheet` 。 |

## <a name="select-workbooks-with-the-file-browser-control"></a>ファイル ブラウザー コントロールを使用してブックを選択する

Power Automate フロー **のスクリプト** 実行ステップを作成する場合は、フローの一部であるブックを選択する必要があります。 ブックの名前を手動で入力する代わりに、ファイル ブラウザーを使用してブックを選択します。

![Power Automate で "スクリプトの実行" アクションを作成する場合のファイル ブラウザー オプション](../images/power-automate-file-browser.png)

Power Automate の制限の詳細と、ブックの動的な選択に対する潜在的な回避策については、Microsoft Power Automate コミュニティのこのスレッド [を参照してください](https://powerusers.microsoft.com/t5/Power-Automate-Ideas/Allow-for-dynamic-quot-file-quot-value-for-excel-quot-get-a-row/idi-p/103091#)。

## <a name="time-zone-differences"></a>タイム ゾーンの違い

Excel ファイルには、固有の場所やタイム ゾーンが存在しない。 ユーザーがブックを開くたび、セッションは日付の計算にユーザーのローカルのタイムゾーンを使用します。 Power Automate は常に UTC を使用します。

スクリプトで日付や時刻を使用する場合、スクリプトをローカルでテストする場合と Power Automate を使用してスクリプトを実行する場合とで動作が異なる可能性があります。 Power Automate を使用すると、変換、書式設定、および時間の調整を行います。 Power [](https://flow.microsoft.com/blog/working-with-dates-and-times/) Automate と Parameters でこれらの関数を使用する方法の手順については、「フロー内の日付と時刻を操作する[ `main` :](../develop/power-automate-integration.md#main-parameters-passing-data-to-a-script)スクリプトにデータを渡す」を参照して、その時間情報をスクリプトに提供する方法について説明します。

## <a name="see-also"></a>関連項目

- [Office スクリプトのトラブルシューティング](troubleshooting.md)
- [Power Automate Officeスクリプトを実行する](../develop/power-automate-integration.md)
- [Excel Online (Business) コネクタのリファレンス ドキュメント](/connectors/excelonlinebusiness/)
