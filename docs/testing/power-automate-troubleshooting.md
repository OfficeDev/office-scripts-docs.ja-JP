---
title: Power Automate with Office スクリプト
description: スクリプトと Power Automate の間の統合に関するヒント、プラットフォーム情報、既知Office問題。
ms.date: 01/14/2021
localization_priority: Normal
ms.openlocfilehash: 59f4cd8b3476c2ee2a1a862f136173a543ba8a15
ms.sourcegitcommit: 45ffe3dbd2c834b78592ad35928cf8096f5e80bc
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 04/14/2021
ms.locfileid: "51755008"
---
# <a name="troubleshooting-information-for-power-automate-with-office-scripts"></a>Power Automate with Office スクリプト

Power Automation を使用すると、スクリプトOfficeを次のレベルに進めできます。 ただし、Power Automate は独立した Excel セッションでスクリプトを代理で実行しますので、いくつかの重要な点に注意してください。

> [!TIP]
> Power Automate を使用して Office スクリプトを使い始める場合は、「Power [Automate](../develop/power-automate-integration.md) を使用して Office スクリプトを実行する」から始め、プラットフォームについて学習してください。

## <a name="avoid-using-relative-references"></a>相対参照の使用を避ける

Power Automate は、選択した Excel ブックでスクリプトを代理で実行します。 この場合、ブックが閉じられます。 Power Automate では、ユーザーの現在の状態 (など) に依存する API の動作 `Workbook.getActiveWorksheet` が異なる場合があります。 これは、API がユーザーのビューまたはカーソルの相対位置に基づいており、その参照が Power Automate フローに存在しないのでです。

一部の相対参照 API は、Power Automate でエラーをスローします。 他のユーザーは、ユーザーの状態を意味する既定の動作を持っています。 スクリプトを設計する場合は、ワークシートと範囲に絶対参照を使用してください。 これにより、ワークシートが再配置された場合でも、Power Automate フローの整合性が保たれる。

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

### <a name="script-methods-with-a-default-behavior-in-power-automate-flows"></a>Power Automate フローの既定の動作を持つスクリプト メソッド

次のメソッドは、ユーザーの現在の状態の代りとして、既定の動作を使用します。

| クラス | Method | Power Automate の動作 |
|--|--|--|
| [Workbook](/javascript/api/office-scripts/excelscript/excelscript.workbook) | `getActiveWorksheet` | ブックの最初のワークシート、またはメソッドによって現在アクティブ化されているワークシートのいずれかを返 `Worksheet.activate` します。 |
| [Worksheet](/javascript/api/office-scripts/excelscript/excelscript.worksheet) | `activate` | の目的でワークシートをアクティブなワークシートとしてマークします `Workbook.getActiveWorksheet` 。 |

## <a name="select-workbooks-with-the-file-browser-control"></a>ファイル ブラウザー コントロールを使用してブックを選択する

Power Automate フローの **スクリプト** の実行ステップを作成する場合は、フローの一部であるブックを選択する必要があります。 ブックの名前を手動で入力する代わりに、ファイル ブラウザーを使用してブックを選択します。

:::image type="content" source="../images/power-automate-file-browser.png" alt-text="[選択ウィンドウのファイル ブラウザーを表示する] オプションを示す Power Automate Run スクリプト アクション。":::

Power Automate の制限に関する詳細なコンテキストと、ブックの動的選択に関する潜在的な回避策の説明については、Microsoft Power Automate コミュニティのこのスレッド [を参照してください](https://powerusers.microsoft.com/t5/Power-Automate-Ideas/Allow-for-dynamic-quot-file-quot-value-for-excel-quot-get-a-row/idi-p/103091#)。

## <a name="time-zone-differences"></a>タイム ゾーンの違い

Excel ファイルには、固有の場所やタイム ゾーンが存在しない。 ユーザーがブックを開くたび、そのユーザーのローカル タイム ゾーンを日付の計算に使用します。 Power Automate は常に UTC を使用します。

スクリプトで日付または時刻を使用する場合、スクリプトがローカルでテストされる場合と Power Automate を使用して実行する場合の動作の違いがあります。 Power Automate を使用すると、変換、書式設定、および調整を行います。 「Power [Automate」](https://flow.microsoft.com/blog/working-with-dates-and-times/)および[ `main` 「Parameters: Passing](../develop/power-automate-integration.md#main-parameters-passing-data-to-a-script) data to a script」のこれらの関数の使い方については、「フロー内の日付と時刻を操作する」を参照してください。

## <a name="see-also"></a>関連項目

- [Office スクリプトのトラブルシューティング](troubleshooting.md)
- [Power Automate Officeスクリプトを実行する](../develop/power-automate-integration.md)
- [Excel Online (Business) コネクタリファレンス ドキュメント](/connectors/excelonlinebusiness/)
