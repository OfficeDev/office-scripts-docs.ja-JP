---
title: JSON を使用して、Office スクリプトとの間でデータを渡す
description: 外部呼び出しとPower Automateで使用するためにデータを JSON オブジェクトに構造化する方法について説明します
ms.date: 06/10/2022
ms.localizationpriority: medium
ms.openlocfilehash: 753097183a18f5d20ca2c78a3748c7a1d968ad42
ms.sourcegitcommit: dd01979d34b3499360d2f79a56f8a8f24f480eed
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 06/15/2022
ms.locfileid: "66088159"
---
# <a name="use-json-to-pass-data-to-and-from-office-scripts"></a>JSON を使用して、Office スクリプトとの間でデータを渡す

[JSON (JavaScript オブジェクト表記)](https://www.w3schools.com/whatis/whatis_json.asp) は、データを格納および転送するための形式です。 各 JSON オブジェクトは、作成時に定義できる名前と値のペアのコレクションです。 JSON は、Excel内の範囲、テーブル、その他のデータ パターンの任意の複雑さを処理できるため、Office スクリプトに役立ちます。 JSON を使用すると、[Web サービス](external-calls.md)からの受信データを解析し、[Power Automate フロー](power-automate-integration.md)を介して複雑なオブジェクトを渡すことができます。

この記事では、Office スクリプトで JSON を使用することに重点を置きます。 まず、W3 Schools の [JSON 入門](https://www.w3schools.com/js/js_json_intro.asp) などの記事から形式の詳細を確認することをお勧めします。

## <a name="parse-json-data-into-a-range-or-table"></a>JSON データを範囲またはテーブルに解析する

JSON オブジェクトの配列は、アプリケーションと Web サービスの間でテーブル データの行を渡す一貫した方法を提供します。 このような場合、各 JSON オブジェクトは行を表し、プロパティは列を表します。 Office スクリプトは JSON 配列をループし、2D 配列として再アセンブルできます。 この配列は、範囲の値として設定され、ブックに格納されます。 プロパティ名をヘッダーとして追加してテーブルを作成することもできます。

次のスクリプトは、テーブルに変換される JSON データを示しています。 データは外部ソースから取得されないことに注意してください。 これについては、この記事の後半で説明します。

```typescript
/**
 * Sample JSON data. This would be replaced by external calls or
 * parameters getting data from Power Automate in a production script.
 */
const jsonData = [
  { "Action": "Edit", /* Action property with value of "Edit". */
    "N": 3370, /* N property with value of 3370. */
    "Percent": 17.85 /* Percent property with value of 17.85. */
  },
  // The rest of the object entries follow the same pattern.
  { "Action": "Paste", "N": 1171, "Percent": 6.2 },
  { "Action": "Clear", "N": 599, "Percent": 3.17 },
  { "Action": "Insert", "N": 352, "Percent": 1.86 },
  { "Action": "Delete", "N": 350, "Percent": 1.85 },
  { "Action": "Refresh", "N": 314, "Percent": 1.66 },
  { "Action": "Fill", "N": 286, "Percent": 1.51 },
];

/**
 * This script converts JSON data to an Excel table.
 */
function main(workbook: ExcelScript.Workbook) {
  // Create a new worksheet to store the imported data.
  const newSheet = workbook.addWorksheet();
  newSheet.activate();

  // Determine the data's shape by getting the properties in one object.
  // This assumes all the JSON objects have the same properties.
  const columnNames = getPropertiesFromJson(jsonData[0]);

  // Create the table headers using the property names.
  const headerRange = newSheet.getRangeByIndexes(0, 0, 1, columnNames.length);
  headerRange.setValues([columnNames]);

  // Create a new table with the headers.
  const newTable = newSheet.addTable(headerRange, true);

  // Add each object in the array of JSON objects to the table.
  const tableValues = jsonData.map(row => convertJsonToRow(row));
  newTable.addRows(-1, tableValues);
}

/**
 * This function turns a JSON object into an array to be used as a table row.
 */
function convertJsonToRow(obj: object) {
  const array: (string | number)[] = [];

  // Loop over each property and get the value. Their order will be the same as the column headers.
  for (let value in obj) {
    array.push(obj[value]);
  }
  return array;
}

/**
 * This function gets the property names from a single JSON object.
 */
function getPropertiesFromJson(obj: object) {
  const propertyArray: string[] = [];
  
  // Loop over each property in the object and store the property name in an array.
  for (let property in obj) {
    propertyArray.push(property);
  }

  return propertyArray;
}
```

> [!TIP]
> JSON の構造がわかっている場合は、独自のインターフェイスを作成して、特定のプロパティを簡単に取得できます。 JSON から配列への変換手順は、型セーフな参照に置き換えることができます。 次のコード スニペットは、新しい `ActionRow` インターフェイスを使用する呼び出しに置き換えられたこれらの手順 (コメントアウト済み) を示しています。 これにより、関数は `convertJsonToRow` 不要になります。
>
> ```typescript
>   // const tableValues = jsonData.map(row => convertJsonToRow(row));
>   // newTable.addRows(-1, tableValues);
>   // }
>
>      const actionRows: ActionRow[] = jsonData as ActionRow[];
>      // Add each object in the array of JSON objects to the table.
>      const tableValues = actionRows.map(row => [row.Action, row.N, row.Percent]);
>      newTable.addRows(-1, tableValues);
>    }
>    
>    interface ActionRow {
>      Action: string;
>      N: number;
>      Percent: number;
>    }
> ```

### <a name="get-json-data-from-external-sources"></a>外部ソースから JSON データを取得する

Office スクリプトを使用して JSON データをブックにインポートするには、2 つの方法があります。

- Power Automate フローを持つ[パラメーター](power-automate-integration.md#main-parameters-pass-data-to-a-script)として。
- `fetch` [外部 Web サービス](external-calls.md)を呼び出す場合。

#### <a name="modify-the-sample-to-work-with-power-automate"></a>Power Automateを操作するようにサンプルを変更する

Power Automateの JSON データは、ジェネリック オブジェクト配列として渡すことができます。 そのデータを `object[]` 受け入れるプロパティをスクリプトに追加します。

```typescript
// For Power Automate, replace the main signature in the previous sample with this one
// and remove the sample data.
function main(workbook: ExcelScript.Workbook, jsonData: object[]) {
```

次に、**スクリプトの実行** アクションに追加`jsonData`するオプションがPower Automate コネクタに表示されます。

:::image type="content" source="../images/json-parameter-power-automate.png" alt-text="jsonData パラメーターを使用したスクリプトの実行アクションを示すExcel Online (Business) コネクタ。":::

#### <a name="modify-the-sample-to-use-a-fetch-call"></a>呼び出しを使用するようにサンプルを変更する`fetch`

Web サービスは、JSON データを使用して呼び出しに `fetch` 応答できます。 これにより、スクリプトに必要なデータがExcelされます。 [Office スクリプトの外部 API 呼び出しのサポートを参照して、外部呼び出しの詳細](external-calls.md)`fetch`と外部呼び出しについて説明します。

```typescript
// For external services, replace the main signature in the previous sample with this one,
// add the fetch call, and remove the sample data.
async function main(workbook: ExcelScript.Workbook) {
  // Replace WEB_SERVICE_URL with the URL of whatever service you need to call.
  const response = await fetch('WEB_SERVICE_URL');
  const jsonData: object[] = await response.json();
```

## <a name="create-json-from-a-range"></a>範囲から JSON を作成する

ワークシートの行と列は、多くの場合、データ値間のリレーションシップを意味します。 テーブルの行は概念的にプログラミング オブジェクトにマップされ、各列はそのオブジェクトのプロパティになります。 次の表のデータを検討してください。 各行は、スプレッドシートに記録されたトランザクションを表します。

|ID |日付     |Amount |ベンダー                        |
|:--|:--------|:------|:-----------------------------|
|1  |6/1/2022 |$43.54 |お客様に最適な Organics Company |
|2  |6/3/2022 |$67.23 |自由のパン屋とカフェ       |
|3  |6/3/2022 |$37.12 |お客様に最適な Organics Company |
|4  |6/6/2022 |$86.95 |Coho Vineyard                 |
|5  |6/7/2022 |$13.64 |自由のパン屋とカフェ       |

各トランザクション (各行) には、"ID"、"Date"、"Amount"、"Vendor" という一連のプロパティが関連付けられます。 これは、Office スクリプトでオブジェクトとしてモデル化できます。

```typescript
// An interface that wraps transaction details as JSON.
interface Transaction {
  "ID": string;
  "Date": number;
  "Amount": number;
  "Vendor": string;
}
```

サンプル テーブルの行はインターフェイスのプロパティに対応するため、スクリプトは各行をオブジェクトに簡単に `Transaction` 変換できます。 これは、Power Automateのデータを出力するときに便利です。 次のスクリプトは、テーブル内の各行を反復処理し、それを `Transaction[]`.

```typescript
function main(workbook: ExcelScript.Workbook) {
  // Get the table on the current worksheet.
  const table = workbook.getActiveWorksheet().getTables()[0];

  // Create an array of Transactions and add each row to it.
  let transactions: Transaction[] = [];
  const dataValues = table.getRangeBetweenHeaderAndTotal().getValues();
  for (let i = 0; i < dataValues.length; i++) {
    let row = dataValues[i];
    let currentTransaction: Transaction = {
      ID: row[table.getColumnByName("ID").getIndex()] as string,
      Date: row[table.getColumnByName("Date").getIndex()] as number,
      Amount: row[table.getColumnByName("Amount").getIndex()] as number,
      Vendor: row[table.getColumnByName("Vendor").getIndex()] as string
    };
    transactions.push(currentTransaction);
  }

  // Do something with the Transaction objects, such as return them to a Power Automate flow.
  console.log(transactions);
}

// An interface that wraps transaction details as JSON.
interface Transaction {
  "ID": string;
  "Date": number;
  "Amount": number;
  "Vendor": string;
}
```

:::image type="content" source="../images/create-json-console-output.png" alt-text="オブジェクトのプロパティ値を示す前のスクリプトからのコンソール出力。":::

### <a name="use-a-generic-object"></a>ジェネリック オブジェクトを使用する

前のサンプルでは、テーブル ヘッダーの値が一貫性があることを前提としています。 テーブルに変数列がある場合は、汎用 JSON オブジェクトを作成する必要があります。 次のスクリプトは、任意のテーブルを JSON としてログに記録するスクリプトを示しています。

```typescript
function main(workbook: ExcelScript.Workbook) {
  // Get the table on the current worksheet.
  const table = workbook.getActiveWorksheet().getTables()[0];

  // Use the table header names as JSON properties.
  const tableHeaders = table.getHeaderRowRange().getValues()[0] as string[];
  
  // Get each data row in the table.
  const dataValues = table.getRangeBetweenHeaderAndTotal().getValues();
  let jsonArray: object[] = [];

  // For each row, create a JSON object and assign each property to it based on the table headers.
  for (let i = 0; i < dataValues.length; i++) {
    // Create a blank generic JSON object.
    let jsonObject: { [key: string]: string } = {};
    for (let j = 0; j < dataValues[i].length; j++) {
      jsonObject[tableHeaders[j]] = dataValues[i][j] as string;
    }

    jsonArray.push(jsonObject);
  }

  // Do something with the objects, such as return them to a Power Automate flow.
  console.log(jsonArray);
}

```

## <a name="see-also"></a>関連項目

- [Office スクリプトでの外部 API 呼び出しのサポート](external-calls.md)
- [サンプル: Office スクリプトで外部フェッチ呼び出しを使用する](../resources/samples/external-fetch-calls.md)
- [Power Automate を使用した Office スクリプトの実行](power-automate-integration.md)