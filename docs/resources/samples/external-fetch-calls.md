---
title: Office スクリプトで外部取得呼び出しを使用する
description: Office スクリプトで外部 API 呼び出しを行う方法について説明します。
ms.date: 06/10/2022
ms.localizationpriority: medium
ms.openlocfilehash: 569d74f1ca8996cd8fe8a4ba3163445d57676d27
ms.sourcegitcommit: dd01979d34b3499360d2f79a56f8a8f24f480eed
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 06/15/2022
ms.locfileid: "66088093"
---
# <a name="use-external-fetch-calls-in-office-scripts"></a>Office スクリプトで外部取得呼び出しを使用する

このスクリプトは、ユーザーのGitHub リポジトリに関する基本情報を取得します。 簡単なシナリオで使用 `fetch` する方法を示します。 その他の外部呼び出しの使用`fetch`の詳細については、[Office スクリプトの外部 API 呼び出しのサポートに関するページを](../../develop/external-calls.md)参照してください。 [JSON]](https://www.w3schools.com/whatis/whatis_json.asp) オブジェクトの操作については、GitHub API によって返されるものなど)、「[JSON を使用して、Office スクリプトとの間でデータを渡す](../../develop/use-json.md)」を参照してください。

[GitHub API リファレンス](https://docs.github.com/rest/reference/repos#list-repositories-for-a-user)で使用されている GItHub API の詳細について説明します。 また、Web ブラウザーにアクセス`https://api.github.com/users/{USERNAME}/repos`して生の API 呼び出しの出力を確認することもできます (必ず、{USERNAME} プレースホルダーをGitHub ID に置き換えてください)。

![リポジトリ情報の取得の例](../../images/git.png)

## <a name="sample-code-get-basic-information-about-users-github-repositories"></a>サンプル コード: ユーザーのGitHub リポジトリに関する基本情報を取得する

```TypeScript
async function main(workbook: ExcelScript.Workbook) {
  // Call the GitHub REST API.
  // Replace the {USERNAME} placeholder with your GitHub username.
  const response = await fetch('https://api.github.com/users/{USERNAME}/repos');
  const repos: Repository[] = await response.json();

  // Create an array to hold the returned values.
  const rows: (string | boolean | number)[][] = [];

  // Convert each repository block into a row.
  for (let repo of repos) {
    rows.push([repo.id, repo.name, repo.license?.name, repo.license?.url]);
  }
  // Create a header row.
  const sheet = workbook.getActiveWorksheet();
  sheet.getRange('A1:D1').setValues([["ID", "Name", "License Name", "License URL"]]);

  // Add the data to the current worksheet, starting at "A2".
  const range = sheet.getRange('A2').getResizedRange(rows.length - 1, rows[0].length - 1);
  range.setValues(rows);
}

// An interface matching the returned JSON for a GitHub repository.
interface Repository {
  name: string,
  id: string,
  license?: License
}

// An interface matching the returned JSON for a GitHub repo license.
interface License {
  name: string,
  url: string
}
```

## <a name="training-video-how-to-make-external-api-calls"></a>トレーニング ビデオ: 外部 API 呼び出しを行う方法

[YouTube でこのサンプルを見る、スディ Ramamurthy のチュートリアルをご覧ください](https://youtu.be/fulP29J418E)。
