---
title: Office スクリプトで外部取得呼び出しを使用する
description: Officeスクリプトで外部 API 呼び出しを行う方法について説明します。
ms.date: 05/14/2021
localization_priority: Normal
ms.openlocfilehash: df8814cbab16969a1140aecfe526fd68e609d43c
ms.sourcegitcommit: 4687693f02fc90a57ba30c461f35046e02e6f5fb
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 05/19/2021
ms.locfileid: "52545753"
---
# <a name="use-external-fetch-calls-in-office-scripts"></a><span data-ttu-id="a6b81-103">Office スクリプトで外部取得呼び出しを使用する</span><span class="sxs-lookup"><span data-stu-id="a6b81-103">Use external fetch calls in Office Scripts</span></span>

<span data-ttu-id="a6b81-104">このスクリプトは、ユーザのGitHubリポジトリに関する基本情報を取得します。</span><span class="sxs-lookup"><span data-stu-id="a6b81-104">This script gets basic information about a user's GitHub repositories.</span></span> <span data-ttu-id="a6b81-105">`fetch`簡単なシナリオでの使用方法を示します。</span><span class="sxs-lookup"><span data-stu-id="a6b81-105">It shows how to use `fetch` in a simple scenario.</span></span> <span data-ttu-id="a6b81-106">使用 `fetch` または他の外部呼び出しの詳細については[、「Officeスクリプト」の外部 API コールサポートを参照してください。](../../develop/external-calls.md)</span><span class="sxs-lookup"><span data-stu-id="a6b81-106">For more information about using `fetch` or other external calls, read [External API call support in Office Scripts](../../develop/external-calls.md)</span></span>

<span data-ttu-id="a6b81-107">使用されている GItHub API の詳細については[、「GitHub API リファレンス」を参照してください](https://docs.github.com/rest/reference/repos#list-repositories-for-a-user)。</span><span class="sxs-lookup"><span data-stu-id="a6b81-107">You can learn more about the GItHub APIs being used in the [GitHub API reference](https://docs.github.com/rest/reference/repos#list-repositories-for-a-user).</span></span> <span data-ttu-id="a6b81-108">Web ブラウザーでアクセスして生の API 呼び出しの出力を確認することもできます `https://api.github.com/users/{USERNAME}/repos` ({USERNAME} プレースホルダーをGitHub ID に置き換えてください)。</span><span class="sxs-lookup"><span data-stu-id="a6b81-108">You can also see the raw API call output by visiting `https://api.github.com/users/{USERNAME}/repos` in a web browser (be sure to replace the {USERNAME} placeholder with your GitHub ID).</span></span>

![リポジトリ情報の取得の例](../../images/git.png)

## <a name="sample-code-get-basic-information-about-users-github-repositories"></a><span data-ttu-id="a6b81-110">サンプル コード: ユーザーのGitHub リポジトリに関する基本情報を取得する</span><span class="sxs-lookup"><span data-stu-id="a6b81-110">Sample code: Get basic information about user's GitHub repositories</span></span>

```TypeScript
async function main(workbook: ExcelScript.Workbook) {
  // Call the GitHub REST API.
  // Replace the {USERNAME} placeholder with your GitHub username.
  const response = await fetch('https://api.github.com/users/{USERNAME}/repos');
  const repos: Repository[] = await response.json();
  
  // Create an array to hold the returned values.
  const rows: (string | boolean | number)[][] = [];

  // Convert each repository block into a row.
  for (let repo of repos){ 
    rows.push([repo.id, repo.name, repo.license?.name, repo.license?.url])
  }

  // Add the data to the current worksheet, starting at "A2".
  const sheet = workbook.getActiveWorksheet();
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

## <a name="training-video-how-to-make-external-api-calls"></a><span data-ttu-id="a6b81-111">トレーニング ビデオ: 外部 API 呼び出しを行う方法</span><span class="sxs-lookup"><span data-stu-id="a6b81-111">Training video: How to make external API calls</span></span>

<span data-ttu-id="a6b81-112">[スーディ・ラマムルティがこのサンプルをYouTubeで歩くのを見てください](https://youtu.be/fulP29J418E)。</span><span class="sxs-lookup"><span data-stu-id="a6b81-112">[Watch Sudhi Ramamurthy walk through this sample on YouTube](https://youtu.be/fulP29J418E).</span></span>
