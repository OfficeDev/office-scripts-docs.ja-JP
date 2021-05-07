---
title: Office スクリプトで外部取得呼び出しを使用する
description: スクリプトで外部 API 呼び出しを行うOfficeします。
ms.date: 04/28/2021
localization_priority: Normal
ms.openlocfilehash: 721bfa39eea1e9973efc7fd13efa5bac734b76dd
ms.sourcegitcommit: f7a7aebfb687f2a35dbed07ed62ff352a114525a
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 05/06/2021
ms.locfileid: "52232523"
---
# <a name="use-external-fetch-calls-in-office-scripts"></a><span data-ttu-id="57c21-103">Office スクリプトで外部取得呼び出しを使用する</span><span class="sxs-lookup"><span data-stu-id="57c21-103">Use external fetch calls in Office Scripts</span></span>

<span data-ttu-id="57c21-104">このスクリプトは、ユーザーのリポジトリに関するGitHub取得します。</span><span class="sxs-lookup"><span data-stu-id="57c21-104">This script gets basic information about a user's GitHub repositories.</span></span> <span data-ttu-id="57c21-105">単純なシナリオでの使 `fetch` い方を示します。</span><span class="sxs-lookup"><span data-stu-id="57c21-105">It shows how to use `fetch` in a simple scenario.</span></span>

<span data-ttu-id="57c21-106">使用されている GItHub API の詳細については、「API リファレンス[」GitHub参照してください](https://docs.github.com/rest/reference/repos#list-repositories-for-a-user)。</span><span class="sxs-lookup"><span data-stu-id="57c21-106">You can learn more about the GItHub APIs being used in the [GitHub API reference](https://docs.github.com/rest/reference/repos#list-repositories-for-a-user).</span></span> <span data-ttu-id="57c21-107">Web ブラウザーにアクセスして、生の API 呼び出しの出力を確認することもできます ({USERNAME} プレースホルダーを Github ID に置き `https://api.github.com/users/{USERNAME}/repos` 換えてください)。</span><span class="sxs-lookup"><span data-stu-id="57c21-107">You can also see the raw API call output by visiting `https://api.github.com/users/{USERNAME}/repos` in a web browser (be sure to replace the {USERNAME} placeholder with your Github ID).</span></span>

![リポジトリ情報の取得例](../../images/git.png)

## <a name="sample-code-get-basic-information-about-users-github-repositories"></a><span data-ttu-id="57c21-109">サンプル コード: ユーザーのリポジトリに関するGitHub取得する</span><span class="sxs-lookup"><span data-stu-id="57c21-109">Sample code: Get basic information about user's GitHub repositories</span></span>

```TypeScript
async function main(workbook: ExcelScript.Workbook) {

  // Replace the {USERNAME} placeholder with your GitHub username.
  const response = await fetch('https://api.github.com/users/{USERNAME}/repos');
  const repos: Repository[] = await response.json();
  
  const rows: (string | boolean | number)[][] = [];
  for (let repo of repos){ 
    rows.push([repo.id, repo.name, repo.license?.name, repo.license?.url])
  }
  const sheet = workbook.getActiveWorksheet();
  const range = sheet.getRange('A2').getResizedRange(rows.length - 1, rows[0].length - 1);
  range.setValues(rows);
  return;
}

interface Repository {
  name: string,
  id: string,
  license?: License 
}

interface License {
  name: string,
  url: string
}
```

## <a name="training-video-how-to-make-external-api-calls"></a><span data-ttu-id="57c21-110">トレーニング ビデオ: 外部 API 呼び出しを行う方法</span><span class="sxs-lookup"><span data-stu-id="57c21-110">Training video: How to make external API calls</span></span>

<span data-ttu-id="57c21-111">[Sudhi Ramamurthy が YouTube でこのサンプルを歩くのを見る](https://youtu.be/fulP29J418E).</span><span class="sxs-lookup"><span data-stu-id="57c21-111">[Watch Sudhi Ramamurthy walk through this sample on YouTube](https://youtu.be/fulP29J418E).</span></span>
