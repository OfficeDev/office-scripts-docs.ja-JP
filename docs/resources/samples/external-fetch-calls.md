---
title: スクリプトで外部フェッチ呼び出しOfficeする
description: スクリプトで外部 API 呼び出しを行うOfficeします。
ms.date: 04/05/2021
localization_priority: Normal
ms.openlocfilehash: a77ceb61c2ff46a7b6226b798462b7be2c8e1c54
ms.sourcegitcommit: 1f003c9924e651600c913d84094506125f1055ab
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 04/26/2021
ms.locfileid: "52026995"
---
# <a name="use-external-fetch-calls-in-office-scripts"></a><span data-ttu-id="81d4d-103">スクリプトで外部フェッチ呼び出しOfficeする</span><span class="sxs-lookup"><span data-stu-id="81d4d-103">Use external fetch calls in Office Scripts</span></span>

<span data-ttu-id="81d4d-104">このスクリプトは、ユーザーのリポジトリに関するGitHub取得します。</span><span class="sxs-lookup"><span data-stu-id="81d4d-104">This script gets basic information about a user's GitHub repositories.</span></span> <span data-ttu-id="81d4d-105">単純なシナリオでの使 `fetch` い方を示します。</span><span class="sxs-lookup"><span data-stu-id="81d4d-105">It shows how to use `fetch` in a simple scenario.</span></span>

<span data-ttu-id="81d4d-106">使用されている GItHub API の詳細については、「API リファレンス[」GitHub参照してください](https://docs.github.com/rest/reference/repos#list-repositories-for-a-user)。</span><span class="sxs-lookup"><span data-stu-id="81d4d-106">You can learn more about the GItHub APIs being used in the [GitHub API reference](https://docs.github.com/rest/reference/repos#list-repositories-for-a-user).</span></span> <span data-ttu-id="81d4d-107">Web ブラウザーにアクセスして、生の API 呼び出しの出力を確認することもできます ({USERNAME} プレースホルダーを Github ID に置き `https://api.github.com/users/{USERNAME}/repos` 換えてください)。</span><span class="sxs-lookup"><span data-stu-id="81d4d-107">You can also see the raw API call output by visiting `https://api.github.com/users/{USERNAME}/repos` in a web browser (be sure to replace the {USERNAME} placeholder with your Github ID).</span></span>

![リポジトリ情報の取得例](../../images/git.png)

## <a name="sample-code-get-basic-information-about-users-github-repositories"></a><span data-ttu-id="81d4d-109">サンプル コード: ユーザーのリポジトリに関するGitHub取得する</span><span class="sxs-lookup"><span data-stu-id="81d4d-109">Sample code: Get basic information about user's GitHub repositories</span></span>

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

## <a name="training-video-how-to-make-external-api-calls"></a><span data-ttu-id="81d4d-110">トレーニング ビデオ: 外部 API 呼び出しを行う方法</span><span class="sxs-lookup"><span data-stu-id="81d4d-110">Training video: How to make external API calls</span></span>

<span data-ttu-id="81d4d-111">[![外部 API 呼び出しの実行方法に関するビデオを見る](../../images/api-vid.png)](https://youtu.be/fulP29J418E "外部 API 呼び出しを行う方法に関するビデオ")</span><span class="sxs-lookup"><span data-stu-id="81d4d-111">[![Watch video on how to make external API calls](../../images/api-vid.png)](https://youtu.be/fulP29J418E "Video on how to make external API calls")</span></span>
