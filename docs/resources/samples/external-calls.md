---
title: スクリプトで外部 API 呼び出Officeする
description: スクリプトで外部 API 呼び出しを行うOfficeします。
ms.date: 03/30/2021
localization_priority: Normal
ms.openlocfilehash: 0ed57ed3b97309dbb7ea196695dcc347e133b3cf
ms.sourcegitcommit: 45ffe3dbd2c834b78592ad35928cf8096f5e80bc
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 04/14/2021
ms.locfileid: "51754804"
---
# <a name="external-api-calls-from-office-scripts"></a>スクリプトからの外部 API Office呼び出し

Officeスクリプトを使用すると、 [外部 API 呼び出しのサポートが制限されます](../../develop/external-calls.md)。

> [!IMPORTANT]
>
> * OAuth2 タイプの認証フローをサインインまたは使用する方法はありません。 すべてのキーと資格情報をハードコード (または別のソースから読み取る) 必要があります。
> * API の資格情報とキーを格納するインフラストラクチャはありません。 これは、ユーザーが管理する必要があります。
> * 外部呼び出しにより、機密データが望ましくないエンドポイントに公開される場合や、内部ブックに外部データが取り込まれたりする場合があります。 管理者は、このような呼び出しに対するファイアウォール保護を確立できます。 外部通話に依存する前に、必ずローカル ポリシーに確認してください。
> * スクリプトが API 呼び出しを使用する場合、Power Automate シナリオでは機能しません。 Power Automate の HTTP アクションまたは同等のアクションを使用して、データを外部サービスから取得または外部サービスにプッシュする必要があります。
> * 外部 API 呼び出しには非同期 API 構文が含まれるので、非同期通信の仕組みについて少し高度な知識が必要です。
> * 依存関係を取得する前に、データ スループットの量を確認してください。 たとえば、外部データセット全体を引き下げないのが最適な選択肢ではなく、代わりにページネーションを使用してデータをチャンク単位で取得する必要があります。

## <a name="useful-knowledge-and-resources"></a>有用な知識とリソース

* [REST API](https://en.wikipedia.org/wiki/Representational_state_transfer): ほとんどの場合、API 呼び出しを使用する方法です。
* [ `async` : この動作を理解します `await` ](https://developer.mozilla.org/docs/Learn/JavaScript/Asynchronous/Async_await)。
* [`fetch`](https://developer.mozilla.org/docs/Web/API/Fetch_API/Using_Fetch): この動作を理解します。

## <a name="steps"></a>手順

1. プレフィックスを `main` 追加して、関数を非同期関数としてマーク `async` します。 たとえば、`async function main(workbook: ExcelScript.Workbook)` などです。
1. どの種類の API 呼び出しを行っていますか? `GET`, `POST`, `PUT`, `DELETE`, `PATCH`? 詳細については、REST API の資料を参照してください。
1. サービス API エンドポイント、認証要件、ヘッダーなどを取得します。
1. コードの完了と開発時間の検証に役立つ入力または `interface` 出力を定義します。 詳細については [、ビデオ](#training-video-how-to-make-external-api-calls) を参照してください。
1. コード、テスト、最適化。 API 呼び出しルーチンの関数を作成して、スクリプトの他の部分から再利用可能にしたり、別のスクリプトで再利用したりすることができます (コピー貼り付けは、この方法ではるかに簡単になります)。

## <a name="scenario"></a>シナリオ

このスクリプトは、ユーザーの GitHub リポジトリに関する基本情報を取得します。

## <a name="resources-used-in-the-sample"></a>サンプルで使用されるリソース

1. [リポジトリ Github API リファレンスを取得します。](https://docs.github.com/rest/reference/repos#list-repositories-for-a-user)
1. API 呼び出しの出力: Web ブラウザーまたは任意の HTTP インターフェイスに移動して入力し `https://api.github.com/users/{USERNAME}/repos` 、{USERNAME} プレースホルダーを Github ID に置き換える。
1. 取得される情報: repo.name、repo.size、repo.owner.id、repo.license?。name

## <a name="sample-code-get-basic-information-about-users-github-repositories"></a>サンプル コード: ユーザーの GitHub リポジトリに関する基本情報を取得する

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

## <a name="training-video-how-to-make-external-api-calls"></a>トレーニング ビデオ: 外部 API 呼び出しを行う方法

[![外部 API 呼び出しの実行方法に関するビデオを見る](../../images/api-vid.png)](https://youtu.be/fulP29J418E "外部 API 呼び出しを行う方法に関するビデオ")
