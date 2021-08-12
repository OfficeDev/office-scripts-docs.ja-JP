---
title: ブックに画像を追加する
description: '[スクリプト] を使用してOfficeをブックに追加し、シート間でコピーする方法について学習します。'
ms.date: 07/12/2021
localization_priority: Normal
ms.openlocfilehash: 7f8442ccfceb701fcca816d4203205dddabe68d9dfe594c87363c7eb42f353dd
ms.sourcegitcommit: 75f7ed8c2d23a104acc293f8ce29ea580b4fcdc5
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 08/11/2021
ms.locfileid: "57846315"
---
# <a name="add-images-to-a-workbook"></a>ブックに画像を追加する

このサンプルでは、スクリプト内のスクリプトを使用してイメージをOffice方法をExcel。

## <a name="scenario"></a>シナリオ

画像は、ブランド化、ビジュアル ID、テンプレートに役立ちます。 これらは、単なる巨大なテーブル以外のブックを作成するのに役立ちます。

最初のサンプルでは、あるワークシートから別のワークシートにイメージをコピーします。 これは、会社のロゴをすべてのシートで同じ位置に配置するために使用できます。

2 番目のサンプルでは、URL からイメージをコピーします。 これは、同僚が共有フォルダーに保存した写真を関連するブックにコピーするために使用できます。

## <a name="sample-excel-file"></a>サンプル Excel ファイル

すぐに <a href="add-images.xlsx"> 使用add-images.xlsx</a> ブックのブックをダウンロードします。 次のスクリプトを追加し、自分でサンプルを試してみてください。

## <a name="sample-code-copy-an-image-across-worksheets"></a>サンプル コード: ワークシート間で画像をコピーする

```TypeScript
/**
 * This script transfers an image from one worksheet to another.
 */
function main(workbook: ExcelScript.Workbook)
{
  // Get the worksheet with the image on it.
  let firstWorksheet = workbook.getWorksheet("FirstSheet");

  // Get the first image from the worksheet.
  // If a script added the image, you could add a name to make it easier to find.
  let image: ExcelScript.Image;
  firstWorksheet.getShapes().forEach((shape, index) => {
    if (shape.getType() === ExcelScript.ShapeType.image) {
      image = shape.getImage();
      return;
    }
  });

  // Copy the image to another worksheet.
  image.getShape().copyTo("SecondSheet");
}
```

## <a name="sample-code-add-an-image-from-a-url-to-a-workbook"></a>サンプル コード: URL からブックにイメージを追加する

```TypeScript
async function main(workbook: ExcelScript.Workbook) {
  // Fetch the image from a URL.
  const link = "https://raw.githubusercontent.com/OfficeDev/office-scripts-docs/master/docs/images/git-octocat.png";
  const response = await fetch(link);

  // Store the response as an ArrayBuffer, since it is a raw image file.
  const data = await response.arrayBuffer();

  // Convert the image data into a base64-encoded string.
  const image = convertToBase64(data);

  // Add the image to a worksheet.
  workbook.getWorksheet("WebSheet").addImage(image);
}

/**
 * Converts an ArrayBuffer containing a .png image into a base64-encoded string.
 */
function convertToBase64(input: ArrayBuffer) {
  const uInt8Array = new Uint8Array(input);
  const count = uInt8Array.length;

  // Allocate the necessary space up front.
  const charCodeArray = new Array(count) as string[];
  
  // Convert every entry in the array to a character.
  for (let i = count; i >= 0; i--) { 
    charCodeArray[i] = String.fromCharCode(uInt8Array[i]);
  }

  // Convert the characters to base64.
  const base64 = btoa(charCodeArray.join(''));
  return base64;
}
```
