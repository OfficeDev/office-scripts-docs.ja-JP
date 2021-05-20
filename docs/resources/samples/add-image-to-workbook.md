---
title: ブックにイメージを追加する
description: Office スクリプトを使用して、ワークブックに画像を追加し、シート間でコピーする方法について説明します。
ms.date: 05/17/2021
localization_priority: Normal
ms.openlocfilehash: 99c3cc2cacf6e535bdb882bb8414d23fd105be35
ms.sourcegitcommit: 4687693f02fc90a57ba30c461f35046e02e6f5fb
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 05/19/2021
ms.locfileid: "52546038"
---
# <a name="add-images-to-a-workbook"></a><span data-ttu-id="2e709-103">ブックにイメージを追加する</span><span class="sxs-lookup"><span data-stu-id="2e709-103">Add images to a workbook</span></span>

<span data-ttu-id="2e709-104">このサンプルでは、ExcelのOffice スクリプトを使用してイメージを操作する方法を示します。</span><span class="sxs-lookup"><span data-stu-id="2e709-104">This sample shows how to work with images using an Office Script in Excel.</span></span>

## <a name="scenario"></a><span data-ttu-id="2e709-105">シナリオ</span><span class="sxs-lookup"><span data-stu-id="2e709-105">Scenario</span></span>

<span data-ttu-id="2e709-106">イメージは、ブランド化、視覚的な ID、テンプレートに役立ちます。</span><span class="sxs-lookup"><span data-stu-id="2e709-106">Images help with branding, visual identity, and templates.</span></span> <span data-ttu-id="2e709-107">彼らは単なる巨大なテーブル以上のワークブックを作るのに役立ちます。</span><span class="sxs-lookup"><span data-stu-id="2e709-107">They help make a workbook more than just a giant table.</span></span>

<span data-ttu-id="2e709-108">最初のサンプルでは、あるワークシートから別のワークシートにイメージをコピーします。</span><span class="sxs-lookup"><span data-stu-id="2e709-108">The first sample copies an image from one worksheet to another.</span></span> <span data-ttu-id="2e709-109">これは、すべてのシート上で会社のロゴを同じ位置に配置するために使用できます。</span><span class="sxs-lookup"><span data-stu-id="2e709-109">This could be used to put your company's logo in the same position on every sheet.</span></span>

<span data-ttu-id="2e709-110">2 番目のサンプルでは、URL からイメージをコピーします。</span><span class="sxs-lookup"><span data-stu-id="2e709-110">The second sample copies an image from a URL.</span></span> <span data-ttu-id="2e709-111">この機能を使用して、同僚が共有フォルダに保存した写真を関連するブックにコピーできます。</span><span class="sxs-lookup"><span data-stu-id="2e709-111">This could be used to copy photos that a colleague stored in a shared folder to a related workbook.</span></span>

## <a name="sample-excel-file"></a><span data-ttu-id="2e709-112">サンプル Excel ファイル</span><span class="sxs-lookup"><span data-stu-id="2e709-112">Sample Excel file</span></span>

<span data-ttu-id="2e709-113">これらのサンプルで使用 <a href="add-images.xlsx">add-images.xlsx</a> ファイルをダウンロードして、自分で試してみてください!</span><span class="sxs-lookup"><span data-stu-id="2e709-113">Download the file <a href="add-images.xlsx">add-images.xlsx</a> used in these samples and try it out yourself!</span></span>

## <a name="sample-code-copy-an-image-across-worksheets"></a><span data-ttu-id="2e709-114">サンプル コード: ワークシート間でイメージをコピーする</span><span class="sxs-lookup"><span data-stu-id="2e709-114">Sample code: Copy an image across worksheets</span></span>

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

## <a name="sample-code-add-an-image-from-a-url-to-a-workbook"></a><span data-ttu-id="2e709-115">サンプル コード: URL からブックにイメージを追加する</span><span class="sxs-lookup"><span data-stu-id="2e709-115">Sample code: Add an image from a URL to a workbook</span></span>

```TypeScript
async function main(workbook: ExcelScript.Workbook) {
  // Fetch the image from a URL.
  const link = "https://github.com/OfficeDev/office-scripts-docs/blob/master/docs/images/git-octocat.png";
  const response = await fetch(link);

  // Store the response as an ArrayBuffer, since it is a raw image file.
  const data = await response.arrayBuffer();

  // Convert the image data into a base64-encoded string.
  const image = convertToBase64(data);

  // Add the image to a worksheet.
  workbook.getWorksheet("WebSheet").addImage(image)
}

/**
 * Converts an ArrayBuffer containing a .png image into a base64-encoded string.
 */
function convertToBase64(input: ArrayBuffer) {
  const uInt8Array = new Uint8Array(input);
  const count = uInt8Array.length;

  // Allocate the necessary space up front.
  const charCodeArray = new Array(count) 
  
  // Convert every entry in the array to a character.
  for (let i = count; i >= 0; i--) { 
    charCodeArray[i] = String.fromCharCode(uInt8Array[i]);
  }

  // Convert the characters to base64.
  const base64 = btoa(charCodeArray.join(''));
  return base64;
}
```
