---
title: ブックに画像を追加する
description: '[スクリプト] を使用してOfficeをブックに追加し、シート間でコピーする方法について学習します。'
ms.date: 07/12/2021
localization_priority: Normal
ms.openlocfilehash: 993444aa328356f872db90d1b9d2403bf28be4de
ms.sourcegitcommit: a86b91c7e104bb7c26efd56de53b9e3976a34828
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 07/12/2021
ms.locfileid: "53394417"
---
# <a name="add-images-to-a-workbook"></a><span data-ttu-id="bcd40-103">ブックに画像を追加する</span><span class="sxs-lookup"><span data-stu-id="bcd40-103">Add images to a workbook</span></span>

<span data-ttu-id="bcd40-104">このサンプルでは、スクリプト内のスクリプトを使用してイメージをOffice方法をExcel。</span><span class="sxs-lookup"><span data-stu-id="bcd40-104">This sample shows how to work with images using an Office Script in Excel.</span></span>

## <a name="scenario"></a><span data-ttu-id="bcd40-105">シナリオ</span><span class="sxs-lookup"><span data-stu-id="bcd40-105">Scenario</span></span>

<span data-ttu-id="bcd40-106">画像は、ブランド化、ビジュアル ID、テンプレートに役立ちます。</span><span class="sxs-lookup"><span data-stu-id="bcd40-106">Images help with branding, visual identity, and templates.</span></span> <span data-ttu-id="bcd40-107">これらは、単なる巨大なテーブル以外のブックを作成するのに役立ちます。</span><span class="sxs-lookup"><span data-stu-id="bcd40-107">They help make a workbook more than just a giant table.</span></span>

<span data-ttu-id="bcd40-108">最初のサンプルでは、あるワークシートから別のワークシートにイメージをコピーします。</span><span class="sxs-lookup"><span data-stu-id="bcd40-108">The first sample copies an image from one worksheet to another.</span></span> <span data-ttu-id="bcd40-109">これは、会社のロゴをすべてのシートで同じ位置に配置するために使用できます。</span><span class="sxs-lookup"><span data-stu-id="bcd40-109">This could be used to put your company's logo in the same position on every sheet.</span></span>

<span data-ttu-id="bcd40-110">2 番目のサンプルでは、URL からイメージをコピーします。</span><span class="sxs-lookup"><span data-stu-id="bcd40-110">The second sample copies an image from a URL.</span></span> <span data-ttu-id="bcd40-111">これは、同僚が共有フォルダーに保存した写真を関連するブックにコピーするために使用できます。</span><span class="sxs-lookup"><span data-stu-id="bcd40-111">This could be used to copy photos that a colleague stored in a shared folder to a related workbook.</span></span>

## <a name="sample-excel-file"></a><span data-ttu-id="bcd40-112">サンプル Excel ファイル</span><span class="sxs-lookup"><span data-stu-id="bcd40-112">Sample Excel file</span></span>

<span data-ttu-id="bcd40-113">すぐに <a href="add-images.xlsx"> 使用add-images.xlsx</a> ブックのブックをダウンロードします。</span><span class="sxs-lookup"><span data-stu-id="bcd40-113">Download <a href="add-images.xlsx">add-images.xlsx</a> for a ready-to-use workbook.</span></span> <span data-ttu-id="bcd40-114">次のスクリプトを追加し、自分でサンプルを試してみてください。</span><span class="sxs-lookup"><span data-stu-id="bcd40-114">Add the following scripts and try the sample yourself!</span></span>

## <a name="sample-code-copy-an-image-across-worksheets"></a><span data-ttu-id="bcd40-115">サンプル コード: ワークシート間で画像をコピーする</span><span class="sxs-lookup"><span data-stu-id="bcd40-115">Sample code: Copy an image across worksheets</span></span>

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

## <a name="sample-code-add-an-image-from-a-url-to-a-workbook"></a><span data-ttu-id="bcd40-116">サンプル コード: URL からブックにイメージを追加する</span><span class="sxs-lookup"><span data-stu-id="bcd40-116">Sample code: Add an image from a URL to a workbook</span></span>

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
