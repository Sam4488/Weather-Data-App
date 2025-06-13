import ExcelJS from "exceljs"
import { saveAs } from "file-saver"

export async function injectHeaderToExcel() {
  const workbook = new ExcelJS.Workbook()

  // Load the template file from /public/template.xlsx
  const templateRes = await fetch("/template.xlsx")
  const templateBlob = await templateRes.blob()
  const templateBuffer = await templateBlob.arrayBuffer()

  // Read the existing workbook from buffer
  await workbook.xlsx.load(templateBuffer)

  const sheet = workbook.worksheets[0] // First sheet

  // Insert 5 empty rows at the top to shift all data down
  sheet.spliceRows(1, 0, [], [], [], [], [])

  // Merge first 5 rows and 3 columns for left image (A1:C5)
  sheet.mergeCells("A1:C5")

  // Load the left image from /public/header.jpg
  const imageRes = await fetch("/header.jpg")
  const imageBlob = await imageRes.blob()
  const imageBuffer = await imageBlob.arrayBuffer()

  // Add left image to workbook
  const imageId = workbook.addImage({
    buffer: imageBuffer,
    extension: "jpeg",
  })

  // Insert left image to cover merged area (A1:C5)
  sheet.addImage(imageId, {
    tl: { col: 0, row: 0 },
    br: { col: 3, row: 5 }, // Covers A1:C5
    editAs: "oneCell"
  })


  //sheet.mergeCells("H1:J5")

 
const image2Res = await fetch("/images.png")
const image2Blob = await image2Res.blob()
const image2Buffer = await image2Blob.arrayBuffer()
// ...rest of code...

  // Add right image to workbook
  const image2Id = workbook.addImage({
    buffer: image2Buffer,
    extension: "png",
  })

  // Insert right image to cover merged area (H1:J5)
  sheet.addImage(image2Id, {
    tl: { col: 7, row: 0 }, // H is column 7 (0-based)
    br: { col: 10, row: 5 }, // J is column 9, br is exclusive so use 10
    editAs: "oneCell"
  })

  // Optionally, add header text below image
  sheet.mergeCells("D2:G2")
  sheet.getCell("D2").value = "ANDRITZ"
  sheet.getCell("D2").font = { bold: true, size: 14 }
  sheet.getCell("D2").alignment = { horizontal: "center" }

  sheet.mergeCells("D3:G5")
  sheet.getCell("D3").value = "WMS REPORT"
  sheet.getCell("D3").font = { bold: true, size: 14 }
  sheet.getCell("D3").alignment = { horizontal: "center" }

  // Export and download
  const updatedBuffer = await workbook.xlsx.writeBuffer()
  saveAs(new Blob([updatedBuffer]), "UpdatedWeatherReport.xlsx")
}