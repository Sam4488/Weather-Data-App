import React from "react"
import { injectHeaderToExcel } from "../utils/injectHeaderToExcel"

export default function HeaderExcelDownloadButton() {
  const handleClick = async () => {
    try {
      await injectHeaderToExcel()
    } catch (err) {
      console.error("Excel generation failed:", err)
    }
  }

  return (
    <button
      onClick={handleClick}
      className="px-4 py-2 bg-blue-600 text-white rounded hover:bg-blue-700"
    >
      Download Excel with Header
    </button>
  )
}
