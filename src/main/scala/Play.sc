import java.io._
import java.io.File
import java.io.FileInputStream
import java.io.FileNotFoundException
import java.io.IOException

import org.apache.poi.ss.usermodel._
import org.apache.poi.xssf.usermodel.XSSFWorkbook




object Test {



  def copyfile(): Unit = {

    val src = "C:\\work\\scala_work\\39_DWHP_Jira_17122018.xlsx"
    val dest = "C:\\work\\scala_work\\40_DWHP_Jira_17122018.xlsx"
    var inputChannel = new FileInputStream(src).getChannel()
    var outputChannel = new FileOutputStream(dest).getChannel()
    outputChannel.transferFrom(inputChannel, 0, inputChannel.size())
    inputChannel.close()
    outputChannel.close()

  }

  private val filepath = "C:\\work\\scala_work\\39_DWHP_Jira_17122018.xlsx"
  private val columnMap: Map[Int, String] = Map(1 -> "Project", 2 -> "Key", 3 -> "Summary", 4 -> "Issue Type", 5 -> "Status", 6 -> "Assignee", 7 -> "Reporter", 8 -> "Created", 9 -> "Updated", 10 -> "Due Date", 11 -> "Description", 12 -> "All Comments", 13 -> "Prioritat", 14 -> "Milestones", 15 -> "Last Comment")


  def readExcel(filepath: String): Unit = {
    try {

      val excelFile = new FileInputStream(new File(filepath))
      val wb = new XSSFWorkbook(excelFile)
      val worksheet = wb.getSheetAt(0)
      val iterator = worksheet.iterator()

      /*
            while (iterator.hasNext) {
              val currentrow = iterator.next()

              println(currentrow.getCell(1))

        */

      /*
      val cellIterator = currentrow.iterator()
      while (cellIterator.hasNext) {
        val currentCell = cellIterator.next()
        if (currentCell.getCellType == CellType.STRING) {
          println(currentCell.getStringCellValue() + "--")
        } else if (currentCell.getCellType == CellType.NUMERIC) {
          println(currentCell.getNumericCellValue() + "--")
        }
      }

    }
    */
    }
    catch {
      case ex: FileNotFoundException => ex.printStackTrace()
      case ex: IOException => ex.printStackTrace()
    }
  }

  /*
    def findDiffRows (worksheet: Iterator[String]): List[String] =
    {
      List[String]()
    }

    def updateStatus (lst: List[String]): wb = {}


    def findObsolteIssues (worksheet: Iterator[String] ): Unit = {

    }
   */
  def main(args: Array[String]): Unit = {

    val workbook = new XSSFWorkbook()
    val createHelper = workbook.getCreationHelper()
    val sheet = workbook.createSheet("DWHP")

    val headerFont = workbook.createFont()
    headerFont.setBold(true)
    headerFont.setFontHeightInPoints(12)
    headerFont.setColor(IndexedColors.BLUE_GREY.getIndex())


    val headerCellStyle = workbook.createCellStyle()
    headerCellStyle.setFont(headerFont)

    val headerRow = sheet.createRow(0)

    for ((key, value) <- columnMap) {
      val cell = headerRow.createCell(key)
      cell.setCellValue(value)
      cell.setCellStyle(headerCellStyle)
    }

    val dateCellStyle = workbook.createCellStyle()
    dateCellStyle.setDataFormat(createHelper.createDataFormat().getFormat("dd-MM-yyyy"))


    try {

      val excelFile = new FileInputStream(new File(filepath))
      val wb = new XSSFWorkbook(excelFile)
      val worksheet = wb.getSheetAt(0)
      val sheetIterator = worksheet.iterator()


      var rowNum = 0
      while (sheetIterator.hasNext) {

        val oldRow = sheetIterator.next()

        rowNum = rowNum + 1
        val row = sheet.createRow(rowNum)


        val oldRowIterator = oldRow.iterator()

        var cell = 0
        while (oldRowIterator.hasNext) {
          val oldCell = oldRowIterator.next().toString
          row.createCell(cell).setCellValue(oldCell)
          cell = cell + 1
        }

      }


      for ((key, value) <- columnMap) {
        sheet.autoSizeColumn(key)
      }

      val fileOut = new FileOutputStream("C:\\work\\scala_work\\copyDWHP.xlsx")
      workbook.write(fileOut)
      fileOut.close()

      // Closing the workbook
      workbook.close()

      excelFile.close()
      wb.close()

    }
    catch {
      case ex: FileNotFoundException => ex.printStackTrace()
      case ex: IOException => ex.printStackTrace()
    }

    }

}

Test.main(Array("bla bla"))
Test.copyfile()
//Test.readExcel(filepath)
//val sheet = wb.getSheetAt(1)
//val headerRow = sheet.getRow(0)
//println(sheet)
//println(headerRow)