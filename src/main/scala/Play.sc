import java.io._
import java.io.File
import java.io.FileInputStream
import java.io.FileNotFoundException
import java.io.IOException

import org.apache.poi.ss.usermodel._
import org.apache.poi.xssf.usermodel.XSSFWorkbook


object Test {

  private val filepath = "C:\\work\\Jira\\UpdateJiraTickets\\data\\dwhp.xlsx"
  private val filepath2 = "C:\\work\\Jira\\UpdateJiraTickets\\data\\40_DWHP_Jira_08012019.xlsx"

  private val columnMap: Map[String, Int] = Map("Issue Type" -> 0,
    "Key" -> 1,
    "Summary" -> 2,
    "Status" -> 3,
    "Assignee" -> 4,
    "Reporter" -> 5,
    "Created" -> 6,
    "Updated" -> 7,
    "Due Date" -> 8,
    "Description" -> 9,
    "All Comments" -> 10,
    "Prioritat" -> 11,
    "Milestones" -> 12,
    "Last Comment" -> 13)

  private val columnToCopyMap: Map[String, Int] = Map("Prioritat" -> 0,
    "Milestones" -> 1,
    "Last Comment" -> 2,
    "Assignee" -> 4)

  private var PositionMap = scala.collection.mutable.Map[Int, Int]()

  def copyfile(): Unit = {
    val src = "C:\\work\\scala_work\\39_DWHP_Jira_17122018.xlsx"
    val dest = "C:\\work\\scala_work\\40_DWHP_Jira_17122018.xlsx"
    var inputChannel = new FileInputStream(src).getChannel()
    var outputChannel = new FileOutputStream(dest).getChannel()
    outputChannel.transferFrom(inputChannel, 0, inputChannel.size())
    inputChannel.close()
    outputChannel.close()
  }


  def readExcel(filepath: String): Unit = {
    try {

      val excelFile = new FileInputStream(new File(filepath))
      val wb = new XSSFWorkbook(excelFile)
      val worksheet = wb.getSheetAt(0)
      val iterator = worksheet.iterator()


      while (iterator.hasNext) {
        val currentrow = iterator.next()
        println(currentrow.getCell(1))

      }

    }
    catch {
      case ex: FileNotFoundException => ex.printStackTrace()
      case ex: IOException => ex.printStackTrace()
    }
  }

  def main(args: Array[String]): Unit = {

    println("Start")
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
      val cell = headerRow.createCell(value)
      cell.setCellValue(key)
      cell.setCellStyle(headerCellStyle)
    }

    val dateCellStyle = workbook.createCellStyle()
    dateCellStyle.setDataFormat(createHelper.createDataFormat().getFormat("dd-MM-yyyy"))

    println("before try")
    try {

      /**********copy from the new report*********************/
      val excelFile = new FileInputStream(new File(filepath))
      val wb = new XSSFWorkbook(excelFile)
      val worksheet = wb.getSheetAt(0)
      val sheetIterator = worksheet.iterator()

      println("before while")
      var rowNum = 0
      while (sheetIterator.hasNext) {

        val Row = sheetIterator.next()
        val RowIterator = Row.iterator()

        if (rowNum == 0) {
          var cell0 = 0
          while (RowIterator.hasNext) {
            val oldCell = RowIterator.next().toString

            if (columnMap.contains(oldCell)) {
              PositionMap += (cell0 -> columnMap(oldCell))
            }
            cell0 = cell0 + 1
          }
          /*
           for ((k, v) <- columnMap2) {
             println(f"key - $k%s, value - $v%s")
           }
           */
        }
        else {
          val row = sheet.createRow(rowNum)

          var cell = 0
          while (RowIterator.hasNext) {
            val oldCell = RowIterator.next().toString

            if (PositionMap.contains(cell)) {
              row.createCell(PositionMap(cell)).setCellValue(oldCell)
            }
            cell = cell + 1
          }
        }
        rowNum = rowNum + 1
      }


      /**********copy from the old report*********************/
      val excelFileOld = new FileInputStream(new File(filepath2))
      val wbOld = new XSSFWorkbook(excelFile)
      val worksheetOld = wbOld.getSheetAt(0)
      val sheetIteratorOld = worksheetOld.iterator()



      println("before while2")
      var rowNum2 = 0
      while (sheetIteratorOld.hasNext) {

        val oldRow = sheetIteratorOld.next()
        val oldRowIterator = oldRow.iterator()

        oldRow.getCell(1)


        if (rowNum2 == 0) {
          var cell0 = 0
          while (oldRowIterator.hasNext) {
            val oldCell = oldRowIterator.next().toString

            if (columnMap.contains(oldCell)) {
              PositionMap += (cell0 -> columnMap(oldCell))
            }
            cell0 = cell0 + 1
          }
          /*
           for ((k, v) <- columnMap2) {
             println(f"key - $k%s, value - $v%s")
           }
           */
        }
        else {
          val row = sheet.createRow(rowNum)

          var cell = 0
          while (oldRowIterator.hasNext) {
            val oldCell = oldRowIterator.next().toString

            if (PositionMap.contains(cell)) {
              row.createCell(PositionMap(cell)).setCellValue(oldCell)
            }
            cell = cell + 1
          }
        }
        rowNum = rowNum + 1
      }

      for ((key, value) <- columnMap) {
        sheet.autoSizeColumn(value)
      }

      val fileOut = new FileOutputStream("C:\\work\\Jira\\UpdateJiraTickets\\data\\copyDWHP.xlsx")
      workbook.write(fileOut)

      // Closing the all
      fileOut.close()
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

//val filepath = "C:\\work\\Jira\\UpdateJiraTickets\\data\\dwhp.xlsx"
//Test.readExcel(filepath)
Test.main(Array("bla bla"))
//Test.copyfile()
//Test.readExcel(filepath)
//val sheet = wb.getSheetAt(1)
//val headerRow = sheet.getRow(0)
//println(sheet)
//println(headerRow)
