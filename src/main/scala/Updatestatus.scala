import java.io._
import org.apache.poi.ss.util._
import org.apache.poi.xssf.usermodel._


object Updatestatus extends App {
  val wb = new XSSFWorkbook(
    getClass.getResourceAsStream("C:/work/Jira/38_DWHP_Jira_10122018.xlsx")
  )
}