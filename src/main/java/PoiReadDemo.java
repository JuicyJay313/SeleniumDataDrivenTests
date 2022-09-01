import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.IOException;

public class PoiReadDemo {
    public static void main(String[] args) throws IOException {
        String excelFilePath = "C:\\Users\\Julien-Pier.Gagnon\\Documents\\Coding\\Selenium\\SeleniumDataDrivenTests\\src\\main\\resources\\Data\\FirstDataBatch.xlsx";
        FileInputStream inputStream = new FileInputStream(excelFilePath);

        XSSFWorkbook workbook = new XSSFWorkbook(inputStream);
        XSSFSheet sheet = workbook.getSheetAt(0);

        int rows = sheet.getLastRowNum();
        int cols = sheet.getRow(1).getLastCellNum();

        for(int r = 1; r <= rows; r++){
            XSSFRow row = sheet.getRow(r);
            for(int c = 0; c <= cols; c++){
                XSSFCell cell = row.getCell(c);
            }
        }
    }
}
