import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.Iterator;

public class ParseExcel {

    private String regexPattern= "";
    private String fileName = "";


    public static void main(String[] args) throws FileNotFoundException {
        String value;
        FileInputStream fis = new FileInputStream(new File("C:\\Asha_Sasi\\test.xlsx"));
        try {
            Workbook wb = new XSSFWorkbook(fis);
            Sheet sheet = wb.getSheetAt(0);
            Iterator<Row> itr = sheet.iterator();

            while (itr.hasNext()) {
                Row row = itr.next();
                Cell cell = row.getCell(2);
                value = cell.getStringCellValue();
                if(value.matches("^[a-zA-Z]*$")) {
                    System.out.println(value);
                }

            }
        } catch (IOException e) {
            e.printStackTrace();
        }
    }


}

