package tests;

import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.Test;

import java.io.FileInputStream;
import java.io.IOException;

public class ExcelRead {

    @Test
    public void read_from_excel_file() throws IOException {

        String path = "Sample Data.xlsx";
//To be able to read from Excel file, we need to load FileInputStream
        FileInputStream fileInputStream=new FileInputStream(path);

       //1 Create workbook
        XSSFWorkbook workbook= new XSSFWorkbook(fileInputStream);
        // 2 - We need to get specific sheet from currently opened workbook
        XSSFSheet sheet= workbook.getSheet("Employees");
        //3 Select row and sell
        System.out.println("sheet.getRow(1).getCell(0) = " + sheet.getRow(1).getCell(0));
        
        System.out.println("sheet.getRow(3).getCell(2) = " + sheet.getRow(3).getCell(2));
        
        //return the count of used cells only, will not count empty rows and cells(starts form 1)
        
        int usedRows= sheet.getPhysicalNumberOfRows();
        
        // returns the number from top cell to bottom cell(starts from 0)
        int lastUsedRow=sheet.getLastRowNum();
        
        //TODO: Create a logic to print neena's name dynamically

        for (int rowNum = 0; rowNum < usedRows; rowNum++) {

            if(sheet.getRow(rowNum).getCell(0).toString().equals("Neena")){
                System.out.println("Neena's name = "+ sheet.getRow(rowNum).getCell(0));
            }
        }
        for (int rowNum = 0; rowNum <usedRows ; rowNum++) {
            if(sheet.getRow(rowNum).getCell(0).toString().equals("Steven")){
                System.out.println("Steven's Job: "+ sheet.getRow(rowNum).getCell(2).toString());
            }

        }

    }
}
