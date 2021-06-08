/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package excel_read_write;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.Iterator;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 *
 * @author root_user
 */
public class Excel_read_write {

    /**
     * @param args the command line arguments
     */
    public static void main(String[] args) throws IOException {

        //this is for .xlsx file reading 
//        FileInputStream fileInputStream = new FileInputStream("E:\\company_list\\cloud_revel\\product_section\\excel_read_write\\excel_file\\Book1.xlsx");
//        HSSFWorkbook workBook = new HSSFWorkbook(fileInputStream);
//        HSSFSheet sheet = workBook.getSheetAt(0);
//        FormulaEvaluator formulaEvaluator = workBook.getCreationHelper().createFormulaEvaluator();
//        for (Row row : sheet) {
//            for (Cell cell : row) {
//                switch (formulaEvaluator.evaluateInCell(cell).getCellType()) {
//                    case Cell.CELL_TYPE_NUMERIC:
//                        System.out.println(cell.getNumericCellValue() + "\t\t");
//                        break;
//                    case Cell.CELL_TYPE_STRING:
//                        System.out.println(cell.getStringCellValue() + "\t\t");
//                        break;
//                }
//            }
//            System.out.println("");
//        }
        File file = new File("E:\\company_list\\cloud_revel\\product_section\\excel_read_write\\excel_file\\Book1.xlsx");
        FileInputStream fileInputStream = new FileInputStream(file);
        XSSFWorkbook fWorkbook = new XSSFWorkbook(fileInputStream);
        XSSFSheet sheet = fWorkbook.getSheetAt(0);
        Iterator<Row> iterator = sheet.iterator();
        while (iterator.hasNext()) {
            Row row = iterator.next();
            Iterator<Cell> cellIterator = row.cellIterator();
            while (cellIterator.hasNext()) {
                Cell cell = cellIterator.next();
                switch (cell.getCellType()) {
                    case Cell.CELL_TYPE_STRING:
                        System.out.println(cell.getStringCellValue() + "\t\t\t");
                    case Cell.CELL_TYPE_NUMERIC:
                        System.out.println(cell.getNumericCellValue() + "\t\t");
                }
            }
        }

    }

}
