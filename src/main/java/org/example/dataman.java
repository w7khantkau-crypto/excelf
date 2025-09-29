package org.example;
import java.io.File;
import java.io.FileInputStream;
import java.util.Iterator;

import org.apache.poi.ss.formula.functions.Column;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class dataman {
    private static final String test = "src/main/java/org/example/test.xls";
    public static void readAllSheets(){//read all the sheets
        try {
            FileInputStream file = new FileInputStream(new File(test));
            Workbook workbook = new XSSFWorkbook(file);
            DataFormatter formatter = new DataFormatter();
            Iterator<Sheet> sheetIterator = workbook.sheetIterator();
            while(sheetIterator.hasNext()) {
                Sheet sheet = sheetIterator.next();
                System.out.println("Sheet Name: " + sheet.getSheetName());
                System.out.println("Row Count: " + sheet.getRow(0).getLastCellNum());
                System.out.println("-----------------");
                Iterator<Row> rowIterator = sheet.rowIterator();
                while(rowIterator.hasNext()){
                   Row row = rowIterator.next();
                   Iterator<Cell> cellIterator = row.cellIterator();
                   while(cellIterator.hasNext()){
                       Cell cell = cellIterator.next();

                       String cellValue = formatter.formatCellValue(cell);
                       System.out.print(cellValue+" \t");
                   }
                   System.out.println();
                }
                System.out.println("-----------------");
            }
            workbook.close();
        }

        catch (Exception e) {
            e.printStackTrace();
        }
    }
    public static void readSaleLog(){//can be reused for different sheet read
        try {
            FileInputStream file = new FileInputStream(new File(test));
            Workbook workbook = new XSSFWorkbook(file);
            DataFormatter formatter = new DataFormatter();
            Sheet sheet = workbook.getSheetAt(0);
            System.out.println("Sheet Name: " + sheet.getSheetName());
            System.out.println("Row Count: " + sheet.getRow(0).getLastCellNum());
            System.out.println("-----------------");
            Iterator<Row> rowIterator = sheet.rowIterator();
            while(rowIterator.hasNext()){
                Row row = rowIterator.next();
                Iterator<Cell> cellIterator = row.cellIterator();

                while(cellIterator.hasNext()){
                    Cell cell = cellIterator.next();
                    String cellValue = formatter.formatCellValue(cell);
                    System.out.print(cellValue+" \t");
                }
                System.out.println();
            }
            System.out.println("-----------------");
            workbook.close();
        } catch (Exception e) {
            e.printStackTrace();
        }
    };
}
