package com.example.testapachepoi;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;

public class ReadExcelFileToList {
    public static List<TermAndDefinition> readExcelData(String fileName){
        List<TermAndDefinition> termAndDefinitionList = new ArrayList<>();

        try {
            //Create the input stream from the xlsx/xls file
            FileInputStream excelFileInputStream = new FileInputStream(fileName);

            //Create Workbook instance for xlsx/xls file input stream
            Workbook workbook = null;
            if (fileName.toLowerCase().endsWith("xlsx") ||
                    fileName.toLowerCase().endsWith("xls")
            ) {
                workbook = new XSSFWorkbook(excelFileInputStream);
            }

            //Get the number of sheets in the xlsx file
            int numberOfSheets = workbook.getNumberOfSheets();

            for (int i = 0; i < numberOfSheets; i++) {
                Sheet sheet = workbook.getSheetAt(i);

                Iterator<Row> rowIterator = sheet.iterator();

                rowIterator.next();
                rowIterator.next();
                rowIterator.next();

                while (rowIterator.hasNext()) {
                    String term = "";
                    String definition = "";

                    Row row = rowIterator.next();

                    Iterator<Cell> cellIterator = row.cellIterator();
                    while (cellIterator.hasNext()) {
                        Cell cell = cellIterator.next();

                        if (cell.getCellType() == 1){
                            switch (cell.getCellType()) {
                                case Cell.CELL_TYPE_STRING:
                                    if (term.equalsIgnoreCase("") )
                                        term = cell.getStringCellValue().trim();
                                    else if (definition.equalsIgnoreCase("") )
                                        definition = cell.getStringCellValue().trim();
                                    break;
                            }
                            TermAndDefinition newTermAndDefinition = new TermAndDefinition(term, definition);
                            termAndDefinitionList.add(newTermAndDefinition);
                        }
                    }
                }
            }
        } catch(IOException e){
            e.printStackTrace();
        }
        return termAndDefinitionList;
    }

    private static String FILE_NAME = "AgileTerm_TermAndDefinition_Template";
    private static String EXTENSION_PART = ".xlsx";

    public static void main (String args[]){
        List<TermAndDefinition> list = readExcelData(FILE_NAME + EXTENSION_PART );
        for (TermAndDefinition termAndDef: list) {
            System.out.println(termAndDef);
        }
    }
}
