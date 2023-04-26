package com.demo;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.util.Arrays;
import java.util.Iterator;

public class ReadData {
    private static final String FILENAME = "C:\\Users\\jasva\\IdeaProjects\\Cappsule_Assessment\\src\\main\\resources\\Excel\\Data_Technical test.xlsx";

    public static void main(String[] args) {
        try (FileOutputStream fileOutputStream = new FileOutputStream("C:\\Users\\jasva\\IdeaProjects\\Cappsule_Assessment\\src\\main\\resources\\Excel\\data.xlsx"); Workbook workbook = new XSSFWorkbook()) {
            // Open the file
            FileInputStream fstream = new FileInputStream(FILENAME);
            Workbook wb = WorkbookFactory.create(fstream);
            //create a data formatter for formatting values
            int i = 0;
            DataFormatter dataFormatter = new DataFormatter();
            // getting the Mastersheet
            Sheet masterSheet = wb.getSheet("Master");
            // getting the Testsheet
            Sheet Test = wb.getSheet("Test");
            // creating a new sheet for storing the result
            Sheet sheet = workbook.createSheet("Sheet3");
            ;
            // the row iterator for the Testsheet
            Iterator<Row> iterator = Test.rowIterator();
            // iterating through the Testsheet


            while (iterator.hasNext()) {
                Row currentRow = iterator.next();
                // getting the cell value of the 2nd column
                Cell cell1 = currentRow.getCell(1);
                String cellValue = dataFormatter.formatCellValue(cell1);
                // converting the cell value to lowercase and removing the 's' and '-'
                cellValue = cellValue.toLowerCase().replace("s", "");
                cellValue = cellValue.replaceAll("[-'mgtable]", "");
                Iterator<Row> iterator1 = masterSheet.rowIterator();

                // iterating through the Mastersheet
                while (iterator1.hasNext()) {
                    Row currentRow1 = iterator1.next();
                    // getting the cell value of the 3rd column
                    Cell cell2 = currentRow1.getCell(2);
                    String cellValue1 = dataFormatter.formatCellValue(cell2);
                    // converting the cell value to lowercase and removing the 's' and '-'
                    cellValue1 = cellValue1.toLowerCase().replace("s", "");
                    cellValue1 = cellValue1.replaceAll("[-'mgtable]", "");

                    // Remove all spaces
                    String sortedString1 = cellValue.replaceAll("\\s+", "");
                    String sortedString2 = cellValue1.replaceAll("\\s+", "");

                    // Convert to char arrays and sort
                    char[] charArray1 = sortedString1.toCharArray();
                    char[] charArray2 = sortedString2.toCharArray();
                    Arrays.sort(charArray1);
                    Arrays.sort(charArray2);
                    // Convert back to strings and compare
                    String sorted1 = new String(charArray1);
                    String sorted2 = new String(charArray2);
                    if (sorted1.equals(sorted2)) {
                        Cell cell3 = currentRow.getCell(3);
                        Cell cell4 = currentRow1.getCell(4);
                        // getting the cell value of the 4th column
                        String cellValue3 = dataFormatter.formatCellValue(cell3);
                        String cellValue4 = dataFormatter.formatCellValue(cell4);

                        // Remove all spaces and convert to lowercase
                        String sortedString3 = cellValue3.replaceAll("\\s+", "").toLowerCase();
                        String sortedString4 = cellValue4.replaceAll("\\s+", "").toLowerCase();

                        // Convert to char arrays and sort
                        char[] charArray3 = sortedString3.toCharArray();
                        char[] charArray4 = sortedString4.toCharArray();
                        // Convert back to strings and compare
                        Arrays.sort(charArray3);
                        Arrays.sort(charArray4);
                        // Convert back to strings and compare
                        String sorted3 = new String(charArray3);
                        String sorted4 = new String(charArray4);

                        if (sorted3.equals(sorted4)) {
                            Row row = sheet.createRow(i);
                            i++;
                            System.out.println(" ");
                            // iterating through the Testsheet
                            Iterator<Cell> cellIterator12 = currentRow1.cellIterator();
                            while (cellIterator12.hasNext()) {
                                Cell currentCell1 = cellIterator12.next();
                                String cellValue6 = dataFormatter.formatCellValue(currentCell1);
                                // creating a new cell in the result sheet
                                Cell cell = row.createCell(currentCell1.getColumnIndex());
//                                if(currentCell1.getCellType()==CellType.NUMERIC){
//                                    cell.setCellValue(i);
//                                }
                                // setting the cell value
                                cell.setCellValue(cellValue6);
                                System.out.print(cellValue6 + "\t" + "  ");


                            }
                        }
                    }
                }

            }
            workbook.write(fileOutputStream);
        } catch (Exception e) {// Catch exception if any
            System.err.println("Error: " + e.getMessage());
        }
    }
}


//while (iterator.hasNext()) {
//        Row currentRow = iterator.next();
//        int rows = currentRow.getRowNum();
//        if (rows != 0) {
//
//        Iterator<Cell> cellIterator = currentRow.cellIterator();
//        while (cellIterator.hasNext()) {
//        Cell currentCell = cellIterator.next();
//        String cellValue = dataFormatter.formatCellValue(currentCell);
//
//        Iterator<Row> iterator1 = Test.rowIterator();
//        while (iterator1.hasNext()) {
//        Row currentRow1 = iterator.next();
//        int rows1 = currentRow.getRowNum();
//        if (rows1 != 0) {
//        Iterator<Cell> cellIterator1 = currentRow1.cellIterator();
//        while (cellIterator1.hasNext()) {
//        Cell currentCell1 = cellIterator1.next();
//        String cellValue1 = dataFormatter.formatCellValue(currentCell1);
//        if (cellValue.equals(cellValue1)) {
//        System.out.print(cellValue + "\t" + "  ");
//        }
//        }
//        }
//        }
//        System.out.print(cellValue + "\t" + "  ");
//
//
//        }
//        System.out.println();
//        }
//        }