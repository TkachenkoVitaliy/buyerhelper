package com.tkachenko.buyerhelper.service.excel;

import com.tkachenko.buyerhelper.utils.AcceptMmkProperties;
import com.tkachenko.buyerhelper.utils.ExcelUtils;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.stereotype.Service;

import java.io.FileInputStream;

import java.io.FileOutputStream;
import java.nio.file.Path;
import java.util.Iterator;

@Service
public class ExcelService {
    private final static String firstSheetForDeleteName = "НАСТРОЙКИ";
    private final static String secondSheetForDeleteName = "Updates History";
    private final static String acceptLibraryName = "AcceptLibrary.xlsx";

    public String getAcceptLibraryName() {
        return acceptLibraryName;
    }

    public void refactorSummaryFile(Path fileSummaryPath) {
        String stringFileSummaryPath = fileSummaryPath.toString();

        try {
            FileInputStream fileInputStream = new FileInputStream(stringFileSummaryPath);

            XSSFWorkbook workbook = new XSSFWorkbook(fileInputStream);

            ExcelUtils.deleteSheetIfExists(workbook, firstSheetForDeleteName);
            ExcelUtils.deleteSheetIfExists(workbook, secondSheetForDeleteName);

            FileOutputStream fileOutputStream = new FileOutputStream(stringFileSummaryPath);
            workbook.write(fileOutputStream);
            workbook.close();

            fileOutputStream.flush();
            fileOutputStream.close();
            fileInputStream.close();

        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    public void parseMmkAccept (Path fileMmkAcceptPath, Path fileMMkAcceptRefactoredPath) {
        String stringFileMMkAcceptPath = fileMmkAcceptPath.toString();
        String stringFileMMkAcceptRefactoredPath = fileMMkAcceptRefactoredPath.toString();

        try {
            FileInputStream fileInputStream = new FileInputStream(stringFileMMkAcceptPath);
            XSSFWorkbook workbookInput = new XSSFWorkbook(fileInputStream);
            XSSFSheet sheetInput = workbookInput.getSheetAt(0);
            sheetInput.removeRow(sheetInput.getRow(sheetInput.getFirstRowNum()));
            int firstRowNum = sheetInput.getFirstRowNum(); //TODO method to find first not null row
            Row headerRow = sheetInput.getRow(firstRowNum);
            String[] columnsNames = AcceptMmkProperties.getArrayColumns();
            int[] columnsIndexes = new int[columnsNames.length];

            for (int i = 0; i < columnsNames.length; i++) {
                String columnName = columnsNames[i];

                for (Cell cell: headerRow) {
                    String cellValue = cell.getStringCellValue();
                    if (cellValue.equals(columnName)) {
                        int columnIndex = cell.getColumnIndex();
                        columnsIndexes[i] = columnIndex;
                        break;
                    }
                }
            }

            XSSFWorkbook workbookOut = new XSSFWorkbook();
            XSSFSheet sheetOut = workbookOut.createSheet("Data");


            Iterator<Row> rowIterator = sheetInput.iterator();
            while (rowIterator.hasNext()) {
                Row rowInput = rowIterator.next();
                int rowNum = rowInput.getRowNum();
                Row rowOutput = sheetOut.createRow(rowNum);
                if (!rowIterator.hasNext()) break; //TODO refactor skip summary row in the of file
                for (int i = 0; i < columnsIndexes.length; i++) {
                    Cell cellInput = rowInput.getCell(columnsIndexes[i]);
                    if (cellInput != null) {
                        ExcelUtils.copyCellValue(cellInput, rowOutput, i);
                    }

                }
            }

            FileOutputStream fileOutputStream = new FileOutputStream(stringFileMMkAcceptRefactoredPath);
            workbookOut.write(fileOutputStream);
            workbookOut.close();
            workbookInput.close();

            fileOutputStream.flush();
            fileOutputStream.close();

        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    public void addToAcceptLibrary (Path fileMMkAcceptRefactoredPath) {
        String stringFileMmkAcceptRefactoredPath = fileMMkAcceptRefactoredPath.toString();
        String stringAcceptLibraryPath = fileMMkAcceptRefactoredPath.getParent().resolve(acceptLibraryName).toString();
        try {
            FileInputStream fileInputStreamMAR = new FileInputStream(stringFileMmkAcceptRefactoredPath);
            XSSFWorkbook workbookInputMAR = new XSSFWorkbook(fileInputStreamMAR);
            XSSFSheet sheetMAR = workbookInputMAR.getSheetAt(0);
            int headerRowNum = sheetMAR.getFirstRowNum();
            int firstRowNum = headerRowNum + 1;
            int lastRowNum = sheetMAR.getLastRowNum();// -1 TODO rewrite for check if last row is blank


            FileInputStream fileInputStreamAL = new FileInputStream(stringAcceptLibraryPath);
            XSSFWorkbook workbookInputAL = new XSSFWorkbook(fileInputStreamAL);
            XSSFSheet sheetAL = workbookInputAL.getSheetAt(0);
            int headerRowNumAL = sheetAL.getFirstRowNum();
            int firstRowNumAL = headerRowNumAL + 1;
            int lastRowNumAL;
            lastRowNumAL = sheetAL.getLastRowNum();


            for (int i = firstRowNum; i <= lastRowNum; i++) {
                Row rowMAR = sheetMAR.getRow(i);
                if (ExcelUtils.isRowEmpty(rowMAR)) continue;
                boolean findPosition = false;
                int samePositionRowNum = -1;

                for (Row rowAL : sheetAL) {
                    if (rowAL.getRowNum() != headerRowNumAL) {
                        if (ExcelUtils.isSamePosition(rowMAR, rowAL)) {
                            findPosition = true;
                            samePositionRowNum = rowAL.getRowNum();
                        }
                    }

                }

                if(findPosition) {
                    Row rewriteRowAL = sheetAL.getRow(samePositionRowNum);
                    for (Cell cell: rowMAR) {
                        int columnIndex = cell.getColumnIndex();
                        ExcelUtils.copyCellValue(cell, rewriteRowAL, columnIndex);
                    }
                } else {
                    lastRowNumAL++;
                    Row rowNewAL = sheetAL.createRow(lastRowNumAL);
                    for (Cell cell: rowMAR) {
                        int columnIndex = cell.getColumnIndex();
                        ExcelUtils.copyCellValue(cell, rowNewAL, columnIndex);
                    }
                }
            }

            FileOutputStream fileOutputStream = new FileOutputStream(stringAcceptLibraryPath);
            workbookInputAL.write(fileOutputStream);
            workbookInputAL.close();

            fileOutputStream.flush();
            fileOutputStream.close();

            fileInputStreamAL.close();
            fileInputStreamMAR.close();
        }

         catch (Exception e) {
            e.printStackTrace();
        }
    }
}
