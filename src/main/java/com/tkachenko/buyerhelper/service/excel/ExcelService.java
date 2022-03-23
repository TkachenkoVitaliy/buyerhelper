package com.tkachenko.buyerhelper.service.excel;

import com.tkachenko.buyerhelper.service.mmk.MonthSheets;
import com.tkachenko.buyerhelper.utils.AcceptMmkProperties;
import com.tkachenko.buyerhelper.utils.ExcelUtils;
import org.apache.poi.openxml4j.util.ZipSecureFile;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.*;
import org.springframework.stereotype.Service;

import java.io.FileInputStream;

import java.io.FileNotFoundException;
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

    public void rewriteSummaryFile(Path fileSummaryPath) {
        final String ACCEPT_COST_NAME = "Стоимость, руб";
        final String SHIPPED_COST_NAME = "Стоимость отгр, руб";
        final String FINAL_COST_NAME = "Итоговая стоимость, тн";

        final String PRICE_NAME = "Цена с НДС, руб/тн";
        final String ACCEPT_WEIGHT_NAME = "Акцепт, тн";
        final String SHIPPED_WEIGHT_NAME = "Отгруженно, тн";
        final String NEW_PRICE_NAME = "Пересмотр, руб/тн";

        final String PROFILE_NAME = "Размеры/профиль";

        try {
            FileInputStream summaryFileInputStream = new FileInputStream(fileSummaryPath.toString());
            ZipSecureFile.setMinInflateRatio(0);
            XSSFWorkbook summaryWorkbook = new XSSFWorkbook(summaryFileInputStream);
            int numberOfSheets = summaryWorkbook.getNumberOfSheets();
            XSSFSheet basicSheetForStyle = summaryWorkbook.getSheet(MonthSheets.JANUARY.getSheetName());
            XSSFRow basicHeaderForStyle = basicSheetForStyle.getRow(basicSheetForStyle.getFirstRowNum());
            //CellStyle headerStyle = basicHeaderForStyle.getRowStyle();
            XSSFRow basicRowForStyle = basicSheetForStyle.getRow(basicSheetForStyle.getFirstRowNum() + 1);
            //CellStyle rowStyle = basicRowForStyle.getRowStyle();

            for (int i = 0; i < numberOfSheets; i++) {
                XSSFSheet currentSheet = summaryWorkbook.getSheetAt(i);
                if (MonthSheets.findBySheetName(currentSheet.getSheetName()) != null) {
                    int headerRowIndex = currentSheet.getFirstRowNum();
                    int firstRowIndex = headerRowIndex + 1;
                    int lastRowIndex = currentSheet.getLastRowNum();

                    XSSFRow headerRow = currentSheet.getRow(headerRowIndex);
                    if (MonthSheets.findBySheetName(currentSheet.getSheetName())!=(MonthSheets.JANUARY))
                        ExcelUtils.copyRowStyleWithoutBlank(basicHeaderForStyle, headerRow);

                    int profileNameColIndex = ExcelUtils.findColumnByValue(headerRow, PROFILE_NAME);
                    int acceptCostColIndex = ExcelUtils.findColumnByValue(headerRow,ACCEPT_COST_NAME);
                    int shippedCostColIndex = ExcelUtils.findColumnByValue(headerRow, SHIPPED_COST_NAME);
                    int finalCostColIndex = ExcelUtils.findColumnByValue(headerRow, FINAL_COST_NAME);
                    int priceColIndex = ExcelUtils.findColumnByValue(headerRow, PRICE_NAME);;
                    int acceptWeightColIndex = ExcelUtils.findColumnByValue(headerRow, ACCEPT_WEIGHT_NAME);
                    int shippedWeightColIndex = ExcelUtils.findColumnByValue(headerRow, SHIPPED_WEIGHT_NAME);
                    int newPriceColIndex = ExcelUtils.findColumnByValue(headerRow, NEW_PRICE_NAME);

                    XSSFCellStyle priceCellStyle = null;

                    for(int j = firstRowIndex; j <= lastRowIndex; j++) {
                        XSSFRow currentRow = currentSheet.getRow(j);
                        XSSFCell profileCell = currentRow.getCell(profileNameColIndex);
                        if(profileCell != null && profileCell.getCellType() != CellType.BLANK) {
                            String profileValue = profileCell.getStringCellValue();
                            profileCell.setCellValue(profileValue.replaceAll("x", "*"));
                        }
                        String priceColForFormula = ExcelUtils.getExcelColAddress(priceColIndex);
                        String acceptWeightColForFormula = ExcelUtils.getExcelColAddress(acceptWeightColIndex);
                        String shippedWeightColForFormula = ExcelUtils.getExcelColAddress(shippedWeightColIndex);
                        String newPriceColForFormula = ExcelUtils.getExcelColAddress(newPriceColIndex);

                        int rowNumForFormula = j + 1;
                        XSSFCell priceCell = currentRow.getCell(priceColIndex);
                        if(priceCell.getCellStyle() != null) priceCellStyle = priceCell.getCellStyle();

                        XSSFCell acceptCostCell = currentRow.createCell(acceptCostColIndex, CellType.FORMULA);
                        XSSFCell shippedCostCell = currentRow.createCell(shippedCostColIndex, CellType.FORMULA);
                        XSSFCell finalCostCell = currentRow.createCell(finalCostColIndex, CellType.FORMULA);

                        acceptCostCell.setCellFormula(priceColForFormula + rowNumForFormula + "*"
                                + acceptWeightColForFormula + rowNumForFormula);
                        shippedCostCell.setCellFormula(priceColForFormula + rowNumForFormula + "*"
                                + shippedWeightColForFormula + rowNumForFormula);
                        finalCostCell.setCellFormula("IF(" +  newPriceColForFormula + rowNumForFormula + "=0,"
                                + priceColForFormula + rowNumForFormula + "*" + shippedWeightColForFormula + rowNumForFormula
                                + "," + newPriceColForFormula + rowNumForFormula + "*"
                                + shippedWeightColForFormula + rowNumForFormula + ")");

                        if(priceCellStyle != null) {
                            acceptCostCell.setCellStyle(priceCellStyle);
                            shippedCostCell.setCellStyle(priceCellStyle);
                            finalCostCell.setCellStyle(priceCellStyle);
                        }

                        ExcelUtils.copyRowStyleWithoutBlank(basicRowForStyle, currentRow);
                    }

                }
            }

            FileOutputStream summaryFileOutputStream = new FileOutputStream(fileSummaryPath.toString());
            summaryWorkbook.write(summaryFileOutputStream);
            summaryWorkbook.close();
            summaryFileOutputStream.flush();
            summaryFileOutputStream.close();
            summaryFileInputStream.close();
        } catch (Exception e) {
            e.printStackTrace();
        }
    }
}
