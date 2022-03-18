package com.tkachenko.buyerhelper.service.mmk;

import com.tkachenko.buyerhelper.utils.ExcelUtils;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Row;

import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.nio.file.Path;

public class MmkProfileParser {
    private final String TARGET_COLUMN_NAME = "Размеры/профиль";
    private final String PRODUCT_TYPE_COLUMN_NAME = "Вид продукции";
    private final String NEW_ORACLE_ORDER_NAME = "Номер СПЕЦ";
    private final String NEW_ORACLE_POS_NAME = "Номер позиции";

    private final String TAPE = "Лента";
    private final String COIL = "Рулон";
    private final String PLATE = "Лист";
    private final String REBAR_COILS = "Профиль арматурный_моток";

    private final String FILLER = "x";

    private final Path mmkOracleFile;
    private final Path mmkAcceptLibraryFile;

    public MmkProfileParser(Path mmkOracleFile, Path mmkAcceptLibraryFile) {
        this.mmkOracleFile = mmkOracleFile;
        this.mmkAcceptLibraryFile = mmkAcceptLibraryFile;
    }

    public void parse() {
        try {
            FileInputStream mmkOracleInputStream = new FileInputStream(mmkOracleFile.toString());
            XSSFWorkbook mmkOracleWorkbook = new XSSFWorkbook(mmkOracleInputStream);
            XSSFSheet mmkOracleOldSheet = mmkOracleWorkbook.getSheetAt(0);
            Row headerMmkOracleOldSheet = mmkOracleOldSheet.getRow(mmkOracleOldSheet.getFirstRowNum());
            XSSFSheet mmkOracleNewSheet = mmkOracleWorkbook.getSheetAt(1);
            Row headerMmkOracleNewSheet = mmkOracleNewSheet.getRow(mmkOracleNewSheet.getFirstRowNum());
            int targetColumnIndex = ExcelUtils.findColumnByValue(headerMmkOracleNewSheet, TARGET_COLUMN_NAME);
            int newOracleOrderIndex = ExcelUtils.findColumnByValue(headerMmkOracleNewSheet, NEW_ORACLE_ORDER_NAME);
            int newOraclePosIndex = ExcelUtils.findColumnByValue(headerMmkOracleNewSheet, NEW_ORACLE_POS_NAME);

            FileInputStream mmkAcceptLibraryInputStream = new FileInputStream(mmkAcceptLibraryFile.toString());
            XSSFWorkbook mmkAcceptLibraryWorkbook = new XSSFWorkbook(mmkAcceptLibraryInputStream);
            XSSFSheet mmkAcceptLibrarySheet = mmkAcceptLibraryWorkbook.getSheetAt(0);


            int firstRowForParse = mmkOracleNewSheet.getFirstRowNum() + 1;
            int lastRowForParse = mmkOracleNewSheet.getLastRowNum();

            for (int i = firstRowForParse; i <= lastRowForParse; i++) {
                Row currentRow = mmkOracleNewSheet.getRow(i);
                Row currentOldRow = mmkOracleOldSheet.getRow(i);
                firstStageParse(headerMmkOracleOldSheet, currentOldRow, headerMmkOracleNewSheet, currentRow,
                        targetColumnIndex);
                secondStageParse(mmkAcceptLibrarySheet, headerMmkOracleNewSheet, currentRow, targetColumnIndex,
                        newOracleOrderIndex, newOraclePosIndex);
                NicheProfileParserFromOracle.nicheParse(headerMmkOracleOldSheet, headerMmkOracleNewSheet,
                        currentOldRow, currentRow);
                NicheProfileParserFromAccept.nicheParse(mmkAcceptLibrarySheet, headerMmkOracleNewSheet, currentRow,
                        targetColumnIndex, newOracleOrderIndex, newOraclePosIndex);
            }

            FileOutputStream mmkOracleOutputStream = new FileOutputStream(mmkOracleFile.toString());
            mmkOracleWorkbook.write(mmkOracleOutputStream);
            mmkOracleWorkbook.close();

            mmkOracleOutputStream.flush();
            mmkOracleOutputStream.close();

            mmkOracleInputStream.close();
            mmkAcceptLibraryInputStream.close();


        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    public void firstStageParse(Row oldHeader, Row oldRow, Row newHeader, Row newRow, int targetColIndex) {
        final String FIRST_MEASURE = "Толщ";
        final String SECOND_MEASURE = "Ширина";
        final String THIRD_MEASURE = "Длина";
        final String PROFILE = "Профиль";

        int productTypeColIndex = ExcelUtils.findColumnByValue(newHeader, PRODUCT_TYPE_COLUMN_NAME);
        int firstMeasureIndex = ExcelUtils.findColumnByValue(oldHeader, FIRST_MEASURE);
        int secondMeasureIndex = ExcelUtils.findColumnByValue(oldHeader, SECOND_MEASURE);
        int thirdMeasureIndex = ExcelUtils.findColumnByValue(oldHeader, THIRD_MEASURE);
        int profileIndex = ExcelUtils.findColumnByValue(oldHeader, PROFILE);

        String productTypeValue = newRow.getCell(productTypeColIndex).getStringCellValue();


        if (newRow.getCell(targetColIndex) == null) newRow.createCell(targetColIndex);
        Cell targetCell = newRow.getCell(targetColIndex);

        parseProfileForType(productTypeValue, oldRow, targetCell, firstMeasureIndex, secondMeasureIndex,
                thirdMeasureIndex, profileIndex);



    }

    public void secondStageParse(XSSFSheet mmkAcceptLibrarySheet, Row targetHeader, Row targetRow, int targetColIndex,
                                 int newOracleOrderIndex, int newOraclePosIndex ) {
        final String FIRST_MEASURE = "Толщина от";
        final String SECOND_MEASURE = "Ширина от";
        final String THIRD_MEASURE = "Длина от";
        final String PROFILE = "Альт. Профиль";

        final String ACCEPT_LIBRARY_ORDER_NAME = "Номер заказа";
        final String ACCEPT_LIBRARY_POS_NAME = "Номер строки";

        int productTypeColIndex = ExcelUtils.findColumnByValue(targetHeader, PRODUCT_TYPE_COLUMN_NAME);
        String productTypeValue = targetRow.getCell(productTypeColIndex).getStringCellValue();

        Row headerMmkAcceptLibrary = mmkAcceptLibrarySheet.getRow(mmkAcceptLibrarySheet.getFirstRowNum());
        int acceptLibraryOrderIndex = ExcelUtils.findColumnByValue(headerMmkAcceptLibrary,ACCEPT_LIBRARY_ORDER_NAME);
        int acceptLibraryPosIndex = ExcelUtils.findColumnByValue(headerMmkAcceptLibrary,ACCEPT_LIBRARY_POS_NAME);
        int firstRowAcceptLibrary = mmkAcceptLibrarySheet.getFirstRowNum()+1;
        int lastRowAcceptLibrary = mmkAcceptLibrarySheet.getLastRowNum();

        int firstMeasureIndex = ExcelUtils.findColumnByValue(headerMmkAcceptLibrary, FIRST_MEASURE);
        int secondMeasureIndex = ExcelUtils.findColumnByValue(headerMmkAcceptLibrary, SECOND_MEASURE);
        int thirdMeasureIndex = ExcelUtils.findColumnByValue(headerMmkAcceptLibrary, THIRD_MEASURE);
        int profileIndex = ExcelUtils.findColumnByValue(headerMmkAcceptLibrary, PROFILE);

        Cell targetCell = targetRow.getCell(targetColIndex);
        if(targetCell == null) targetRow.createCell(targetColIndex);

        if(targetCell.getCellType() == CellType.BLANK) {
            for(int i = firstRowAcceptLibrary; i <= lastRowAcceptLibrary; i++) {
                Row rowFrom = mmkAcceptLibrarySheet.getRow(i);
                if (ExcelUtils.isSamePosition(targetRow, newOracleOrderIndex, newOraclePosIndex,
                        rowFrom, acceptLibraryOrderIndex, acceptLibraryPosIndex)) {
                    parseProfileForType(productTypeValue, rowFrom, targetCell, firstMeasureIndex,
                                    secondMeasureIndex, thirdMeasureIndex, profileIndex);
                }
            }
        }
    }


// innerMethod
    private void parseProfileForType(String productTypeValue, Row rowFrom, Cell targetCell,
                                    int firstMeasureIndex, int secondMeasureIndex, int thirdMeasureIndex,
                                    int profileIndex) {

        DataFormatter getStringValueFormatter = new DataFormatter();
        if (productTypeValue.contains(COIL) | productTypeValue.contains(TAPE)) {

            if (rowFrom.getCell(firstMeasureIndex) != null && rowFrom.getCell(secondMeasureIndex) != null
            && rowFrom.getCell(firstMeasureIndex).getCellType() !=CellType.BLANK
                    && rowFrom.getCell(secondMeasureIndex).getCellType() != CellType.BLANK) {
                String firstMeasureValue = getStringValueFormatter.formatCellValue(rowFrom.getCell(firstMeasureIndex));
                String secondMeasureValue = getStringValueFormatter.formatCellValue(rowFrom.getCell(secondMeasureIndex));
                targetCell.setCellValue(firstMeasureValue + FILLER + secondMeasureValue);
            }

        } else if (productTypeValue.contains(PLATE)) {

            if (rowFrom.getCell(firstMeasureIndex) != null && rowFrom.getCell(secondMeasureIndex) != null
                    && rowFrom.getCell(thirdMeasureIndex) != null
                    && rowFrom.getCell(firstMeasureIndex).getCellType() != CellType.BLANK
                    && rowFrom.getCell(secondMeasureIndex).getCellType() != CellType.BLANK
                    && rowFrom.getCell(thirdMeasureIndex).getCellType() != CellType.BLANK) {

                String firstMeasureValue = getStringValueFormatter.formatCellValue(rowFrom.getCell(firstMeasureIndex));
                String secondMeasureValue = getStringValueFormatter.formatCellValue(rowFrom.getCell(secondMeasureIndex));
                String thirdMeasureValue = getStringValueFormatter.formatCellValue(rowFrom.getCell(thirdMeasureIndex));
                targetCell.setCellValue(firstMeasureValue + FILLER + secondMeasureValue + FILLER + thirdMeasureValue);
            }

        } else if(productTypeValue.contains(REBAR_COILS)){

            if (rowFrom.getCell(profileIndex) != null && rowFrom.getCell(profileIndex).getCellType() != CellType.BLANK) {

                String profileValue = getStringValueFormatter.formatCellValue(rowFrom.getCell(profileIndex));
                targetCell.setCellValue(profileValue + " бунт");

            }

        } else {

            if (rowFrom.getCell(profileIndex) != null) {

                String profileValue = getStringValueFormatter.formatCellValue(rowFrom.getCell(profileIndex));
                targetCell.setCellValue(profileValue);

            }

        }
    }

}
