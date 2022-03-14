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

    private final String TAPE = "Лента";
    private final String COIL = "Рулон";
    private final String PLATE = "Лист";


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
            int targetRowIndex = ExcelUtils.findColumnByValue(headerMmkOracleNewSheet, TARGET_COLUMN_NAME);

            FileInputStream mmkAcceptLibraryInputStream = new FileInputStream(mmkAcceptLibraryFile.toString());
            XSSFWorkbook mmkAcceptLibraryWorkbook = new XSSFWorkbook(mmkAcceptLibraryInputStream);
            XSSFSheet mmkAcceptLibrarySheet = mmkAcceptLibraryWorkbook.getSheetAt(0);


            int firstRowForParse = mmkOracleNewSheet.getFirstRowNum() + 1;
            int lastRowForParse = mmkOracleNewSheet.getLastRowNum();

            for (int i = firstRowForParse; i <= lastRowForParse; i++) {
                Row currentRow = mmkOracleNewSheet.getRow(i);
                Row currentOldRow = mmkOracleOldSheet.getRow(i);
                firstStageParse(headerMmkOracleOldSheet, currentOldRow, headerMmkOracleNewSheet, currentRow, targetRowIndex);
                secondStageParse();
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

    public void firstStageParse(Row oldHeader, Row oldRow, Row header, Row row, int targetColIndex) {
        final String FIRST_MEASURE = "Толщ";
        final String SECOND_MEASURE = "Ширина";
        final String THIRD_MEASURE = "Длина";
        final String PROFILE = "Профиль";
        final String FILLER = "x";

        int productTypeColIndex = ExcelUtils.findColumnByValue(header, PRODUCT_TYPE_COLUMN_NAME);
        int firstMeasureIndex = ExcelUtils.findColumnByValue(oldHeader, FIRST_MEASURE);
        int secondMeasureIndex = ExcelUtils.findColumnByValue(oldHeader, SECOND_MEASURE);
        int thirdMeasureIndex = ExcelUtils.findColumnByValue(oldHeader, THIRD_MEASURE);
        int profileIndex = ExcelUtils.findColumnByValue(oldHeader, PROFILE);

        String productTypeValue = row.getCell(productTypeColIndex).getStringCellValue();
        DataFormatter getStringValueFormatter = new DataFormatter();


        if (row.getCell(targetColIndex) == null) row.createCell(targetColIndex);
        Cell targetCell = row.getCell(targetColIndex);

        if (productTypeValue.contains(COIL) | productTypeValue.contains(TAPE)) {

            if (oldRow.getCell(firstMeasureIndex) != null && oldRow.getCell(secondMeasureIndex) != null) {
                String firstMeasureValue = getStringValueFormatter.formatCellValue(oldRow.getCell(firstMeasureIndex));
                String secondMeasureValue = getStringValueFormatter.formatCellValue(oldRow.getCell(secondMeasureIndex));
                targetCell.setCellValue(firstMeasureValue + FILLER + secondMeasureValue);
            }

        } else if (productTypeValue.contains(PLATE)) {

            if (oldRow.getCell(firstMeasureIndex) != null && oldRow.getCell(secondMeasureIndex) != null
                    && oldRow.getCell(thirdMeasureIndex) != null) {

                String firstMeasureValue = getStringValueFormatter.formatCellValue(oldRow.getCell(firstMeasureIndex));
                String secondMeasureValue = getStringValueFormatter.formatCellValue(oldRow.getCell(secondMeasureIndex));
                String thirdMeasureValue = getStringValueFormatter.formatCellValue(oldRow.getCell(thirdMeasureIndex));
                targetCell.setCellValue(firstMeasureValue + FILLER + secondMeasureValue + FILLER + thirdMeasureValue);
            }

        } else {

            if (oldRow.getCell(profileIndex) != null) {

                String profileValue = getStringValueFormatter.formatCellValue(oldRow.getCell(profileIndex));
                targetCell.setCellValue(profileValue);

            }

        }

    }

    public void secondStageParse() {

    }

}
