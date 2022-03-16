package com.tkachenko.buyerhelper.service.mmk;

import com.tkachenko.buyerhelper.utils.ExcelUtils;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.nio.file.Path;

public class MmkAcceptMonthParser {

    private final String ACCEPT_MONTH_HEADER_NAME = "Месяц акцепта";
    private final String NEW_ORACLE_ORDER_NAME = "Номер СПЕЦ";
    private final Path mmkOracleFile;

    public MmkAcceptMonthParser(Path mmkOracleFile) {
        this.mmkOracleFile = mmkOracleFile;
    }

    public void parseMonth() {
        try {
            FileInputStream mmkOracleInputStream = new FileInputStream(mmkOracleFile.toString());
            XSSFWorkbook mmkOracleWorkbook = new XSSFWorkbook(mmkOracleInputStream);
            XSSFSheet mmkOracleNewSheet = mmkOracleWorkbook.getSheetAt(1);
            Row headerMmkOracleNewSheet = mmkOracleNewSheet.getRow(mmkOracleNewSheet.getFirstRowNum());
            int monthAcceptIndex = ExcelUtils.findColIndexByStringValue(ACCEPT_MONTH_HEADER_NAME, headerMmkOracleNewSheet);
            int orderIndex = ExcelUtils.findColIndexByStringValue(NEW_ORACLE_ORDER_NAME, headerMmkOracleNewSheet);
            int firstRow = mmkOracleNewSheet.getFirstRowNum()+1;
            int lastRow = mmkOracleNewSheet.getLastRowNum();

            for (int i = firstRow; i <= lastRow; i++) {
                Row targetRow = mmkOracleNewSheet.getRow(i);
                Cell targetCell = targetRow.getCell(monthAcceptIndex);
                if(targetCell == null) {
                    targetRow.createCell(monthAcceptIndex);
                }
                if(targetCell.getCellType() == CellType.BLANK) {
                    double orderNumber = targetRow.getCell(orderIndex).getNumericCellValue();
                    double acceptMonth = findSameOrderWithAcceptMonth(orderNumber, orderIndex, monthAcceptIndex, mmkOracleNewSheet,
                            firstRow, lastRow);
                    targetCell.setCellValue(acceptMonth);
                }
            }

            FileOutputStream mmkOracleOutputStream = new FileOutputStream(mmkOracleFile.toString());
            mmkOracleWorkbook.write(mmkOracleOutputStream);
            mmkOracleWorkbook.close();

            mmkOracleOutputStream.flush();
            mmkOracleOutputStream.close();

            mmkOracleInputStream.close();

        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    public double findSameOrderWithAcceptMonth(double orderNumber, int orderIndex,int monthAcceptIndex,XSSFSheet sheet,
                                               int firstRow, int lastRow) {
        for(int j = firstRow; j <= lastRow; j++) {
            Row currentRow = sheet.getRow(j);
            Cell orderCell = currentRow.getCell(orderIndex);
            if(orderCell.getNumericCellValue() == orderNumber) {
                Cell monthAcceptCell = currentRow.getCell(monthAcceptIndex);
                if(monthAcceptCell != null && monthAcceptCell.getCellType() != CellType.BLANK) {
                    return monthAcceptCell.getNumericCellValue();
                }
            }
        }

        return 0.0;
    }
}
