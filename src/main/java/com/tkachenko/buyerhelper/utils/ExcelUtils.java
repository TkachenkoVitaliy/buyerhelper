package com.tkachenko.buyerhelper.utils;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellCopyPolicy;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelUtils {

    public static void deleteSheetIfExists (XSSFWorkbook workbook, String sheetName) {
        int sheetIndex = workbook.getSheetIndex(sheetName);
        if (sheetIndex >= 0) {
            workbook.removeSheetAt(sheetIndex);
        }
    }

    public static boolean isSamePosition (Row rowFirst, Row rowSecond) {
        final int firstCol = 0;
        final int secondCol = 1;

        Cell cellFirst;
        Cell cellSecond;

        cellFirst = rowFirst.getCell(firstCol);
        cellSecond = rowSecond.getCell(firstCol);
        if (cellFirst == null | cellSecond == null) {
            return false;
        }

        if (cellFirst.getNumericCellValue() != cellSecond.getNumericCellValue() ) {
            return false;
        }
        cellFirst = rowFirst.getCell(secondCol);
        cellSecond = rowSecond.getCell(secondCol);
        if (cellFirst.getNumericCellValue() != cellSecond.getNumericCellValue() ) {
            return false;
        } else {
            return true;
        }
    }

    public static boolean isRowEmpty (Row row) {
        for (int c = row.getFirstCellNum(); c < row.getLastCellNum(); c++) {
            Cell cell = row.getCell(c);
            if (cell != null && cell.getCellType() != CellType.BLANK) return false;
        }
        return true;
    }

    public static void copyCellValueXSSF (XSSFCell from, XSSFCell to) {

        switch (from.getCellType()) {
            case STRING:
            case NUMERIC:
            case FORMULA:
            case BOOLEAN:
            case ERROR:
                to.copyCellFrom(from, new CellCopyPolicy());
                break;
            case BLANK:
                to.setBlank();
                break;
            case _NONE:
                break;
        }
    }

    public static void copyCellValue (Cell from, Cell to) {

        switch (from.getCellType()) {
            case STRING:
                to.setCellValue(from.getStringCellValue());
                break;
            case NUMERIC:
                to.setCellValue(from.getNumericCellValue());
                break;
            case FORMULA:
                to.setCellFormula(from.getCellFormula());
                break;
            case BOOLEAN:
                to.setCellValue(from.getBooleanCellValue());
                break;
            case ERROR:
                to.setCellErrorValue(from.getErrorCellValue());
                break;
            case BLANK:
                to.setBlank();
                break;
            case _NONE:
                break;
        }
    }

    public static void copyCellValue (Cell from, Row rowTo, int indexTo) {
        Cell to = rowTo.createCell(indexTo);
        switch (from.getCellType()) {
            case STRING:
                to.setCellValue(from.getStringCellValue());
                break;
            case NUMERIC:
                to.setCellValue(from.getNumericCellValue());
                break;
            case FORMULA:
                to.setCellFormula(from.getCellFormula());
                break;
            case BOOLEAN:
                to.setCellValue(from.getBooleanCellValue());
                break;
            case ERROR:
                to.setCellErrorValue(from.getErrorCellValue());
                break;
            case BLANK:
                to.setBlank();
                break;
            case _NONE:
                break;
        }
    }

    public static int findColIndexByStringValue(String value, Row row) {
        int result = -1;
        for (Cell cell : row) {
            if (cell.getCellType() == CellType.STRING) {
                if(cell.getStringCellValue().equals(value)) {
                    result = cell.getColumnIndex();
                    return result;
                }
            }
        }
        return result;
    }

    public static int findColumnByValue(Row searchableRow, String searchableValue) {
        for (Cell cell : searchableRow) {
            if (cell != null && cell.getCellType() != CellType.BLANK) {
                if(cell.getStringCellValue().equals(searchableValue)) {
                    return cell.getColumnIndex();
                }
            }
        }
        return -1;
    }
}
