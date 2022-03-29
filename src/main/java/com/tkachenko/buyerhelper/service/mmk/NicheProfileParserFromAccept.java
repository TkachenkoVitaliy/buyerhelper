package com.tkachenko.buyerhelper.service.mmk;

import com.tkachenko.buyerhelper.utils.ExcelUtils;
import com.tkachenko.buyerhelper.utils.RegexUtils;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;


public class NicheProfileParserFromAccept {
    private final static String ANGLES = "Уголок г/к";
    private final static String REBARS_COILS = "Профиль арматурный_моток";
    private final static String COLD_ROLLED_SPECIAL_SECTIONS = "Спецпрофиль х/г";
    private final static String U_CHANNELS = "Швеллер г/к";
    private final static String FILLER = "x";

    private final static String NEW_PRODUCT_TYPE_COLUMN_NAME = "Вид продукции";

    private final static String ACCEPT_LIBRARY_ORDER_NAME = "Номер заказа";
    private final static String ACCEPT_LIBRARY_POS_NAME = "Номер строки";
    private final static String ADDITIONAL_REQ_NAME = "Доп.тех.требования";
    private final static String ACCEPT_LIBRARY_PROFILE_NAME = "Наименование позиции";
    private final static String LENGTH_HEADER_ACCEPT_NAME = "Длина от";


    public static void nicheParse (XSSFSheet mmkAcceptLibrarySheet, Row targetHeader, Row targetRow, int targetColIndex,
                                   int newOracleOrderIndex, int newOraclePosIndex) {
        Row headerMmkAcceptLibrary = mmkAcceptLibrarySheet.getRow(mmkAcceptLibrarySheet.getFirstRowNum());
        int firstRowAcceptLibrary = mmkAcceptLibrarySheet.getFirstRowNum()+1;
        int lastRowAcceptLibrary = mmkAcceptLibrarySheet.getLastRowNum();
        int acceptLibraryOrderIndex = ExcelUtils.findColumnByValue(headerMmkAcceptLibrary,ACCEPT_LIBRARY_ORDER_NAME);
        int acceptLibraryPosIndex = ExcelUtils.findColumnByValue(headerMmkAcceptLibrary,ACCEPT_LIBRARY_POS_NAME);
        int acceptLibraryAdditionalReqIndex = ExcelUtils.findColumnByValue(headerMmkAcceptLibrary, ADDITIONAL_REQ_NAME);
        int acceptLibraryProfileNameIndex = ExcelUtils.findColumnByValue(headerMmkAcceptLibrary, ACCEPT_LIBRARY_PROFILE_NAME);
        int lengthHeaderAcceptIndex = ExcelUtils.findColumnByValue(headerMmkAcceptLibrary, LENGTH_HEADER_ACCEPT_NAME);

        int newProductTypeIndex = ExcelUtils.findColIndexByStringValue(NEW_PRODUCT_TYPE_COLUMN_NAME, targetHeader);
        String productTypeValue = targetRow.getCell(newProductTypeIndex).getStringCellValue();

        Cell targetCell = targetRow.getCell(targetColIndex);
        if(targetCell == null) targetRow.createCell(targetColIndex);
        DataFormatter formatter = new DataFormatter();
        String targetCellStringValue = formatter.formatCellValue(targetCell);

        if(targetCell.getCellType() == CellType.BLANK || targetCellStringValue.equals("")) {

            for(int i = firstRowAcceptLibrary; i <= lastRowAcceptLibrary; i++) {
                Row rowFrom = mmkAcceptLibrarySheet.getRow(i);
                if (ExcelUtils.isSamePosition(targetRow, newOracleOrderIndex, newOraclePosIndex,
                        rowFrom, acceptLibraryOrderIndex, acceptLibraryPosIndex)) {

                    if(productTypeValue.equals(ANGLES)) {
                        String additionalReq = rowFrom.getCell(acceptLibraryAdditionalReqIndex).
                                getStringCellValue();

                        DataFormatter dataFormatter = new DataFormatter();
                        String lengthValue = dataFormatter.formatCellValue(rowFrom.getCell(lengthHeaderAcceptIndex));

                        final String FIRST_MEASURE_REGEX = "(Ширина полки=[0-9]{1,3})";
                        final String REMOVE_FIRST_MEASURE_REGEX = "(Ширина полки=)";
                        final String SECOND_MEASURE_REGEX = "(Толщина полки профиля=[0-9]{1,2})";
                        final String REMOVE_SECOND_MEASURE_REGEX = "(Толщина полки профиля=)";

                        String firstMeasureValue = RegexUtils.regexWithRemove(additionalReq,
                                FIRST_MEASURE_REGEX, REMOVE_FIRST_MEASURE_REGEX);
                        String secondMeasureValue = RegexUtils.regexWithRemove(additionalReq,
                                SECOND_MEASURE_REGEX, REMOVE_SECOND_MEASURE_REGEX);

                        targetCell.setCellValue(firstMeasureValue+FILLER+firstMeasureValue+FILLER+secondMeasureValue+
                                FILLER+lengthValue);
                    }

                    if(productTypeValue.equals(REBARS_COILS)) {
                        String profileName = rowFrom.getCell(acceptLibraryProfileNameIndex).
                                getStringCellValue();
                        final String PROFILE_REGEX = "(\\|[0-9]{1,2})";

                        String profile = RegexUtils.regex(profileName, PROFILE_REGEX);
                        targetCell.setCellValue(profile+" бунт");
                    }

                    if(productTypeValue.equals(COLD_ROLLED_SPECIAL_SECTIONS)) {
                        String specialProfileName = rowFrom.getCell(acceptLibraryProfileNameIndex).
                                getStringCellValue();
                        final String SPECIAL_PROFILE_REGEX = "(\\|[0-9*.]{1,20}\\|)";

                        String specialProfile = RegexUtils.regex(specialProfileName,SPECIAL_PROFILE_REGEX);

                        targetCell.setCellValue(specialProfile);
                    }

                    if(productTypeValue.equals(U_CHANNELS)) {
                        String additionalReq = rowFrom.getCell(acceptLibraryAdditionalReqIndex).
                                getStringCellValue();

                        DataFormatter dataFormatter = new DataFormatter();
                        String lengthValue = dataFormatter.formatCellValue(rowFrom.getCell(lengthHeaderAcceptIndex));

                        final String PROFILE_NUMBER_REGEX = "(Номер профиля горячекатаного проката=)([0-9.УВ]{1,5})";
                        final String REMOVE_PROFILE_NUMBER_REGEX = "(Номер профиля горячекатаного проката=)";

                        String profileNumberValue = RegexUtils.regexWithRemove(additionalReq,
                                PROFILE_NUMBER_REGEX, REMOVE_PROFILE_NUMBER_REGEX);

                        targetCell.setCellValue(profileNumberValue+FILLER+lengthValue);
                    }
                }
            }
        }
    }
}
