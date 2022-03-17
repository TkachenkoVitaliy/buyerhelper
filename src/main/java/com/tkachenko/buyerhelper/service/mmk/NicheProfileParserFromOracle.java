package com.tkachenko.buyerhelper.service.mmk;

import com.tkachenko.buyerhelper.utils.ExcelUtils;
import com.tkachenko.buyerhelper.utils.RegexUtils;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;


public class NicheProfileParserFromOracle {

    private final String ROLLED_TAPE = "Лента";
    private final String ROLLED_SHEET = "Лист";
    private final String ROLLED_COIL = "Рулон";
    private final String REBARS_BAR = "Профиль арматурный"; //TODO equals - not contains!!!
    private final String REBARS_COILS = "Профиль арматурный_моток";
    private final String SPECIAL_SECTIONS = "Спецпрофиль";
    private final String ANGLES = "Уголок г/к";
    private final String U_CHANNELS = "Швеллер г/к";
    private final String FILLER = "x";

    private final String NEW_PROFILE_HEADER = "Размеры/профиль";
    private final String OLD_HEADER_ADDITIONAL_REQ = "Тесты";
    private final String NEW_PRODUCT_TYPE_COLUMN_NAME = "Вид продукции";

    public void nicheParse (Row oldHeader, Row newHeader, Row oldRow, Row newRow) {

        int newProfileIndex = ExcelUtils.findColIndexByStringValue(NEW_PROFILE_HEADER, newHeader);
        int oldAdditionalReqIndex = ExcelUtils.findColIndexByStringValue(OLD_HEADER_ADDITIONAL_REQ, oldHeader);
        int newProductTypeIndex = ExcelUtils.findColIndexByStringValue(NEW_PRODUCT_TYPE_COLUMN_NAME, newHeader);

        if(oldRow.getCell(oldAdditionalReqIndex) == null) return;

        String productTypeValue = newRow.getCell(newProductTypeIndex).getStringCellValue();
        String additionalReq = oldRow.getCell(oldAdditionalReqIndex).getStringCellValue();

        Cell targetCell = newRow.getCell(newProfileIndex);
        if(targetCell == null) newRow.createCell(newProfileIndex);

        if(productTypeValue.contains(ROLLED_TAPE) | productTypeValue.contains(ROLLED_COIL)) {
            final String HEIGHT_REGEX = "(Толщина)\\s\\d{1,3}(.\\d[^,]{0,2})?";
            final String REMOVE_HEIGHT_REGEX = "(Толщина\\s)";
            final String WIDTH_REGEX = "(Ширина)\\s\\d{2,4}";
            final String REMOVE_WIDTH_REGEX = "(Ширина\\s)";

            String heightValue = RegexUtils.regexWithRemove(additionalReq, HEIGHT_REGEX, REMOVE_HEIGHT_REGEX);
            String widthValue = RegexUtils.regexWithRemove(additionalReq, WIDTH_REGEX, REMOVE_WIDTH_REGEX);

            targetCell.setCellValue(heightValue+FILLER+widthValue);
        }

        if(productTypeValue.contains(ROLLED_SHEET) | productTypeValue.contains(ANGLES)) {
            final String MEASURES_REGEX = "(Размер)\\s([0-9|.]{1,4})x([0-9|.]{1,4})x([0-9|.]{1,5})";
            final String REMOVE_MEASURES_REGEX = "(Размер\\s)";

            String measuresValue = RegexUtils.regexWithRemove(additionalReq, MEASURES_REGEX, REMOVE_MEASURES_REGEX);

            targetCell.setCellValue(measuresValue);
        }

        if(productTypeValue.contains(SPECIAL_SECTIONS) | productTypeValue.contains(U_CHANNELS) |
                productTypeValue.equals(REBARS_BAR)) {
            final String PROFILE_REGEX = "(Номер профиля )([0-9УВх]{1,100})";
            final String REMOVE_PROFILE_REGEX = "(Номер профиля )";
            final String LENGTH_REGEX = "(Длина )([0-9]{1,100})";
            final String REMOVE_LENGTH_REGEX = "(Длина )";

            String profileValue = RegexUtils.regexWithRemove(additionalReq, PROFILE_REGEX, REMOVE_PROFILE_REGEX);
            String lengthValue = RegexUtils.regexWithRemove(additionalReq, LENGTH_REGEX, REMOVE_LENGTH_REGEX);

            targetCell.setCellValue(profileValue+FILLER+lengthValue);
        }

        if(productTypeValue.contains(REBARS_COILS)) {
            final String PROFILE_REGEX = "(Номер профиля )([0-9УВх]{1,100})";
            final String REMOVE_PROFILE_REGEX = "(Номер профиля )";

            String profileValue = RegexUtils.regexWithRemove(additionalReq, PROFILE_REGEX, REMOVE_PROFILE_REGEX);

            targetCell.setCellValue(profileValue+" бунт");
        }
    }
}
