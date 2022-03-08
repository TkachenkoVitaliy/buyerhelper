package com.tkachenko.buyerhelper.service.mmk;

import com.tkachenko.buyerhelper.property.FileStorageProperties;
import com.tkachenko.buyerhelper.utils.ExcelUtils;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Service;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.nio.file.Path;
import java.nio.file.Paths;

@Service
public class MmkService {

    private final Path mmkToOtherFactorySettings;
    private final String settingsFileName = "mmkToOtherFactorySetting.xlsx";
    private final String newSheetName = "OracleNewPage";
    private final int settingsOracleHeaderIndex = 2;
    private final int settingPasteCellIndex = 1;
    private final int settingFactoryHeaderIndex = 0;

    @Autowired
    public MmkService(FileStorageProperties fileStorageProperties) {
        this.mmkToOtherFactorySettings = Paths.get(fileStorageProperties.getUploadDir()).toAbsolutePath().normalize()
                .resolve(settingsFileName);
    }

    public void parseMmkToOtherFactoryFormat(Path fileMmkOraclePath) {
        try {
            FileInputStream inputStreamSettings = new FileInputStream(mmkToOtherFactorySettings.toAbsolutePath().toString());
            FileInputStream inputStreamMmk = new FileInputStream(fileMmkOraclePath.toAbsolutePath().toString());
            XSSFWorkbook settingsWorkbook = new XSSFWorkbook(inputStreamSettings);
            XSSFWorkbook mmkWorkbook = new XSSFWorkbook(inputStreamMmk);
            XSSFSheet settingsSheet = settingsWorkbook.getSheetAt(0);
            XSSFSheet mmkOracleSheet = mmkWorkbook.getSheetAt(0);
            XSSFSheet mmkNewSheet = mmkWorkbook.createSheet(newSheetName);
            int settingLastRowIndex = settingsSheet.getLastRowNum();
            Row rowNewSheetHeader = mmkNewSheet.createRow(0);

            for (int i = 0; i <= settingLastRowIndex; i++) {
                Cell cellFrom = settingsSheet.getRow(i).getCell(settingFactoryHeaderIndex);
                Cell cellTo = rowNewSheetHeader.createCell(i+1);
                ExcelUtils.copyCellValueXSSF((XSSFCell) cellFrom, (XSSFCell) cellTo);
            }

            int firstParseRowMmkIndex = mmkOracleSheet.getFirstRowNum()+1;
            int lastParseRowMmkIndex = mmkOracleSheet.getLastRowNum();

            Row headerOracleRow = mmkOracleSheet.getRow(firstParseRowMmkIndex-1);
            Row currentParseRow;
            for (int i = firstParseRowMmkIndex; i <=lastParseRowMmkIndex; i++) {
                currentParseRow = mmkOracleSheet.getRow(i);
                Row newSheetRow = mmkNewSheet.createRow(i);
                for (int k = 0; k <=settingLastRowIndex; k++) {
                    Cell cellForSearchCol = settingsSheet.getRow(k).getCell(settingsOracleHeaderIndex);
                    if(cellForSearchCol != null && cellForSearchCol.getCellType()!= CellType.BLANK) {
                        String valueForSearch = cellForSearchCol.getStringCellValue();
                        int colIndexForPaste = (int) settingsSheet.getRow(k).getCell(settingPasteCellIndex).getNumericCellValue();
                        int colIndexForCopy = ExcelUtils.findColIndexByStringValue(valueForSearch, headerOracleRow);
                        Cell cellFrom = currentParseRow.getCell(colIndexForCopy);
                        Cell cellTo = newSheetRow.createCell(colIndexForPaste);
                        if(cellFrom != null) ExcelUtils.copyCellValueXSSF((XSSFCell) cellFrom, (XSSFCell) cellTo);
                    }
                }
            }

            FileOutputStream outputStreamMmk = new FileOutputStream(fileMmkOraclePath.toAbsolutePath().toString());
            mmkWorkbook.write(outputStreamMmk);
            mmkWorkbook.close();
            settingsWorkbook.close();

            outputStreamMmk.flush();
            outputStreamMmk.close();
            inputStreamMmk.close();
            inputStreamSettings.close();

        } catch (Exception e) {
            e.printStackTrace();
        }
    }
}
