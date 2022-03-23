package com.tkachenko.buyerhelper.service.mmk;

import com.tkachenko.buyerhelper.utils.ExcelUtils;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.stereotype.Service;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

@Service
public class MmkService {

    private final String settingsFileName = "mmkToOtherFactorySetting.xlsx";
    private final Path mmkToOtherFactorySettings = Paths.get("./src/data").resolve(settingsFileName);
    private final String newSheetName = "OracleNewPage";
    private final int settingsOracleHeaderIndex = 2;
    private final int settingPasteCellIndex = 1;
    private final int settingFactoryHeaderIndex = 0;

    private final String priceHeader = "Цена";
    private final String factoryValue = "MMK";
    private final int yearValue = 2022;
    private final String factoryHeader = "Поставщик";
    private final String yearHeader = "Год акцепта";
    private int factoryColIndex = -1;
    private int yearColIndex = -1;

//    @Autowired
//    public MmkService(FileStorageProperties fileStorageProperties) {
//        this.mmkToOtherFactorySettings = Paths.get(fileStorageProperties.getUploadDir()).toAbsolutePath().normalize()
//                .resolve(settingsFileName);
//    }

    public void parseMmkToOtherFactoryFormat(Path fileMmkOraclePath, Path fileMmkAcceptLibraryPath,
                                             Path fileMmkDependenciesPath) {
        try {
            FileInputStream inputStreamSettings = new FileInputStream(mmkToOtherFactorySettings.toAbsolutePath()
                    .toString());
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
                if(cellFrom.getStringCellValue().equals(factoryHeader)) {
                    factoryColIndex = i + 1;
                }
                if(cellFrom.getStringCellValue().equals(yearHeader)) {
                    yearColIndex = i + 1;
                }
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
                        if(valueForSearch.equals(priceHeader)) {
                            cellTo.setCellValue(cellTo.getNumericCellValue() * 1.2);
                        }
                    }
                }

                if (factoryColIndex >= 0) {
                    newSheetRow.createCell(factoryColIndex).setCellValue(factoryValue);
                }
                if (yearColIndex >= 0) {
                    newSheetRow.createCell(yearColIndex).setCellValue(yearValue);
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

            MmkAcceptMonthParser mmkAcceptMonthParser = new MmkAcceptMonthParser(fileMmkOraclePath);
            mmkAcceptMonthParser.parseMonth();

            MmkProfileParser mmkProfileParser = new MmkProfileParser(fileMmkOraclePath, fileMmkAcceptLibraryPath);
            mmkProfileParser.parse();

            //MmkFormulasSetter mmkFormulasSetter = new MmkFormulasSetter(fileMmkOraclePath);
            //mmkFormulasSetter.setFormulas();

            MmkBranchSellTypeAndClientSetter mmkBranchSellTypeAndClientSetter =
                    new MmkBranchSellTypeAndClientSetter(fileMmkOraclePath, fileMmkDependenciesPath);
            mmkBranchSellTypeAndClientSetter.setBranchSellTypeAndClient();

        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    public void addOracleToSummaryFile(Path fileMmkOraclePath, Path fileSummaryPath) {
        final String ORACLE_ACCEPT_MONTH_COL_NAME = "Месяц акцепта";

        try {
            FileInputStream summaryFileInputStream = new FileInputStream(fileSummaryPath.toString());
            XSSFWorkbook summaryWorkbook = new XSSFWorkbook(summaryFileInputStream);
//          Перебрать все страницы и создать XSSFSheet
            int numberOfSheetsSummaryWorkbook = summaryWorkbook.getNumberOfSheets();

            FileInputStream oracleFileInputStream = new FileInputStream(fileMmkOraclePath.toString());
            XSSFWorkbook oracleWorkbook = new XSSFWorkbook(oracleFileInputStream);
            XSSFSheet newOracleSheet = oracleWorkbook.getSheet(newSheetName);
            XSSFRow newOracleHeader = newOracleSheet.getRow(newOracleSheet.getFirstRowNum());
            int newOracleAcceptMonthColIndex = ExcelUtils.findColumnByValue(newOracleHeader, ORACLE_ACCEPT_MONTH_COL_NAME);
            int oracleFirstRow = newOracleSheet.getFirstRowNum() + 1;
            int oracleLastRow = newOracleSheet.getLastRowNum();

            Map<MonthSheets, XSSFSheet> mapMonthSheetsXSSFSheet = new HashMap<>();
            CellCopyPolicy defaultCopyPolicy = new CellCopyPolicy();

            for (int i = 0; i < numberOfSheetsSummaryWorkbook; i++) {
                XSSFSheet currentSummarySheet = summaryWorkbook.getSheetAt(i);
                String currentSummarySheetName = currentSummarySheet.getSheetName();
                if(MonthSheets.findBySheetName(currentSummarySheetName) != null) {
                    MonthSheets currentSheetMonth = MonthSheets.findBySheetName(currentSummarySheetName);
                    mapMonthSheetsXSSFSheet.put(currentSheetMonth, currentSummarySheet);
                }
            }

            for(int j = oracleFirstRow; j <= oracleLastRow; j++) {
                XSSFRow currentOracleRow = newOracleSheet.getRow(j);
                XSSFCell oracleAcceptMonthCell = currentOracleRow.getCell(newOracleAcceptMonthColIndex);
                if (oracleAcceptMonthCell != null && oracleAcceptMonthCell.getCellType() != CellType.BLANK) {
                    int oracleAcceptMonthValue = (int) oracleAcceptMonthCell.getNumericCellValue();
                    MonthSheets oracleMonthSheet = MonthSheets.findByIntValue(oracleAcceptMonthValue);
                    if (mapMonthSheetsXSSFSheet.containsKey(oracleMonthSheet)) {
                        XSSFSheet targetSheet = mapMonthSheetsXSSFSheet.get(oracleMonthSheet);
                        int summaryDestinationRowIndex = targetSheet.getLastRowNum() + 1;
                        XSSFRow summaryDestinationRow = targetSheet.createRow(summaryDestinationRowIndex);
                        ExcelUtils.copyXSSFRow(currentOracleRow, summaryDestinationRow);
                    } else {
                        String newSheetName = oracleMonthSheet.getSheetName();
                        summaryWorkbook.createSheet(newSheetName);
                        XSSFSheet targetSheet = summaryWorkbook.getSheet(newSheetName);
                        mapMonthSheetsXSSFSheet.put(oracleMonthSheet, targetSheet);
                        int summaryDestinationHeaderIndex = targetSheet.getLastRowNum() + 1;
                        XSSFRow summaryDestinationHeader = targetSheet.createRow(summaryDestinationHeaderIndex);
                        XSSFRow summaryDestinationRow = targetSheet.createRow(summaryDestinationHeaderIndex+1);
                        ExcelUtils.copyXSSFRow(newOracleHeader, summaryDestinationHeader);
                        ExcelUtils.copyXSSFRow(currentOracleRow, summaryDestinationRow);
                    }
                }
            }

            FileOutputStream summaryFileOutputStream = new FileOutputStream(fileSummaryPath.toString());
            summaryWorkbook.write(summaryFileOutputStream);
            summaryWorkbook.close();

            summaryFileOutputStream.flush();
            summaryFileOutputStream.close();

            summaryFileInputStream.close();
            oracleFileInputStream.close();
        }

        catch (Exception e) {
            e.printStackTrace();
        }

    }
}

