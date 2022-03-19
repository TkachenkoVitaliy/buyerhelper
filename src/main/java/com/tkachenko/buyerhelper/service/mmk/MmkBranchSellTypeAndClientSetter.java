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

public class MmkBranchSellTypeAndClientSetter {
    private final Path fileMmkOraclePath;
    private final Path fileMmkDependenciesPath;
    private final String ORACLE_SHEET_NAME = "OracleNewPage";
    private final String TRANSIT_SALES_DEP_SHEET_NAME = "Прямые транзиты";
    private final String CONTAINERS_DEP_SHEET_NAME = "Контейнеры";
    private final String EXCEPTIONS_DEP_SHEET_NAME = "Исключения";

    private final String CONSIGNEE_HEADER_COL_NAME = "Грузополучатель";
    private final String BRANCH_HEADER_COL_NAME = "База";
    private final String SELL_TYPE_HEADER_COL_NAME= "Вид поставки";
    private final String CLIENT_HEADER_COL_NAME = "Транзитн. Клиент";

    public MmkBranchSellTypeAndClientSetter(Path fileMmkOraclePath, Path fileMmkDependenciesPath) {
        this.fileMmkOraclePath = fileMmkOraclePath;
        this.fileMmkDependenciesPath = fileMmkDependenciesPath;
    }

    public void setBranchSellTypeAndClient() {
        try {
            FileInputStream oracleInputStream = new FileInputStream(fileMmkOraclePath.toString());
            XSSFWorkbook oracleWorkbook = new XSSFWorkbook(oracleInputStream);
            XSSFSheet oracleSheet = oracleWorkbook.getSheet(ORACLE_SHEET_NAME);
            Row oracleHeader = oracleSheet.getRow(oracleSheet.getFirstRowNum());
            int firstRowNum = oracleSheet.getFirstRowNum() + 1;
            int lastRowNum = oracleSheet.getLastRowNum();

            FileInputStream dependenciesInputStream = new FileInputStream(fileMmkDependenciesPath.toString());
            XSSFWorkbook dependenciesWorkbook = new XSSFWorkbook(dependenciesInputStream);

            int oracleConsigneeColIndex = ExcelUtils.findColumnByValue(oracleHeader, CONSIGNEE_HEADER_COL_NAME);
            int oracleBranchColIndex = ExcelUtils.findColumnByValue(oracleHeader, BRANCH_HEADER_COL_NAME);
            int oracleSellTypeColIndex = ExcelUtils.findColumnByValue(oracleHeader, SELL_TYPE_HEADER_COL_NAME);
            int oracleClientColIndex = ExcelUtils.findColumnByValue(oracleHeader, CLIENT_HEADER_COL_NAME);

            for (int i = firstRowNum; i <= lastRowNum; i++) {
                Row currentOracleRow = oracleSheet.getRow(i);
                setFromDefaultStorage(currentOracleRow, oracleConsigneeColIndex, oracleBranchColIndex,
                        oracleSellTypeColIndex, dependenciesWorkbook);
                setFromTransits(currentOracleRow, oracleConsigneeColIndex, oracleBranchColIndex, oracleSellTypeColIndex,
                        oracleClientColIndex, dependenciesWorkbook);

                //some methods

            }

            FileOutputStream oracleOutputStream = new FileOutputStream(fileMmkOraclePath.toString());
            oracleWorkbook.write(oracleOutputStream);
            oracleWorkbook.close();

            oracleOutputStream.flush();
            oracleOutputStream.close();
            oracleInputStream.close();
            dependenciesInputStream.close();

        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    private void setFromDefaultStorage(Row oracleCurrentRow, int oracleConsigneeColIndex, int oracleBranchColIndex,
                                       int oracleSellTypeColIndex ,XSSFWorkbook dependenciesWorkbook) {
        final String DEFAULT_STORAGE_DEP_SHEET_NAME = "Склады";
        XSSFSheet defaultStorageSheet = dependenciesWorkbook.getSheet(DEFAULT_STORAGE_DEP_SHEET_NAME);
        Row dependencyHeader = defaultStorageSheet.getRow(defaultStorageSheet.getFirstRowNum());
        int firstDependencyRow = defaultStorageSheet.getFirstRowNum()+1;
        int lastDependencyRow = defaultStorageSheet.getLastRowNum();

        int dependencyConsigneeColIndex = ExcelUtils.findColumnByValue(dependencyHeader, CONSIGNEE_HEADER_COL_NAME);
        int dependencyBranchColIndex = ExcelUtils.findColumnByValue(dependencyHeader, BRANCH_HEADER_COL_NAME);
        int dependencySellTypeColIndex = ExcelUtils.findColumnByValue(dependencyHeader, SELL_TYPE_HEADER_COL_NAME);

        Cell oracleConsigneeCell = oracleCurrentRow.getCell(oracleConsigneeColIndex);

        if (oracleConsigneeCell != null) {
            for(int j = firstDependencyRow; j <= lastDependencyRow; j++) {
                Row currentDependencyRow = defaultStorageSheet.getRow(j);
                Cell dependencyConsigneeCell = currentDependencyRow.getCell(dependencyConsigneeColIndex);
                String oracleConsigneeValue = oracleConsigneeCell.getStringCellValue().replaceAll("\"", "");
                if(oracleConsigneeValue.equals(dependencyConsigneeCell.getStringCellValue())) {
                    String branchValue = currentDependencyRow.getCell(dependencyBranchColIndex).getStringCellValue();
                    String sellTypeValue = currentDependencyRow.getCell(dependencySellTypeColIndex).getStringCellValue();

                    if(oracleCurrentRow.getCell(oracleBranchColIndex) == null) oracleCurrentRow.createCell(oracleBranchColIndex);
                    if(oracleCurrentRow.getCell(oracleSellTypeColIndex) == null) oracleCurrentRow.createCell(oracleSellTypeColIndex);
                    Cell oracleBranchCell = oracleCurrentRow.getCell(oracleBranchColIndex);
                    Cell oracleSellTypeCell = oracleCurrentRow.getCell(oracleSellTypeColIndex);

                    oracleBranchCell.setCellValue(branchValue);
                    oracleSellTypeCell.setCellValue(sellTypeValue);
                }
            }
        }


    }

    public void setFromTransits(Row oracleCurrentRow, int oracleConsigneeColIndex, int oracleBranchColIndex,
                     int oracleSellTypeColIndex, int oracleClientColIndex ,XSSFWorkbook dependenciesWorkbook) {
        final String TRANSIT_SALES_DEP_SHEET_NAME = "Прямые транзиты";
        XSSFSheet transitSheet = dependenciesWorkbook.getSheet(TRANSIT_SALES_DEP_SHEET_NAME);
        Row dependencyHeader = transitSheet.getRow(transitSheet.getFirstRowNum());
        int firstDependencyRow = transitSheet.getFirstRowNum()+1;
        int lastDependencyRow = transitSheet.getLastRowNum();

        int dependencyConsigneeColIndex = ExcelUtils.findColumnByValue(dependencyHeader, CONSIGNEE_HEADER_COL_NAME);
        int dependencyBranchColIndex = ExcelUtils.findColumnByValue(dependencyHeader, BRANCH_HEADER_COL_NAME);
        int dependencySellTypeColIndex = ExcelUtils.findColumnByValue(dependencyHeader, SELL_TYPE_HEADER_COL_NAME);
        int dependencyClientColIndex = ExcelUtils.findColumnByValue(dependencyHeader, CLIENT_HEADER_COL_NAME);

        Cell oracleConsigneeCell = oracleCurrentRow.getCell(oracleConsigneeColIndex);

        if (oracleConsigneeCell != null) {
            for(int k = firstDependencyRow; k <= lastDependencyRow; k++) {
                Row currentDependencyRow = transitSheet.getRow(k);
                Cell dependencyConsigneeCell = currentDependencyRow.getCell(dependencyConsigneeColIndex);
                String oracleConsigneeValue = oracleConsigneeCell
                        .getStringCellValue().replaceAll("\"", "");
                if(oracleConsigneeValue.equals(dependencyConsigneeCell.getStringCellValue())) {
                    String branchValue = currentDependencyRow.getCell(dependencyBranchColIndex).getStringCellValue();
                    String sellTypeValue = currentDependencyRow.getCell(dependencySellTypeColIndex).getStringCellValue();
                    String clientValue = null;
                    if(currentDependencyRow.getCell(dependencyClientColIndex) != null &&
                            currentDependencyRow.getCell(dependencyClientColIndex).getCellType() != CellType.BLANK) {
                        clientValue = currentDependencyRow.getCell(dependencyClientColIndex).getStringCellValue();
                    }

                    if(oracleCurrentRow.getCell(oracleBranchColIndex) == null) oracleCurrentRow.createCell(oracleBranchColIndex);
                    if(oracleCurrentRow.getCell(oracleSellTypeColIndex) == null) oracleCurrentRow.createCell(oracleSellTypeColIndex);
                    if(oracleCurrentRow.getCell(oracleClientColIndex) == null) oracleCurrentRow.createCell(oracleClientColIndex);

                    Cell oracleBranchCell = oracleCurrentRow.getCell(oracleBranchColIndex);
                    Cell oracleSellTypeCell = oracleCurrentRow.getCell(oracleSellTypeColIndex);
                    Cell oracleClientCell = oracleCurrentRow.getCell(oracleClientColIndex);

                    oracleBranchCell.setCellValue(branchValue);
                    oracleSellTypeCell.setCellValue(sellTypeValue);
                    if(clientValue != null) oracleClientCell.setCellValue(clientValue);
                }
            }
        }
    }
}