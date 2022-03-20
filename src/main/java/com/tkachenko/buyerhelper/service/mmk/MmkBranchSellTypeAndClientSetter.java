package com.tkachenko.buyerhelper.service.mmk;

import com.tkachenko.buyerhelper.utils.ExcelUtils;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import javax.xml.crypto.Data;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.nio.file.Path;

public class MmkBranchSellTypeAndClientSetter {
    private final Path fileMmkOraclePath;
    private final Path fileMmkDependenciesPath;
    private final String ORACLE_SHEET_NAME = "OracleNewPage";

    private final String CONSIGNEE_HEADER_COL_NAME = "Грузополучатель";
    private final String BRANCH_HEADER_COL_NAME = "База";
    private final String SELL_TYPE_HEADER_COL_NAME= "Вид поставки";
    private final String CLIENT_HEADER_COL_NAME = "Транзитн. Клиент";
    private final String STATION_HEADER_COL_NAME = "Станция";
    private final String ORDER_HEADER_COL_NAME = "Номер СПЕЦ";
    private final String POSITION_HEADER_COL_NAME = "Номер позиции";

    public MmkBranchSellTypeAndClientSetter(Path fileMmkOraclePath, Path fileMmkDependenciesPath) {
        this.fileMmkOraclePath = fileMmkOraclePath;
        this.fileMmkDependenciesPath = fileMmkDependenciesPath;
    }

    public void setBranchSellTypeAndClient() {
        try {
            FileInputStream oracleInputStream = new FileInputStream(fileMmkOraclePath.toString());
            XSSFWorkbook oracleWorkbook = new XSSFWorkbook(oracleInputStream);
            XSSFSheet oldOracleSheet = oracleWorkbook.getSheetAt(0);
            Row oldOracleHeader = oldOracleSheet.getRow(oldOracleSheet.getFirstRowNum());
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
            int oracleOrderColIndex = ExcelUtils.findColumnByValue(oracleHeader, ORDER_HEADER_COL_NAME);
            int oraclePositionColIndex = ExcelUtils.findColumnByValue(oracleHeader, POSITION_HEADER_COL_NAME);
            int oldOracleStationColIndex = ExcelUtils.findColumnByValue(oldOracleHeader, STATION_HEADER_COL_NAME);

            for (int i = firstRowNum; i <= lastRowNum; i++) {
                Row currentOracleRow = oracleSheet.getRow(i);
                setFromDefaultStorage(currentOracleRow, oracleConsigneeColIndex, oracleBranchColIndex,
                        oracleSellTypeColIndex, dependenciesWorkbook);
                setFromTransits(currentOracleRow, oracleConsigneeColIndex, oracleBranchColIndex, oracleSellTypeColIndex,
                        oracleClientColIndex, dependenciesWorkbook);
                Row currentOldOracleRow = oldOracleSheet.getRow(i);
                setFromContainers(currentOldOracleRow, oldOracleStationColIndex, currentOracleRow,
                        oracleConsigneeColIndex, oracleBranchColIndex, oracleSellTypeColIndex, oracleClientColIndex,
                        dependenciesWorkbook);
                setExceptions (currentOracleRow, oracleOrderColIndex, oraclePositionColIndex, oracleBranchColIndex,
                        oracleSellTypeColIndex, oracleClientColIndex, dependenciesWorkbook);
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

    public void setFromContainers(Row oldOracleCurrentRow, int oldOracleStationColIndex, Row oracleCurrentRow,
                                  int oracleConsigneeColIndex, int oracleBranchColIndex, int oracleSellTypeColIndex,
                                  int oracleClientColIndex, XSSFWorkbook dependenciesWorkbook) {
        final String CONTAINERS_DEP_SHEET_NAME = "Контейнеры";
        XSSFSheet containersSheet = dependenciesWorkbook.getSheet(CONTAINERS_DEP_SHEET_NAME);
        Row dependencyHeader = containersSheet.getRow(containersSheet.getFirstRowNum());
        int firstDependencyRow = containersSheet.getFirstRowNum()+1;
        int lastDependencyRow = containersSheet.getLastRowNum();

        int dependencyStationColIndex = ExcelUtils.findColumnByValue(dependencyHeader, STATION_HEADER_COL_NAME);
        int dependencyConsigneeColIndex = ExcelUtils.findColumnByValue(dependencyHeader, CONSIGNEE_HEADER_COL_NAME);
        int dependencyBranchColIndex = ExcelUtils.findColumnByValue(dependencyHeader, BRANCH_HEADER_COL_NAME);
        int dependencySellTypeColIndex = ExcelUtils.findColumnByValue(dependencyHeader, SELL_TYPE_HEADER_COL_NAME);
        int dependencyClientColIndex = ExcelUtils.findColumnByValue(dependencyHeader, CLIENT_HEADER_COL_NAME);

        Cell oracleConsigneeCell = oracleCurrentRow.getCell(oracleConsigneeColIndex);
        Cell oldOracleStationCell = oldOracleCurrentRow.getCell(oldOracleStationColIndex);

        if(oracleConsigneeCell != null && oldOracleStationCell != null) {
            for(int l = firstDependencyRow; l <= lastDependencyRow; l++) {
                Row currentDependencyRow = containersSheet.getRow(l);
                Cell dependencyConsigneeCell = currentDependencyRow.getCell(dependencyConsigneeColIndex);
                Cell dependencyStationCell = currentDependencyRow.getCell(dependencyStationColIndex);

                String oracleConsigneeValue = oracleConsigneeCell
                        .getStringCellValue().replaceAll("\"", "");
                String oldOracleStationValue = oldOracleStationCell.getStringCellValue();

                if(oracleConsigneeValue.equals(dependencyConsigneeCell.getStringCellValue()) &&
                oldOracleStationValue.equals(dependencyStationCell.getStringCellValue())) {
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

    public void setExceptions (Row oracleCurrentRow, int oracleOrderColIndex, int oraclePositionColIndex,
                               int oracleBranchColIndex, int oracleSellTypeColIndex, int oracleClientColIndex,
                               XSSFWorkbook dependenciesWorkbook) {
        final String EXCEPTIONS_DEP_SHEET_NAME = "Исключения";
        XSSFSheet exceptionsSheet = dependenciesWorkbook.getSheet(EXCEPTIONS_DEP_SHEET_NAME);
        Row dependencyHeader = exceptionsSheet.getRow(exceptionsSheet.getFirstRowNum());
        int firstDependencyRow = exceptionsSheet.getFirstRowNum()+1;
        int lastDependencyRow = exceptionsSheet.getLastRowNum();

        int dependencyOrderColIndex = ExcelUtils.findColumnByValue(dependencyHeader, ORDER_HEADER_COL_NAME);
        int dependencyPositionColIndex = ExcelUtils.findColumnByValue(dependencyHeader, POSITION_HEADER_COL_NAME);
        int dependencyBranchColIndex = ExcelUtils.findColumnByValue(dependencyHeader, BRANCH_HEADER_COL_NAME);
        int dependencySellTypeColIndex = ExcelUtils.findColumnByValue(dependencyHeader, SELL_TYPE_HEADER_COL_NAME);
        int dependencyClientColIndex = ExcelUtils.findColumnByValue(dependencyHeader, CLIENT_HEADER_COL_NAME);

        Cell oracleOrderCell = oracleCurrentRow.getCell(oracleOrderColIndex);
        Cell oraclePositionCell = oracleCurrentRow.getCell(oraclePositionColIndex);

        DataFormatter formatter = new DataFormatter();

        String oracleOrderValue = formatter.formatCellValue(oracleOrderCell);
        String oraclePositionValue = formatter.formatCellValue(oraclePositionCell);

        for(int n = firstDependencyRow; n <= lastDependencyRow; n++) {
            Row dependencyCurrentRow = exceptionsSheet.getRow(n);
            Cell dependencyOrderCell = dependencyCurrentRow.getCell(dependencyOrderColIndex);
            Cell dependencyPositionCell = dependencyCurrentRow.getCell(dependencyPositionColIndex);

            String dependencyOrderValue = formatter.formatCellValue(dependencyOrderCell);
            String dependencyPositionValue = formatter.formatCellValue(dependencyPositionCell);

            if(oracleOrderValue.equals(dependencyOrderValue)) {
                if(dependencyPositionValue.equals("0")) {
                    String branchValue = dependencyCurrentRow.getCell(dependencyBranchColIndex).getStringCellValue();
                    String sellTypeValue = dependencyCurrentRow.getCell(dependencySellTypeColIndex).getStringCellValue();
                    String clientValue = null;
                    if(dependencyCurrentRow.getCell(dependencyClientColIndex) != null &&
                            dependencyCurrentRow.getCell(dependencyClientColIndex).getCellType() != CellType.BLANK) {
                        clientValue = dependencyCurrentRow.getCell(dependencyClientColIndex).getStringCellValue();
                    }

                    if(oracleCurrentRow.getCell(oracleBranchColIndex) == null)
                        oracleCurrentRow.createCell(oracleBranchColIndex);
                    if(oracleCurrentRow.getCell(oracleSellTypeColIndex) == null)
                        oracleCurrentRow.createCell(oracleSellTypeColIndex);
                    if(oracleCurrentRow.getCell(oracleClientColIndex) == null)
                        oracleCurrentRow.createCell(oracleClientColIndex);

                    Cell oracleBranchCell = oracleCurrentRow.getCell(oracleBranchColIndex);
                    Cell oracleSellTypeCell = oracleCurrentRow.getCell(oracleSellTypeColIndex);
                    Cell oracleClientCell = oracleCurrentRow.getCell(oracleClientColIndex);

                    oracleBranchCell.setCellValue(branchValue);
                    oracleSellTypeCell.setCellValue(sellTypeValue);
                    if(clientValue != null) oracleClientCell.setCellValue(clientValue);
                } else {
                    if(oraclePositionValue.equals(dependencyPositionValue)) {
                        String branchValue = dependencyCurrentRow.getCell(dependencyBranchColIndex).getStringCellValue();
                        String sellTypeValue = dependencyCurrentRow.getCell(dependencySellTypeColIndex).getStringCellValue();
                        String clientValue = null;
                        if(dependencyCurrentRow.getCell(dependencyClientColIndex) != null &&
                                dependencyCurrentRow.getCell(dependencyClientColIndex).getCellType() != CellType.BLANK) {
                            clientValue = dependencyCurrentRow.getCell(dependencyClientColIndex).getStringCellValue();
                        }

                        if(oracleCurrentRow.getCell(oracleBranchColIndex) == null)
                            oracleCurrentRow.createCell(oracleBranchColIndex);
                        if(oracleCurrentRow.getCell(oracleSellTypeColIndex) == null)
                            oracleCurrentRow.createCell(oracleSellTypeColIndex);
                        if(oracleCurrentRow.getCell(oracleClientColIndex) == null)
                            oracleCurrentRow.createCell(oracleClientColIndex);

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
}
