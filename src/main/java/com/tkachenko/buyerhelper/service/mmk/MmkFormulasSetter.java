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

public class MmkFormulasSetter {

    private final String ACCEPT_COST_NAME = "Стоимость, руб";
    private final String SHIPPED_COST_NAME = "Стоимость отгр, руб";
    private final String FINAL_COST_NAME = "Итоговая стоимость, тн";
    private final String SHEET_NAME = "OracleNewPage";
    private final String PRICE_NAME = "Цена с НДС, руб/тн";
    private final String ACCEPT_WEIGHT_NAME = "Акцепт, тн";
    private final String SHIPPED_WEIGHT_NAME = "Отгруженно, тн";
    private final String NEW_PRICE_NAME = "Пересмотр, руб/тн";

    private final Path mmkOracleFile;

    public MmkFormulasSetter(Path mmkOracleFile) {
        this.mmkOracleFile = mmkOracleFile;
    }

    public void setFormulas() {
        try {
            FileInputStream mmkOracleInputStream = new FileInputStream(mmkOracleFile.toString());
            XSSFWorkbook mmkOracleWorkbook = new XSSFWorkbook(mmkOracleInputStream);
            XSSFSheet mmkOracleNewSheet = mmkOracleWorkbook.getSheet(SHEET_NAME);
            int headerRowIndex = mmkOracleNewSheet.getFirstRowNum();
            int firstRowIndex = headerRowIndex + 1;
            int lastRowIndex = mmkOracleNewSheet.getLastRowNum();
            Row headerRow = mmkOracleNewSheet.getRow(headerRowIndex);
            int acceptCostColIndex = ExcelUtils.findColumnByValue(headerRow, ACCEPT_COST_NAME); //target
            int shippedCostColIndex = ExcelUtils.findColumnByValue(headerRow, SHIPPED_COST_NAME); //target
            int finalCostColIndex = ExcelUtils.findColumnByValue(headerRow, FINAL_COST_NAME); //target
            int priceColIndex = ExcelUtils.findColumnByValue(headerRow, PRICE_NAME);
            String priceColForFormula = ExcelUtils.getExcelColAddress(priceColIndex);
            int acceptWeightColIndex = ExcelUtils.findColumnByValue(headerRow, ACCEPT_WEIGHT_NAME);
            String acceptWeightColForFormula = ExcelUtils.getExcelColAddress(acceptWeightColIndex);
            int shippedWeightColIndex = ExcelUtils.findColumnByValue(headerRow, SHIPPED_WEIGHT_NAME);
            String shippedWeightColForFormula = ExcelUtils.getExcelColAddress(shippedWeightColIndex);
            int newPriceColIndex = ExcelUtils.findColumnByValue(headerRow, NEW_PRICE_NAME);
            String newPriceColForFormula = ExcelUtils.getExcelColAddress(newPriceColIndex);

            for (int i = firstRowIndex; i <= lastRowIndex; i++) {
                Row currentRow = mmkOracleNewSheet.getRow(i);
                int rowNumForFormula = i + 1;
                Cell acceptCostCell = currentRow.createCell(acceptCostColIndex, CellType.FORMULA);
                Cell shippedCostCell = currentRow.createCell(shippedCostColIndex, CellType.FORMULA);
                Cell finalCostCell = currentRow.createCell(finalCostColIndex, CellType.FORMULA);
                acceptCostCell.setCellFormula(priceColForFormula + rowNumForFormula + "*"
                        + acceptWeightColForFormula + rowNumForFormula);
                shippedCostCell.setCellFormula(priceColForFormula + rowNumForFormula + "*"
                        + shippedWeightColForFormula + rowNumForFormula);
                finalCostCell.setCellFormula("IF(" +  newPriceColForFormula + rowNumForFormula + "=0,"
                        + priceColForFormula + rowNumForFormula + "*" + shippedWeightColForFormula + rowNumForFormula
                        + "," + newPriceColForFormula + rowNumForFormula + "*"
                        + shippedWeightColForFormula + rowNumForFormula + ")");
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
}
