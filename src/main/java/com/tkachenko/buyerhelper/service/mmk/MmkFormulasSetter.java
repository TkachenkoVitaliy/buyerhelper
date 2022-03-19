package com.tkachenko.buyerhelper.service.mmk;

import com.tkachenko.buyerhelper.utils.ExcelUtils;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.nio.file.Path;

public class MmkFormulasSetter {

    private final String ACCEPT_COST_NAME = "Стоимость, руб";
    private final String SHIPPED_COST_NAME = "Стоимость отгр, руб";
    private final String FINAL_COST_NAME = "Итоговая стоимость, тн";
    private final String SHEET_NAME = "OracleNewPage";

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
            int acceptCostColIndex = ExcelUtils.findColumnByValue(headerRow, ACCEPT_COST_NAME);
            int shippedCostColIndex = ExcelUtils.findColumnByValue(headerRow, SHIPPED_COST_NAME);
            int finalCostColIndex = ExcelUtils.findColumnByValue(headerRow, FINAL_COST_NAME);


        } catch (Exception e) {
            e.printStackTrace();
        }

    }
}
