package com.tkachenko.buyerhelper.service;

import com.tkachenko.buyerhelper.entity.FileEntity;
import com.tkachenko.buyerhelper.property.FileStorageProperties;
import com.tkachenko.buyerhelper.service.mmk.MonthSheets;
import com.tkachenko.buyerhelper.utils.ExcelUtils;
import com.tkachenko.buyerhelper.utils.FileUtils;
import org.apache.poi.openxml4j.util.ZipSecureFile;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.*;
import org.springframework.stereotype.Service;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.nio.file.*;
import java.util.ArrayList;
import java.util.List;

@Service
public class SummarySplitter {

    private final Path fileStorageLocation;
    private final FileDBService fileDBService;

    public SummarySplitter (FileStorageProperties fileStorageProperties, FileDBService fileDBService) {
        this.fileDBService = fileDBService;
        this.fileStorageLocation = Paths.get(fileStorageProperties.getUploadDir()).toAbsolutePath().normalize();
    }

    private final static String KRASNODAR = "Краснодар";
    private final static String ROSTOV = "Ростов";
    private final static String MOSCOW = "Москва";
    private final static String NOVGOROD = "Новгород";
    private final static String KAZAN = "Казань";
    private final static String NAB_CHELNY = "Наб. Челны";
    private final static String IZHEVSK = "Ижевск";
    private final static String PERM = "Пермь";
    private final static String UFA = "Уфа";
    private final static String CHELYABINSK = "Челябинск";
    private final static String YEKATERINBURG = "Екатеринбург";
    private final static String TUMEN = "Тюмень";
    private final static String OMSK = "Омск";
    private final static String NOVOSIBIRSK = "Новосибирск";
    private final static String NOVOKUZNETSK = "Новокузнецк";
    private final static String EXTENSION = ".xlsx";
    private final static String SUMMARY = "SUMMARY.xlsx";
    private final static String NEW_SUMMARY = "new" + SUMMARY;
    private final static String DIRECTORY_NAME = "forZip";
    private final static String[] citiesArray = {KRASNODAR, ROSTOV, MOSCOW, NOVGOROD, KAZAN, NAB_CHELNY, IZHEVSK, PERM,
    UFA, CHELYABINSK, YEKATERINBURG, TUMEN, OMSK, NOVOSIBIRSK, NOVOKUZNETSK};

    public void splitFiles() {
        Path newSummaryFilePath = copySummaryFile();

        for (String cityName : citiesArray) {
            Path targetFilePath = splitNewSummaryFile(newSummaryFilePath, cityName);
            deleteUnnecessaryBranches(targetFilePath, cityName);
        }
        try {
            Files.deleteIfExists(newSummaryFilePath);
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    private Path copySummaryFile() {
        final String PRICE_NAME = "Цена с НДС, руб/тн";
        final String ACCEPT_COST_NAME = "Стоимость, руб";
        final String SHIPPED_COST_NAME = "Стоимость отгр, руб";
        final String NEW_PRICE_NAME = "Пересмотр, руб/тн";
        final String FINAL_COST_NAME = "Итоговая стоимость, тн";

        FileEntity summaryFileEntity = fileDBService.getActualFileByStorageName(SUMMARY);
        Path oldSummaryFilePath = FileUtils.getEntityPath(fileStorageLocation, summaryFileEntity);
        Path targetDirectoryPath = fileStorageLocation.resolve(DIRECTORY_NAME);
        try {
            Files.createDirectories(targetDirectoryPath);
        } catch (IOException e) {
            e.printStackTrace();
        }
        Path newSummaryFilePath = targetDirectoryPath.resolve(NEW_SUMMARY);

        try {
            ZipSecureFile.setMinInflateRatio(0);
            FileInputStream oldSummaryInputStream = new FileInputStream(oldSummaryFilePath.toString());
            XSSFWorkbook oldWorkbook = new XSSFWorkbook(oldSummaryInputStream);
            XSSFRow sourceHeaderStyle = null;

            XSSFWorkbook newWorkbook = new XSSFWorkbook();
            CellCopyPolicy defaultCopyPolicy = new CellCopyPolicy();
            for (Sheet oldSheet : oldWorkbook) {
                XSSFSheet oldMonthSheet = (XSSFSheet) oldSheet;
                XSSFSheet newMonthSheet = newWorkbook.createSheet(oldMonthSheet.getSheetName());
                int headerRowNum = oldMonthSheet.getFirstRowNum();
                int lastRowNum = oldMonthSheet.getLastRowNum();
                XSSFRow oldHeaderRow = oldMonthSheet.getRow(headerRowNum);
                int firstColIndex = oldHeaderRow.getFirstCellNum();
                int lastColIndex = oldHeaderRow.getLastCellNum() -1;
                int priceColIndex = ExcelUtils.findColumnByValue(oldHeaderRow, PRICE_NAME);
                int acceptCostColIndex = ExcelUtils.findColumnByValue(oldHeaderRow, ACCEPT_COST_NAME);
                int shippedCostColIndex = ExcelUtils.findColumnByValue(oldHeaderRow, SHIPPED_COST_NAME);
                int newPriceColIndex = ExcelUtils.findColumnByValue(oldHeaderRow, NEW_PRICE_NAME);
                int finalCostColIndex = ExcelUtils.findColumnByValue(oldHeaderRow, FINAL_COST_NAME);

                for (int i = headerRowNum; i <= lastRowNum; i++) {
                    XSSFRow oldRow = oldMonthSheet.getRow(i);
                    XSSFRow newRow = newMonthSheet.createRow(i);
                    int targetColIndex = firstColIndex;
                    for (int j= firstColIndex; j <= lastColIndex; j++) {
                        if (j != priceColIndex && j !=acceptCostColIndex && j != shippedCostColIndex
                                && j != newPriceColIndex && j != finalCostColIndex ) {
                            XSSFCell oldCell = oldRow.getCell(j);
                            XSSFCell newCell = newRow.createCell(targetColIndex);
                            newCell.copyCellFrom(oldCell, defaultCopyPolicy);
                            targetColIndex++;
                        }
                    }
                }

                if(newMonthSheet.getSheetName().equals(MonthSheets.JANUARY.getSheetName())) {
                    sourceHeaderStyle = newMonthSheet.getRow(newMonthSheet.getFirstRowNum());
                } else {
                    if(sourceHeaderStyle != null)  {
                        ExcelUtils.copyRowStyle(sourceHeaderStyle, newMonthSheet.getRow(newMonthSheet.getFirstRowNum()));
                    }
                }
            }

            FileOutputStream fileOutputStream = new FileOutputStream(newSummaryFilePath.toString());
            newWorkbook.write(fileOutputStream);
            fileOutputStream.flush();
            fileOutputStream.close();
            newWorkbook.close();
            oldWorkbook.close();
            oldSummaryInputStream.close();

        } catch (Exception e) {
            e.printStackTrace();
        }


        return newSummaryFilePath;
    }

    private Path splitNewSummaryFile(Path newSummaryFilePath, String cityName) {
        Path targetDirectoryPath = fileStorageLocation.resolve(DIRECTORY_NAME);
        Path targetFilePath = targetDirectoryPath.resolve(cityName + EXTENSION);

        try {
            Files.createDirectories(targetDirectoryPath);
            Files.copy(newSummaryFilePath,targetFilePath, StandardCopyOption.REPLACE_EXISTING);
        } catch (IOException e) {
            e.printStackTrace();
        }

        return targetFilePath;
    }

    private void deleteUnnecessaryBranches(Path targetFilePath, String cityName) {
        final String BRANCH_HEADER_NAME = "База";
        System.out.println(cityName);
        try {
            ZipSecureFile.setMinInflateRatio(0);
            FileInputStream fileInputStream = new FileInputStream(targetFilePath.toString());
            XSSFWorkbook branchWorkbook = new XSSFWorkbook(fileInputStream);
            for (Sheet monthSheet : branchWorkbook) {
                int headerIndex = monthSheet.getFirstRowNum();
                int firstRowIndex = headerIndex + 1;
                int lastRowIndex = monthSheet.getLastRowNum();
                Row headerRow = monthSheet.getRow(headerIndex);
                int branchColIndex = ExcelUtils.findColumnByValue(headerRow, BRANCH_HEADER_NAME);

                List<Integer> rowIndexesList = new ArrayList<>();
                rowIndexesList.add(headerIndex);
                for(int i = firstRowIndex; i <= lastRowIndex ; i++) {
                    XSSFRow currentRow = (XSSFRow) monthSheet.getRow(i);
                    if(currentRow != null) {
                        XSSFCell branchCell = currentRow.getCell(branchColIndex);
                        if(branchCell != null && branchCell.getCellType() != CellType.BLANK
                                && branchCell.getStringCellValue().equals(cityName)) {
                            rowIndexesList.add(i);
                        }
                    }
                }


                for (int i = 0; i < rowIndexesList.size(); i++) {
                    int currentListIndex = i;
                    int rowIndexFromList = rowIndexesList.get(i);
                    if (rowIndexFromList > currentListIndex) monthSheet.shiftRows(rowIndexFromList, rowIndexFromList,
                            currentListIndex-rowIndexFromList);
                }

                for (int i=rowIndexesList.size(); i <= monthSheet.getLastRowNum(); i++) {
                    Row currentRow = monthSheet.getRow(i);
                    if(currentRow != null) monthSheet.removeRow(currentRow);
                }
            }

            FileOutputStream fileOutputStream = new FileOutputStream(targetFilePath.toString());
            branchWorkbook.write(fileOutputStream);
            fileOutputStream.flush();
            branchWorkbook.close();
            fileOutputStream.close();
            fileInputStream.close();
        } catch (Exception e) {
            e.printStackTrace();
        }
    }
}
