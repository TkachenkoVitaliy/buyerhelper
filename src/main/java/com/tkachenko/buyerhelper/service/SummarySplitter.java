package com.tkachenko.buyerhelper.service;

import com.tkachenko.buyerhelper.entity.FileEntity;
import com.tkachenko.buyerhelper.property.FileStorageProperties;
import com.tkachenko.buyerhelper.utils.ExcelUtils;
import com.tkachenko.buyerhelper.utils.FileUtils;
import org.apache.poi.openxml4j.util.ZipSecureFile;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
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
    private final static String DIRECTORY_NAME = "forZip";
    private final static String[] citiesArray = {KRASNODAR, ROSTOV, MOSCOW, NOVGOROD, KAZAN, NAB_CHELNY, IZHEVSK, PERM,
    UFA, CHELYABINSK, YEKATERINBURG, TUMEN, OMSK, NOVOSIBIRSK, NOVOKUZNETSK};

    public void splitFiles() {
        for (String cityName : citiesArray) {
            Path targetFilePath = copySummaryFile(cityName);
            deleteUnnecessaryBranches(targetFilePath, cityName);
        }
    }

    private Path copySummaryFile(String cityName) {
        FileEntity summaryFileEntity = fileDBService.getActualFileByStorageName(SUMMARY);
        Path summaryFilePath = FileUtils.getEntityPath(fileStorageLocation, summaryFileEntity);
        Path targetDirectoryPath = fileStorageLocation.resolve(DIRECTORY_NAME);
        Path targetFilePath = targetDirectoryPath.resolve(cityName + EXTENSION);

        try {
            Files.createDirectories(targetDirectoryPath);
            Files.copy(summaryFilePath,targetFilePath, StandardCopyOption.REPLACE_EXISTING);
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

                for(int i = firstRowIndex; i <= lastRowIndex ; i++) {
                    XSSFRow currentRow = (XSSFRow) monthSheet.getRow(i);
                    if(currentRow != null) {
                        XSSFCell branchCell = currentRow.getCell(branchColIndex);
                        if(branchCell != null && branchCell.getCellType() != CellType.BLANK
                                && !branchCell.getStringCellValue().equals(cityName)) {
                            monthSheet.removeRow(currentRow);
                        }
                    }
                }

                int countRowIndex = monthSheet.getFirstRowNum() + 1;
                List<Integer> rowIndexesList = new ArrayList<>();

                for(Row currentRow : monthSheet) {
                    rowIndexesList.add(currentRow.getRowNum());
                }

                for(int j = 0; j < rowIndexesList.size(); j++) {
                    int currentListIndex = j;
                    int rowIndexFromList = rowIndexesList.get(j);
                    if(rowIndexFromList > currentListIndex) monthSheet.shiftRows(rowIndexFromList, rowIndexFromList,
                            currentListIndex-rowIndexFromList);
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
