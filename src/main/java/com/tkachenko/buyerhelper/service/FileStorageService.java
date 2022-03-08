package com.tkachenko.buyerhelper.service;

import com.tkachenko.buyerhelper.property.FileStorageProperties;
import com.tkachenko.buyerhelper.service.excel.ExcelService;
import com.tkachenko.buyerhelper.service.mmk.MmkService;
import com.tkachenko.buyerhelper.utils.FileUtils;
import com.tkachenko.buyerhelper.utils.DateUtils;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Service;
import org.springframework.web.multipart.MultipartFile;

import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.nio.file.StandardCopyOption;
import java.util.GregorianCalendar;

@Service
public class FileStorageService {

    private final Path fileStorageLocation;
    private final String fileMmkOracleName = "MmkOracle.xlsx";
    private final String fileMmkDependenciesName = "MmkDependencies.xlsx";
    private final String fileOtherFactoryName = "OtherFactories.xlsx";
    private final String fileSummaryName = "SUMMARY.xlsx";
    private final FileDBService fileDBService;
    private final ExcelService excelService;
    private final MmkService mmkService;
    private final String mmkAcceptParentDirectory = "mmkAccept";
    private final String mmkAcceptName = "mmkAccept.xlsx";
    private final String mmkAcceptRefactoredName = "mmkAcceptRefactored.xlsx";
    //TODO use xls or xlsx file extension
    private final String xls =".xls";
    private final String xlsx =".xlsx";


    @Autowired
    public FileStorageService (FileStorageProperties fileStorageProperties, FileDBService fileDBService,
                               ExcelService excelService, MmkService mmkService) {
        this.fileDBService = fileDBService;
        this.fileStorageLocation = Paths.get(fileStorageProperties.getUploadDir()).toAbsolutePath().normalize();
        this.excelService = excelService;
        this.mmkService = mmkService;
    }

    public void storeFiles (MultipartFile fileMmkOracle,
                              MultipartFile fileMmkDependencies,
                              MultipartFile fileOtherFactory) {

        GregorianCalendar gregorianCalendar = new GregorianCalendar();

        String originalFileMmkOracleName = fileMmkOracle.getOriginalFilename();
        String originalFileMmkDependenciesName = fileMmkDependencies.getOriginalFilename();
        String originalFileOtherFactoryName = fileOtherFactory.getOriginalFilename();

        FileUtils.validateExcelExtension(originalFileMmkOracleName);
        FileUtils.validateExcelExtension(originalFileMmkDependenciesName);
        FileUtils.validateExcelExtension(originalFileOtherFactoryName);

        String yearFolder = DateUtils.getYear(gregorianCalendar);
        String monthFolder = DateUtils.getMonth(gregorianCalendar);
        String dayFolder = DateUtils.getDay(gregorianCalendar);
        String timeFolder = DateUtils.getTime(gregorianCalendar);

        Path targetFolder = fileStorageLocation.resolve(yearFolder).resolve(monthFolder)
                .resolve(dayFolder).resolve(timeFolder);


        try {
            Files.createDirectories(targetFolder);
        } catch (IOException e) {
            //TODO сделать собственное исключение в package Exceptions (FileStorageException)
            e.printStackTrace();
        }

        Path fileMmkOraclePath = targetFolder.resolve(fileMmkOracleName);
        Path fileMmkDependenciesPath = targetFolder.resolve(fileMmkDependenciesName);
        Path fileOtherFactoryPath = targetFolder.resolve(fileOtherFactoryName);
        Path fileSummaryPath = targetFolder.resolve(fileSummaryName);

        try {
            Files.copy(fileMmkOracle.getInputStream(), fileMmkOraclePath, StandardCopyOption.REPLACE_EXISTING);
            fileDBService.save(originalFileMmkOracleName, fileMmkOracleName,
                    yearFolder, monthFolder, dayFolder, timeFolder, true);

            Files.copy(fileMmkDependencies.getInputStream(), fileMmkDependenciesPath, StandardCopyOption.REPLACE_EXISTING);
            fileDBService.save(originalFileMmkDependenciesName, fileMmkDependenciesName,
                    yearFolder, monthFolder, dayFolder, timeFolder, true);

            Files.copy(fileOtherFactory.getInputStream(), fileOtherFactoryPath, StandardCopyOption.REPLACE_EXISTING);
            fileDBService.save(originalFileOtherFactoryName, fileOtherFactoryName,
                    yearFolder, monthFolder, dayFolder, timeFolder, true);

            Files.copy(fileOtherFactory.getInputStream(), fileSummaryPath, StandardCopyOption.REPLACE_EXISTING);
            fileDBService.save("CREATED", fileSummaryName,
                    yearFolder, monthFolder, dayFolder, timeFolder, true);
            excelService.refactorSummaryFile(fileSummaryPath);

            mmkService.parseMmkToOtherFactoryFormat(fileMmkOraclePath);

        } catch (IOException e) {
            //TODO сделать собственное исключение в package Exceptions (FileStorageException)
            e.printStackTrace();
        }

    }

    public void storeFiles (MultipartFile mmkAccept) {

        String mmkAcceptOriginalName = mmkAccept.getOriginalFilename();
        FileUtils.validateExcelExtension(mmkAcceptOriginalName);
        Path mmkAcceptPath = fileStorageLocation.resolve(mmkAcceptParentDirectory);

        try {
            Files.createDirectories(mmkAcceptPath);
        } catch (IOException e) {
            e.printStackTrace();
        }

        Path fileMmkAcceptPath = mmkAcceptPath.resolve(mmkAcceptName);
        Path fileMmkAcceptRefactoredPath = mmkAcceptPath.resolve(mmkAcceptRefactoredName);

        try {
            Files.copy(mmkAccept.getInputStream(), fileMmkAcceptPath, StandardCopyOption.REPLACE_EXISTING);
        } catch (IOException e) {
            e.printStackTrace();
        }

        excelService.parseMmkAccept(fileMmkAcceptPath, fileMmkAcceptRefactoredPath);

        excelService.addToAcceptLibrary(fileMmkAcceptRefactoredPath);
    }

}
