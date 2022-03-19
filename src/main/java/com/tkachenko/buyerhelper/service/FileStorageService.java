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
    private final Path mmkAcceptPath;
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
    private final String mmkAcceptLibraryName;
    private final Path fileMmkAcceptLibraryPath;
    //TODO use xls or xlsx file extension
    private final String xls =".xls";
    private final String xlsx =".xlsx";


    @Autowired
    public FileStorageService (FileStorageProperties fileStorageProperties, FileDBService fileDBService,
                               ExcelService excelService, MmkService mmkService) {
        this.fileDBService = fileDBService;
        this.fileStorageLocation = Paths.get(fileStorageProperties.getUploadDir()).toAbsolutePath().normalize();
        this.mmkAcceptPath = fileStorageLocation.resolve(mmkAcceptParentDirectory);
        this.excelService = excelService;
        this.mmkService = mmkService;
        this.mmkAcceptLibraryName = excelService.getAcceptLibraryName();
        this.fileMmkAcceptLibraryPath = mmkAcceptPath.resolve(mmkAcceptLibraryName);
    }



    public void storeFiles (MultipartFile fileMmkOracle,
                              MultipartFile fileMmkDependencies,
                              MultipartFile fileOtherFactory) {

        GregorianCalendar currentDateAndTime = new GregorianCalendar();

        String originalFileMmkOracleName = fileMmkOracle.getOriginalFilename();
        String originalFileMmkDependenciesName = fileMmkDependencies.getOriginalFilename();
        String originalFileOtherFactoryName = fileOtherFactory.getOriginalFilename();

        FileUtils.validateExcelExtension(originalFileMmkOracleName);
        FileUtils.validateExcelExtension(originalFileMmkDependenciesName);
        FileUtils.validateExcelExtension(originalFileOtherFactoryName);

        String yearFolderName = DateUtils.getYear(currentDateAndTime);
        String monthFolderName = DateUtils.getMonth(currentDateAndTime);
        String dayFolderName = DateUtils.getDay(currentDateAndTime);
        String timeFolderName = DateUtils.getTime(currentDateAndTime);

        Path targetFolder = fileStorageLocation.resolve(yearFolderName).resolve(monthFolderName)
                .resolve(dayFolderName).resolve(timeFolderName);


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
                    yearFolderName, monthFolderName, dayFolderName, timeFolderName, true);

            Files.copy(fileMmkDependencies.getInputStream(), fileMmkDependenciesPath, StandardCopyOption.REPLACE_EXISTING);
            fileDBService.save(originalFileMmkDependenciesName, fileMmkDependenciesName,
                    yearFolderName, monthFolderName, dayFolderName, timeFolderName, true);

            Files.copy(fileOtherFactory.getInputStream(), fileOtherFactoryPath, StandardCopyOption.REPLACE_EXISTING);
            fileDBService.save(originalFileOtherFactoryName, fileOtherFactoryName,
                    yearFolderName, monthFolderName, dayFolderName, timeFolderName, true);

            Files.copy(fileOtherFactory.getInputStream(), fileSummaryPath, StandardCopyOption.REPLACE_EXISTING);
            fileDBService.save("CREATED", fileSummaryName,
                    yearFolderName, monthFolderName, dayFolderName, timeFolderName, true);
            excelService.refactorSummaryFile(fileSummaryPath);

            mmkService.parseMmkToOtherFactoryFormat(fileMmkOraclePath, fileMmkAcceptLibraryPath, fileMmkDependenciesPath);
        } catch (IOException e) {
            //TODO сделать собственное исключение в package Exceptions (FileStorageException)
            e.printStackTrace();
        }

    }

    public void storeFiles (MultipartFile mmkAccept) {

        String mmkAcceptOriginalName = mmkAccept.getOriginalFilename();
        FileUtils.validateExcelExtension(mmkAcceptOriginalName);

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
