package com.tkachenko.buyerhelper.service;

import com.tkachenko.buyerhelper.entity.FileEntity;
import com.tkachenko.buyerhelper.property.FileStorageProperties;
import com.tkachenko.buyerhelper.service.excel.ExcelService;
import com.tkachenko.buyerhelper.service.mmk.MmkService;
import com.tkachenko.buyerhelper.utils.FileUtils;
import com.tkachenko.buyerhelper.utils.ZipUtils;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.core.io.Resource;
import org.springframework.core.io.UrlResource;
import org.springframework.stereotype.Service;

import java.io.File;
import java.io.IOException;
import java.net.MalformedURLException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.ArrayList;

@Service
public class FileDownloadService {

    private final Path fileStorageLocation;
    private final FileDBService fileDBService;
    private final ExcelService excelService;
    private final MmkService mmkService;
    private final SummarySplitter summarySplitter;

    private final String SUMMARY_NAME = "SUMMARY.xlsx";
    private final String ZIP_DIRECTORY = "forZip";
    private final String ZIP_EXTENSION = ".zip";

    @Autowired
    public FileDownloadService(FileStorageProperties fileStorageProperties, FileDBService fileDBService,
                               ExcelService excelService, MmkService mmkService, SummarySplitter summarySplitter) {
        this.fileStorageLocation = Paths.get(fileStorageProperties.getUploadDir()).toAbsolutePath().normalize();
        this.fileDBService = fileDBService;
        this.excelService = excelService;
        this.mmkService = mmkService;
        this.summarySplitter = summarySplitter;
    }

    public Resource loadSummaryFileAsResource() {
        try{
            FileEntity actualSummaryEntity = fileDBService.getActualFileByStorageName(SUMMARY_NAME);
            Path actualSummaryFilePath = FileUtils.getEntityPath(fileStorageLocation, actualSummaryEntity);
            Resource resource = new UrlResource(actualSummaryFilePath.toUri());
            if (resource.exists()) {
                return resource;
            } else {
                throw new RuntimeException("File not found " + actualSummaryFilePath.toString());
            }

        } catch (MalformedURLException ex) {
            throw new RuntimeException("File not found");
        }
    }

    public Resource loadBranchesZipFileAsResource() {
        Path zipDirectory = fileStorageLocation.resolve(ZIP_DIRECTORY);
        ArrayList<String> branchFilesAddresses = summarySplitter.splitFiles();

        try{

            //remove old .zip files
            File file = new File(zipDirectory.toString());
            File[] files = file.listFiles();
            for (File currentFile : files) {
                String currentFileName = currentFile.getName();
                if (currentFileName.contains(ZIP_EXTENSION)) {
                    Files.deleteIfExists(currentFile.toPath());
                }
            }


            Path zippedFilePath = ZipUtils.zipListFiles(branchFilesAddresses, zipDirectory);
            Resource resource = new UrlResource(zippedFilePath.toUri());
            if(resource.exists()) {
                return resource;
            } else {
                throw new RuntimeException("File not found " + zippedFilePath.toString());
            }
        } catch (IOException ex) {
            throw new RuntimeException("File not found");
        }
    }
}
