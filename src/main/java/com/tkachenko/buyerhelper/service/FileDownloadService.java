package com.tkachenko.buyerhelper.service;

import com.tkachenko.buyerhelper.entity.FileEntity;
import com.tkachenko.buyerhelper.property.FileStorageProperties;
import com.tkachenko.buyerhelper.service.excel.ExcelService;
import com.tkachenko.buyerhelper.service.mmk.MmkService;
import com.tkachenko.buyerhelper.utils.FileUtils;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.core.io.Resource;
import org.springframework.core.io.UrlResource;
import org.springframework.stereotype.Service;

import java.net.MalformedURLException;
import java.nio.file.Path;
import java.nio.file.Paths;

@Service
public class FileDownloadService {

    private final Path fileStorageLocation;
    private final FileDBService fileDBService;
    private final ExcelService excelService;
    private final MmkService mmkService;
    private final SummarySplitter summarySplitter;

    private final String SUMMARY_NAME = "SUMMARY.xlsx";

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

    public void loadBranchesZipFileAsResource() {
        summarySplitter.splitFiles();
    }
}
