package com.tkachenko.buyerhelper.controller;

import com.tkachenko.buyerhelper.service.FileDownloadService;
import com.tkachenko.buyerhelper.service.FileStorageService;
import org.springframework.core.io.Resource;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.http.HttpHeaders;
import org.springframework.http.MediaType;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.*;
import org.springframework.web.multipart.MultipartFile;

import javax.servlet.http.HttpServletRequest;
import java.io.IOException;

@RestController
public class FileStorageController {

    @Autowired
    private FileStorageService fileStorageService;
    @Autowired
    private FileDownloadService fileDownloadService;

    @PostMapping ("/uploadFiles")
    public String storeFile (@RequestParam("mmkOracle")MultipartFile fileMmkOracle,
                             @RequestParam("mmkDependencies") MultipartFile fileMmkDependencies,
                             @RequestParam("otherFactory") MultipartFile fileOtherFactory) {

        fileStorageService.storeFiles(fileMmkOracle, fileMmkDependencies, fileOtherFactory);
        return "UPLOAD COMPLETE";
    }

    @PostMapping ("/uploadAccept")
    public String storeAccept (@RequestParam("mmkAccept") MultipartFile mmkAccept) {

        fileStorageService.storeFiles(mmkAccept);
        return "ACCEPT UPLOADED";
    }



    @GetMapping("/downloadSummaryFile")
    public ResponseEntity<Resource> downloadSummaryFile(HttpServletRequest request) {
        Resource resource = fileDownloadService.loadSummaryFileAsResource();

        String contentType = null;
        try {
            contentType = request.getServletContext().getMimeType(resource.getFile().getAbsolutePath());
        } catch (IOException ex) {
            System.out.println("Could not determine file type");
        }

        if (contentType == null) {
            contentType = "application/octet-stream";
        }

        return ResponseEntity.ok().contentType(MediaType.parseMediaType(contentType))
                .header(HttpHeaders.CONTENT_DISPOSITION, "attachment; filename=SUMMARY.xlsx").body(resource);
    }

    @GetMapping("/downloadZipFile")
    public ResponseEntity<Resource> downloadZipFile(HttpServletRequest request) {
        Resource resource = fileDownloadService.loadBranchesZipFileAsResource();
        String fileName = resource.getFilename();
        System.out.println("FILE NAME - " + fileName);
        String headerValues = "attachment; filename=" + fileName; // + ".zip";

        String contentType = null;
        try {
            contentType = request.getServletContext().getMimeType(resource.getFile().getAbsolutePath());
        } catch (IOException ex) {
            System.out.println("Could not determine file type");
        }

        if (contentType == null) {
            contentType = "application/octet-stream";
        }

        return ResponseEntity.ok().contentType(MediaType.parseMediaType(contentType))
                .header(HttpHeaders.CONTENT_DISPOSITION, headerValues).body(resource);
    }
}
