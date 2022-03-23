package com.tkachenko.buyerhelper.controller;

import com.tkachenko.buyerhelper.service.FileStorageService;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.bind.annotation.RestController;
import org.springframework.web.multipart.MultipartFile;

@RestController
public class FileStorageController {

    @Autowired
    private FileStorageService fileStorageService;

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
    public String downloadSummaryFile() {
        return "TEST";
    }

    @GetMapping("/downloadZipFile")
    public String downloadZipFile() {
        return "TEST";
    }
}
