package com.tkachenko.buyerhelper.service;

import com.tkachenko.buyerhelper.entity.FileEntity;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Service;
import com.tkachenko.buyerhelper.repository.FileEntityRepository;

@Service
public class FileDBService {

    private final FileEntityRepository fileEntityRepository;

    @Autowired
    public FileDBService(FileEntityRepository fileEntityRepository) {
        this.fileEntityRepository = fileEntityRepository;
    }

    public void save(String originalFileName, String storageFileName,
                     String year, String month, String day, String time, Boolean isActual) {

        FileEntity fileEntity = new FileEntity(originalFileName, storageFileName, year, month, day, time, isActual);
        fileEntityRepository.save(fileEntity);
    }

    public void setPreviousIsActualFalse(String storageFileName) {
        FileEntity previousFile = fileEntityRepository.findByStorageFileNameAndIsActual(storageFileName, true);
        if(previousFile != null) {
            previousFile.setActual(false);
            fileEntityRepository.save(previousFile);
        }
    }

    public FileEntity getActualFileByStorageName (String storageFileName) {
        FileEntity actualFile = fileEntityRepository.findByStorageFileNameAndIsActual(storageFileName, true);
        return actualFile;
    }
}
