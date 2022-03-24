package com.tkachenko.buyerhelper.utils;

import com.tkachenko.buyerhelper.entity.FileEntity;
import com.tkachenko.buyerhelper.exceptions.IncorrectExtensionException;


import java.nio.file.Path;

public class FileUtils {

    public static void validateExcelExtension (String fileName) {
        String fileExtension = fileName.substring(fileName.lastIndexOf("."));
        if (!fileExtension.equals(".xls") && !fileExtension.equals(".xlsx"))
            throw new IncorrectExtensionException("Wrong file extension");
    }

    public static Path getEntityPath(Path fileStorageLocation, FileEntity fileEntity) {

        Path entityFileLocation = fileStorageLocation.resolve(fileEntity.getYear())
                .resolve(fileEntity.getMonth()).resolve(fileEntity.getDay()).
                resolve(fileEntity.getTime()).resolve(fileEntity.getStorageFileName());
        return entityFileLocation;
    }
}
