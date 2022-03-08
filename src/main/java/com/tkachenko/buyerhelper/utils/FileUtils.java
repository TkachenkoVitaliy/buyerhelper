package com.tkachenko.buyerhelper.utils;

import com.tkachenko.buyerhelper.exceptions.IncorrectExtensionException;

public class FileUtils {
    public static void validateExcelExtension (String fileName) {
        String fileExtension = fileName.substring(fileName.lastIndexOf("."));
        if (!fileExtension.equals(".xls") && !fileExtension.equals(".xlsx")) throw new IncorrectExtensionException("Wrong file extension");
    }
}
