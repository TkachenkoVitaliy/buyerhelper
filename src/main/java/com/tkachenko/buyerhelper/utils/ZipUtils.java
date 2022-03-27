package com.tkachenko.buyerhelper.utils;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.nio.file.Path;
import java.util.ArrayList;
import java.util.GregorianCalendar;
import java.util.zip.ZipEntry;
import java.util.zip.ZipOutputStream;

public class ZipUtils {
    private final static String ZIP_EXTENSION = ".zip";

    public static Path zipListFiles (ArrayList<String> listFilesAddresses, Path zipDirectory) {
        GregorianCalendar time = new GregorianCalendar();
        String zippedFileName = (DateUtils.getTime(time) + " " + DateUtils.getDay(time) + " " + DateUtils.getMonth(time)
         + " " + DateUtils.getYear(time) + ZIP_EXTENSION);
        Path zippedFilePath = zipDirectory.resolve(zippedFileName);
        try {
            FileOutputStream fileOutputStream = new FileOutputStream(zippedFilePath.toString());
            ZipOutputStream zipOutputStream = new ZipOutputStream(fileOutputStream);
            for (String fileAddress : listFilesAddresses) {
                File fileToZip = new File(fileAddress);
                FileInputStream fileInputStream = new FileInputStream(fileToZip);
                ZipEntry zipEntry = new ZipEntry(fileToZip.getName());
                zipOutputStream.putNextEntry(zipEntry);

                byte[] bytes = new byte[1024];
                int length;
                while((length = fileInputStream.read(bytes)) >= 0) {
                    zipOutputStream.write(bytes, 0, length);
                }
                fileInputStream.close();
            }
            zipOutputStream.flush();
            zipOutputStream.close();
            fileOutputStream.flush();
            fileOutputStream.close();
        } catch (Exception e) {
            e.printStackTrace();
        }


        return zippedFilePath;
    }
}
