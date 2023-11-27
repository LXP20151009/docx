package org.example;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.zip.ZipEntry;
import java.util.zip.ZipInputStream;

public class DocxProcessor {
    public static void unzipDocx(String docxFilePath, String outputFolder) throws IOException {
        try (ZipInputStream zis = new ZipInputStream(new FileInputStream(docxFilePath))) {
            byte[] buffer = new byte[1024];
            ZipEntry zipEntry = zis.getNextEntry();
            while (zipEntry != null) {
                String fileName = zipEntry.getName();
                File newFile = new File(outputFolder + File.separator + fileName);
                new File(newFile.getParent()).mkdirs();
                try (FileOutputStream fos = new FileOutputStream(newFile)) {
                    int len;
                    while ((len = zis.read(buffer)) > 0) {
                        fos.write(buffer, 0, len);
                    }
                }
                zipEntry = zis.getNextEntry();
            }
        }
    }
    public static void main(String[] args) throws IOException {
        String sourcePath = "D:/batch_word/";//test_word
        String tempPath = "D:tempPath";
        String desPath = "D:/desPath";
        File sourceFile = new File(sourcePath);

        FileInputStream fileInputStream = new FileInputStream(sourceFile);
        if (sourceFile.isDirectory())
            for (File f : sourceFile.listFiles()) {
                if (!f.isDirectory() && f.getName().indexOf("docx") > 0) {
                    DocxProcessor.unzipDocx(f.getAbsolutePath(),
                            tempPath + f.getName().replaceAll("\\.docx", ""));
                }
            }
    }
}