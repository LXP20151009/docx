package org.example;

import java.io.*;
import java.util.zip.*;

public class ZipUnzipDocx {

    public static void main(String[] args) {
        String inputFilePath = "D:/test_word/new.docx";
        String outputFolderPath = "D:/test_word/output-folder";

        // Step 1: Unzip the DOCX file
        unzipDocx(inputFilePath, outputFolderPath);

        // Step 2: Modify the contents (e.g., styles.xml) in the output folder as needed

        // Step 3: Zip the modified contents back into a DOCX file
        zipDocx(outputFolderPath, "D:/test_word/reset/");
    }

    private static void unzipDocx(String inputFilePath, String outputFolderPath) {
        try (ZipInputStream zipInputStream = new ZipInputStream(new FileInputStream(inputFilePath))) {
            byte[] buffer = new byte[1024];
            ZipEntry zipEntry;
            File srcFile= new File(inputFilePath);
            // Create output directory if not exists
            File outputFolder = new File(outputFolderPath+File.separator+srcFile.getName());
            if (!outputFolder.exists()) {
                outputFolder.mkdirs();
            }

            // Extract each entry
            while ((zipEntry = zipInputStream.getNextEntry()) != null) {
                String entryName = zipEntry.getName();
                File entryFile = new File(outputFolderPath, entryName);

                // Create parent directories if not exists
                entryFile.getParentFile().mkdirs();

                // Write the entry to the file
                try (FileOutputStream fileOutputStream = new FileOutputStream(entryFile)) {
                    int length;
                    while ((length = zipInputStream.read(buffer)) > 0) {
                        fileOutputStream.write(buffer, 0, length);
                    }
                }

                zipInputStream.closeEntry();
            }

            System.out.println("Unzipping completed successfully.");
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    private static void zipDocx(String inputFolderPath, String outputFilePath) {
        try (FileOutputStream fileOutputStream = new FileOutputStream(outputFilePath);
             ZipOutputStream zipOutputStream = new ZipOutputStream(fileOutputStream)) {

            File inputFolder = new File(inputFolderPath);
            zipFiles(inputFolder, inputFolder, zipOutputStream);

            System.out.println("Zipping completed successfully.");
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    private static void zipFiles(File rootFolder, File currentFile, ZipOutputStream zipOutputStream) throws IOException {
        byte[] buffer = new byte[1024];
        int length;

        if (currentFile.isDirectory()) {
            String[] files = currentFile.list();

            if (files != null) {
                for (String file : files) {
                    File entryFile = new File(currentFile, file);
                    zipFiles(rootFolder, entryFile, zipOutputStream);
                }
            }
        } else {
            try (FileInputStream fileInputStream = new FileInputStream(currentFile)) {
                String entryName = currentFile.getAbsolutePath().substring(rootFolder.getAbsolutePath().length() + 1);
                ZipEntry zipEntry = new ZipEntry(entryName);

                zipOutputStream.putNextEntry(zipEntry);

                while ((length = fileInputStream.read(buffer)) > 0) {
                    zipOutputStream.write(buffer, 0, length);
                }

                zipOutputStream.closeEntry();
            }
        }
    }
}
