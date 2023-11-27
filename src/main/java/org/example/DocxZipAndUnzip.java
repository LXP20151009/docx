package org.example;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.zip.ZipEntry;
import java.util.zip.ZipInputStream;

public class DocxZipAndUnzip
{
    public static void unzipDocx(String docxFilePath, String outputFolder) throws IOException
    {
        try (ZipInputStream zis = new ZipInputStream(new FileInputStream(docxFilePath)))
        {
            byte[] buffer = new byte[1024];
            ZipEntry zipEntry = zis.getNextEntry();
            while (zipEntry != null)
            {
                String fileName = zipEntry.getName();

                File newFile = new File(outputFolder + File.separator + fileName);
                new File(newFile.getParent()).mkdirs();
                try (FileOutputStream fos = new FileOutputStream(newFile))
                {
                    int len;

                    while ((len = zis.read(buffer)) > 0)
                    {
                        fos.write(buffer, 0, len);
                    }
                }
                zipEntry = zis.getNextEntry();
            }
        }
    }

    public static void zipDocx(String docxFilePath, String outputFolder) throws IOException
    {

        FileInputStream fileInputStream = new FileInputStream(docxFilePath);
        try (ZipInputStream zis = new ZipInputStream(new FileInputStream(docxFilePath)))
        {
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
    public static void writeDocxXml(String xmlContent, String docxFilePath) throws IOException
    {

        FileInputStream fileInputStream = new FileInputStream(docxFilePath);
        try (ZipInputStream zis = new ZipInputStream(new FileInputStream(docxFilePath)))
        {
            byte[] buffer = new byte[1024];
            ZipEntry zipEntry = zis.getNextEntry();
            while (zipEntry != null)
            {
                String fileName = zipEntry.getName();
                File newFile = new File(docxFilePath + File.separator + fileName);
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
    public static void mainAgo(String[] args) throws IOException
    {
        String sourcePath = "D:/batch_word/";//test_word
        String tempPath = "D:tempPath";
        String desPath = "D:/desPath";
        File sourceFile = new File(sourcePath);

        FileInputStream fileInputStream = new FileInputStream(sourceFile);
        if (sourceFile.isDirectory())
            for (File f : sourceFile.listFiles())
            {
                if (!f.isDirectory() && f.getName().indexOf("docx") > 0)
                {
                    DocxZipAndUnzip.unzipDocx(f.getAbsolutePath(),
                            tempPath + f.getName().replaceAll("\\.docx", ""));
                }
            }
    }
    public static void main(String[] args) throws Exception
    {
        String sourcePath = "D:/test_word/realFile.docx";//test_word
        String tempPath = "D:/tempPath";
        String outputFolder="D:/outputFolder";
        String desPath = "D:/desPath";
        byte []fileBuffer =new byte[100];
        byte[]docXmlBuffer = new byte[100];
        File sourceFile = new File(sourcePath);

        FileInputStream fileInputStream = new FileInputStream(sourceFile);

        System.out.println("文件长度："+sourcePath.length()+" Max"+Integer.MAX_VALUE);
        fileBuffer=new byte[(int)sourceFile.length()];
        fileInputStream.read(fileBuffer);
        String xmlContent = CreateXWPFDocumentDumpDocumentXML.doIt()[0];
        docXmlBuffer=xmlContent.getBytes();
      KMP kmp=new KMP();
     int pos= kmp.indexOf(fileBuffer.toString(),xmlContent);
     String changedXmlContent=CreateXWPFDocumentDumpDocumentXML.doIt()[1];;
     File desFile=new File(desPath+File.pathSeparator+sourceFile.getName());
     for(int i=pos;i<xmlContent.length();i++)
     {
         fileBuffer[i+pos]=changedXmlContent.getBytes()[i];
     }
     FileOutputStream fileOutputStream= new FileOutputStream(desFile);
     fileOutputStream.write(fileBuffer);
     fileOutputStream.flush();
     fileInputStream.close();
     fileOutputStream.close();
        //        //String.copyValueOf(fileBuffer,0,fileBuffer.length);
        //        //fileBuffer.toString();
        //
        //
        //        FileOutputStream fileOutputStream= new FileOutputStream
        //                (desPath+File.separator+sourceFile.getName());
        //        fileInputStream.read();
        //        //ZipOutputStream zipOutputStream = new ZipOutputStream();


    }
}