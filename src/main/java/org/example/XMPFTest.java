package org.example;

import com.deepoove.poi.XWPFTemplate;
import com.deepoove.poi.data.*;
import org.apache.poi.hwpf.HWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFDocument;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

public class XMPFTest {

    public static void main(String[] args) throws IOException {
        String srcFile="D:/test_word/pureWord.docx";
        String desFile="D:/test_word/modified_document.docx";
        FileInputStream fis = new FileInputStream(srcFile);
        FileOutputStream fos = new FileOutputStream(desFile);
        XWPFTemplate xwpfTemplate= XWPFTemplate.compile(srcFile);
        //List<XWPFPicture> pictureList=
        XWPFDocument document=        xwpfTemplate.getXWPFDocument();
        String text = "this a paragraph";
        Texts.of().style()
        Documents.of().addParagraph(Paragraphs.of().addPicture(Pictures.of().sizeInCm(14.3,8).center().altMeta().create()));
        DocumentRenderData data = Documents.of().addParagraph(Paragraphs.of(text).create()).create();
        XWPFTemplate template = XWPFTemplate.create(data);

    }
}
