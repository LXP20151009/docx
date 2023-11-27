package org.example;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.StringReader;

import com.sun.org.apache.xml.internal.resolver.Catalog;
import org.apache.poi.xwpf.usermodel.*;
import org.apache.xmlbeans.XmlOptions;

import javax.xml.bind.JAXBContext;
import javax.xml.bind.Unmarshaller;

public class CreateXWPFDocumentDumpDocumentXML {
    
 static String printDocumentXML(XWPFDocument docx) throws Exception {
     
  String xml;
  
//  System.out.println("Contents of org.openxmlformats.schemas.wordprocessingml.x2006.main.CTDocument1:");
//  org.apache.xmlbeans.XmlObject documentXmlObject = docx.getDocument();
//  xml = documentXmlObject.toString();
//  System.out.println(xml);
  
  System.out.println("Contents of whole DocumentDocument:");
  org.openxmlformats.schemas.wordprocessingml.x2006.main.CTDocument1 ctDocument1 = docx.getDocument();
  org.openxmlformats.schemas.wordprocessingml.x2006.main.DocumentDocument documentDocument =
          org.openxmlformats.schemas.wordprocessingml.x2006.main.DocumentDocument.Factory.newInstance(new XmlOptions());
 // documentDocument=
  documentDocument.setDocument(ctDocument1);
  xml = documentDocument.toString();

  //System.out.println(xml);
  return  xml;
 }

 public static void main(String[] args) throws Exception {

 // String fisFile="D:/test_word/realTestGenerate.docx";
  String fisFile="D:/test_word/realFile.docx";
  FileInputStream fis=new FileInputStream(fisFile);
  XWPFDocument docx = new XWPFDocument(fis);
//  XWPFParagraph paragraph = docx.createParagraph();
//  XWPFRun run=paragraph.createRun();
//  run.setBold(true);
//  run.setFontSize(22);
//  run.setText("The paragraph content ...");
//  paragraph = docx.createParagraph();

        String xmlContent=  printDocumentXML(docx);
//        xmlContent=   xmlContent.replaceAll("(a:ext\\ cx=\"\\d+)","a:ext\\ cx=\"5148000");
//        xmlContent=      xmlContent.replaceAll("(5148000\"\\ cy=\"\\d+)","5148000\"\\ cy=\"2880000");

     xmlContent=   xmlContent.replaceAll("(cx=\"\\d+)","cx=\"5148000");
     xmlContent=      xmlContent.replaceAll("(5148000\"\\ cy=\"\\d+)","5148000\"\\ cy=\"2880000");
        String[]strings= xmlContent.split("(</a:prstGeom>)|(</pic:spPr>)");
        StringBuilder stringBuilder=new StringBuilder();
        for(int i=0;i<strings.length;i++)
        {
          //System.out.println("strings"+i+" "+strings[i]);
          if(i%2==1)
          {
              strings[i]="</a:prstGeom>\n"+
                      "<a:ln w=\"9750\">\n" +
                      "                      <a:solidFill>\n" +
                      "<a:srgbClr val=\""+"33A5FF"+"\"/>\n"+
                     // "                        <a:schemeClr val=\"accent1\">\n" +
                     // "                          <a:lumMod val=\"60000\"/>\n" +
                     // "                          <a:lumOff val=\"40000\"/>\n" +
                     // "                        </a:schemeClr>\n" +
                      "                      </a:solidFill>\n" +
                      "                    </a:ln>\n"
                      +"</pic:spPr>\n";

  }
    stringBuilder.append(strings[i]);

   // docx.getDocument();
  //docx.getDocument()
  //if(i>10) break;
}
//     JAXBContext jc = JAXBContext.newInstance(org.openxmlformats.schemas.wordprocessingml.x2006.main.DocumentDocument.class);
//     Unmarshaller unmarshaller = jc.createUnmarshaller();
//     StringReader reader = new StringReader(stringBuilder.toString());
//     org.openxmlformats.schemas.wordprocessingml.x2006.main.DocumentDocument
//             documentDocument =
//             (org.openxmlformats.schemas.wordprocessingml.x2006.main.DocumentDocument)
//                     unmarshaller.unmarshal(reader);
     System.out.println("----");
     System.out.println(stringBuilder.toString());
//  try (FileOutputStream out = new FileOutputStream("./XWPFDocument.docx")) {
//    docx.write(out);
//  }

 }
}