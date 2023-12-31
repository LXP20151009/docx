package org.example;

import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFPicture;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.apache.xmlbeans.XmlException;
import org.apache.xmlbeans.XmlToken;
import org.openxmlformats.schemas.drawingml.x2006.main.CTNonVisualDrawingProps;
import org.openxmlformats.schemas.drawingml.x2006.main.CTPositiveSize2D;
import org.openxmlformats.schemas.drawingml.x2006.picture.CTPicture;
import org.openxmlformats.schemas.drawingml.x2006.wordprocessingDrawing.CTInline;

import java.io.IOException;
import java.io.InputStream;

public class CustomXWPFDocument extends XWPFDocument
{
	public CustomXWPFDocument()
    {
		super();
	}
	
	public CustomXWPFDocument(OPCPackage opcPackage) throws IOException
    {
		super(opcPackage);
	}
	
    public CustomXWPFDocument(InputStream in) throws IOException
    {
        super(in);
    }

    public static void createPicture(XWPFParagraph paragraph,
                              XWPFRun run,
                              XWPFPicture picture,
                              String blipId, int id, long width, long height, long w, String clr,String bid)
    {
        final int EMU = 9525;
       // width *= EMU;
        //height *= EMU;
        //String blipId = getAllPictures().get(id).getPackageRelationship().getId();
     //  run.getCTR().getDrawingList().remove(picture);
      // CTInline inline =run.getCTR().addNewDrawing().addNewInline();
//        XWPFParagraph para=  createParagraph();
//        XWPFRun paraRun = para.createRun();
//
//        CTR ctr= paraRun.getCTR();
       //CTInline inline= createParagraph().createRun().getCTR().addNewDrawing().addNewInline();
       // run.getCTR().getDrawingList().remove(picture);
        CTInline inline= run.getCTR().addNewDrawing().addNewInline();
        String picXml = "" +
                "<a:graphic xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\">" +
                "   <a:graphicData uri=\"http://schemas.openxmlformats.org/drawingml/2006/picture\">" +
                "      <pic:pic xmlns:pic=\"http://schemas.openxmlformats.org/drawingml/2006/picture\">" +
                "         <pic:nvPicPr>" +
                "            <pic:cNvPr id=\"" + id + "\" name=\"Generated\"/>" +
                "            <pic:cNvPicPr/>" +
                "         </pic:nvPicPr>" +
                "         <pic:blipFill>" +
                "            <a:blip r:embed=\"" + bid + "\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\"/>" +
                "            <a:stretch>" +
                "               <a:fillRect/>" +
                "            </a:stretch>" +
                "         </pic:blipFill>" +
                "         <pic:spPr>" +
                "            <a:xfrm>" +
                "               <a:off x=\"0\" y=\"0\"/>" +
                "               <a:ext cx=\"" + width + "\" cy=\"" + height + "\"/>" +
                "            </a:xfrm>" +
                "            <a:prstGeom prst=\"rect\">" +
                "               <a:avLst/>" +
                "            </a:prstGeom>" +
                "<a:ln w=\""+w+"\">"+
                "<a:solidFill>"+
                "<a:srgbClr val=\""+clr+"\"/>"+
                   "</a:solidFill>"+
                       "</a:ln>"+
                "         </pic:spPr>" +
                "      </pic:pic>" +
                "   </a:graphicData>" +
                "</a:graphic>";

        //CTGraphicalObjectData graphicData = inline.addNewGraphic().addNewGraphicData();
        XmlToken xmlToken = null;
        try
        {
            xmlToken = XmlToken.Factory.parse(picXml);
        } catch(XmlException xe) {
            xe.printStackTrace();
        }
        inline.set(xmlToken);
        //graphicData.set(xmlToken);

        inline.setDistT(0);
        inline.setDistB(0);
        inline.setDistL(0);
        inline.setDistR(0);

        CTPositiveSize2D extent = inline.addNewExtent();
        extent.setCx(width);
        extent.setCy(height);

        CTNonVisualDrawingProps docPr = inline.addNewDocPr();
        docPr.setId(id);
        docPr.setName("Picture " + id);
        docPr.setDescr("Generated");


    }

    public static XWPFPicture createAndReturnPicture(
                              XWPFRun run,
                              XWPFPicture picture,
                              String blipId, int id, long width, long height, long w, String clr,String bid) throws XmlException {
        final int EMU = 9525;


        //CTInline inline= createParagraph().createRun().getCTR().addNewDrawing().addNewInline();
        // run.getCTR().getDrawingList().remove(picture);
        //CTInline inline= run.getCTR().addNewDrawing().addNewInline();
        String picXml =
                               "         <pic:nvPicPr>" +
                "            <pic:cNvPr id=\"" + id + "\" name=\"Generated\"/>" +
                "            <pic:cNvPicPr/>" +
                "         </pic:nvPicPr>" +
                "         <pic:blipFill>" +
                "            <a:blip r:embed=\"" + bid + "\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\"/>" +
                "            <a:stretch>" +
                "               <a:fillRect/>" +
                "            </a:stretch>" +
                "         </pic:blipFill>" +
                "         <pic:spPr>" +
                "            <a:xfrm>" +
                "               <a:off x=\"0\" y=\"0\"/>" +
                "               <a:ext cx=\"" + width + "\" cy=\"" + height + "\"/>" +
                "            </a:xfrm>" +
                "            <a:prstGeom prst=\"rect\">" +
                "               <a:avLst/>" +
                "            </a:prstGeom>" +
                "<a:ln w=\"" + w + "\">" +
                "<a:solidFill>" +
                "<a:srgbClr val=\"" + clr + "\"/>" +
                "</a:solidFill>" +
                "</a:ln>" +
                "         </pic:spPr>" ;
       CTPicture ctPicture= CTPicture.Factory.parse(picXml);
       return   new XWPFPicture(ctPicture, run);
        //CTGraphicalObjectData graphicData = inline.addNewGraphic().addNewGraphicData();
//        XmlToken xmlToken = null;
//        try {
//            xmlToken = XmlToken.Factory.parse(picXml);
//
//            return xmlToken;
//        } catch (XmlException xe) {
//            xe.printStackTrace();
//        }
////        inline.set(xmlToken);
////        //graphicData.set(xmlToken);
////
////        inline.setDistT(0);
////        inline.setDistB(0);
////        inline.setDistL(0);
////        inline.setDistR(0);
////
////        CTPositiveSize2D extent = inline.addNewExtent();
////        extent.setCx(width);
////        extent.setCy(height);
////
////        CTNonVisualDrawingProps docPr = inline.addNewDocPr();
////        docPr.setId(id);
////        docPr.setName("Picture " + id);
////        docPr.setDescr("Generated");
//
//        return xmlToken;

}
}