package org.example;

import com.deepoove.poi.XWPFTemplate;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.util.Units;
import org.apache.poi.xwpf.usermodel.*;
import org.apache.xmlbeans.XmlException;
import org.apache.xmlbeans.XmlToken;
import org.openxmlformats.schemas.drawingml.x2006.wordprocessingDrawing.CTAnchor;
import org.openxmlformats.schemas.drawingml.x2006.wordprocessingDrawing.CTInline;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTDrawing;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTString;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTStyle;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTStyles;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

public class AModifyWordImage {

    public static String getTitleLvl(XWPFDocument doc, XWPFParagraph para) {
        String titleLvl = "";
        try {
            //判断该段落是否设置了大纲级别
            if (para.getCTP().getPPr().getOutlineLvl() != null) {
                return String.valueOf(para.getCTP().getPPr().getOutlineLvl().getVal());
            }
        } catch (Exception e) {
        }
        try {
            //判断该段落的样式是否设置了大纲级别
            if (doc.getStyles().getStyle(para.getStyle()).getCTStyle().getPPr().getOutlineLvl() != null) {
                return String.valueOf(doc.getStyles().getStyle(para.getStyle()).getCTStyle().getPPr().getOutlineLvl().getVal());
            }
        } catch (Exception e) {
        }
        try {
            //判断该段落的样式的基础样式是否设置了大纲级别
            if (doc.getStyles().getStyle(doc.getStyles().getStyle(para.getStyle()).getCTStyle().getBasedOn().getVal())
                    .getCTStyle().getPPr().getOutlineLvl() != null) {
                String styleName = doc.getStyles().getStyle(para.getStyle()).getCTStyle().getBasedOn().getVal();
                return String.valueOf(doc.getStyles().getStyle(styleName).getCTStyle().getPPr().getOutlineLvl().getVal());
            }
        } catch (Exception e) {

        }
//        try {
//            if (para.getStyleID() != null) {
//                return para.getStyleID();
//            }
//        } catch (Exception e) {
//
//        }

        return titleLvl;
    }


    public static void setAnchorToInline(XWPFRun run, XWPFPicture picture, long width, long EMUHeight) throws XmlException {
        List<org.openxmlformats.schemas.wordprocessingml.x2006.main.CTDrawing> drawingList
                = run.getCTR().getDrawingList();
        for (org.openxmlformats.schemas.wordprocessingml.x2006.main.CTDrawing drawing : drawingList) {
            for (CTAnchor ctAnchor : drawing.getAnchorList()) {


                if ((ctAnchor.getGraphic().getGraphicData()).toString().
                        indexOf("blip r:embed=\"" + picture.getCTPicture().getBlipFill().getBlip().getEmbed() + "\"") > -1) {

                    ctAnchor.getEffectExtent().setB("10000");
                    ctAnchor.getEffectExtent().setT("10000");
                    ctAnchor.getEffectExtent().setL("10000");
                    ctAnchor.getEffectExtent().setR("10000");
                    String inlineStr = "<wp:inline distT=\"0\" distB=\"0\" distL=\"0\" distR=\"0\" wp14:anchorId=\"2764B781\" wp14:editId=\"3860602A\" xmlns:wpc=\"http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas\" xmlns:cx=\"http://schemas.microsoft.com/office/drawing/2014/chartex\" xmlns:cx1=\"http://schemas.microsoft.com/office/drawing/2015/9/8/chartex\" xmlns:cx2=\"http://schemas.microsoft.com/office/drawing/2015/10/21/chartex\" xmlns:cx3=\"http://schemas.microsoft.com/office/drawing/2016/5/9/chartex\" xmlns:cx4=\"http://schemas.microsoft.com/office/drawing/2016/5/10/chartex\" xmlns:cx5=\"http://schemas.microsoft.com/office/drawing/2016/5/11/chartex\" xmlns:cx6=\"http://schemas.microsoft.com/office/drawing/2016/5/12/chartex\" xmlns:cx7=\"http://schemas.microsoft.com/office/drawing/2016/5/13/chartex\" xmlns:cx8=\"http://schemas.microsoft.com/office/drawing/2016/5/14/chartex\" xmlns:mc=\"http://schemas.openxmlformats.org/markup-compatibility/2006\" xmlns:aink=\"http://schemas.microsoft.com/office/drawing/2016/ink\" xmlns:am3d=\"http://schemas.microsoft.com/office/drawing/2017/model3d\" xmlns:o=\"urn:schemas-microsoft-com:office:office\" xmlns:oel=\"http://schemas.microsoft.com/office/2019/extlst\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" xmlns:m=\"http://schemas.openxmlformats.org/officeDocument/2006/math\" xmlns:v=\"urn:schemas-microsoft-com:vml\" xmlns:wp14=\"http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing\" xmlns:wp=\"http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing\" xmlns:w10=\"urn:schemas-microsoft-com:office:word\" xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\" xmlns:w14=\"http://schemas.microsoft.com/office/word/2010/wordml\" xmlns:w15=\"http://schemas.microsoft.com/office/word/2012/wordml\" xmlns:w16cex=\"http://schemas.microsoft.com/office/word/2018/wordml/cex\" xmlns:w16cid=\"http://schemas.microsoft.com/office/word/2016/wordml/cid\" xmlns:w16=\"http://schemas.microsoft.com/office/word/2018/wordml\" xmlns:w16sdtdh=\"http://schemas.microsoft.com/office/word/2020/wordml/sdtdatahash\" xmlns:w16se=\"http://schemas.microsoft.com/office/word/2015/wordml/symex\" xmlns:wpg=\"http://schemas.microsoft.com/office/word/2010/wordprocessingGroup\" xmlns:wpi=\"http://schemas.microsoft.com/office/word/2010/wordprocessingInk\" xmlns:wne=\"http://schemas.microsoft.com/office/word/2006/wordml\" xmlns:wps=\"http://schemas.microsoft.com/office/word/2010/wordprocessingShape\">\n" +
                            "  <wp:extent cx=\"3779520\" cy=\"2353945\"/>\n" +
                            "  <wp:effectExtent l=\"0\" t=\"0\" r=\"0\" b=\"0\"/>\n" +
                            "  <wp:docPr id=\"1\" name=\"图片 187\" descr=\"2faa35876327dbe6d1a7ef53c72d8b4\"/>\n" +
                            "  <wp:cNvGraphicFramePr>\n" +
                            "    <a:graphicFrameLocks noChangeAspect=\"1\" xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\"/>\n" +
                            "  </wp:cNvGraphicFramePr>\n" +
                            "  <a:graphic xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\">\n" +
                            "    <a:graphicData uri=\"http://schemas.openxmlformats.org/drawingml/2006/picture\">\n" +
                            "      <pic:pic xmlns:pic=\"http://schemas.openxmlformats.org/drawingml/2006/picture\">\n" +
                            "        <pic:nvPicPr>\n" +
                            "          <pic:cNvPr id=\"0\" name=\"图片 187\" descr=\"2faa35876327dbe6d1a7ef53c72d8b4\"/>\n" +
                            "          <pic:cNvPicPr>\n" +
                            "            <a:picLocks noChangeAspect=\"1\" noChangeArrowheads=\"1\"/>\n" +
                            "          </pic:cNvPicPr>\n" +
                            "        </pic:nvPicPr>\n" +
                            "        <pic:blipFill>\n" +
                            "          <a:blip r:embed=\"rId7\" cstate=\"print\">\n" +
                            "            <a:extLst>\n" +
                            "              <a:ext uri=\"{28A0092B-C50C-407E-A947-70E740481C1C}\">\n" +
                            "                <a14:useLocalDpi val=\"0\" xmlns:a14=\"http://schemas.microsoft.com/office/drawing/2010/main\"/>\n" +
                            "              </a:ext>\n" +
                            "            </a:extLst>\n" +
                            "          </a:blip>\n" +
                            "          <a:srcRect l=\"18706\" t=\"40797\" r=\"22948\" b=\"10738\"/>\n" +
                            "          <a:stretch>\n" +
                            "            <a:fillRect/>\n" +
                            "          </a:stretch>\n" +
                            "        </pic:blipFill>\n" +
                            "        <pic:spPr bwMode=\"auto\">\n" +
                            "          <a:xfrm>\n" +
                            "            <a:off x=\"0\" y=\"0\"/>\n" +
                            "            <a:ext cx=\"5148000\" cy=\"2353945\"/>\n" +
                            "          </a:xfrm>\n" +
                            "          <a:prstGeom prst=\"rect\">\n" +
                            "            <a:avLst/>\n" +
                            "          </a:prstGeom>\n" +
                            "          <a:noFill/>\n" +
                            "          <a:ln w=\"9500\">\n" +
                            "            <a:solidFill>\n" +
                            "              <a:srgbClr val=\"33A5FF\"/>\n" +
                            "            </a:solidFill>\n" +
                            "          </a:ln>\n" +
                            "        </pic:spPr>\n" +
                            "      </pic:pic>\n" +
                            "    </a:graphicData>\n" +
                            "  </a:graphic>\n" +
                            "</wp:inline>";
                    org.openxmlformats.schemas.wordprocessingml.x2006.main.
                            CTDrawing fakeDrawing =//(CTDrawing)XmlToken.Factory.parse(inlineStr);
                            (CTDrawing) CTDrawing.Factory.parse(inlineStr);
//                            (CTDrawing) drawingList.get(drawingList.indexOf(drawing))
//                            .set(XmlToken.Factory.parse(inlineStr));
                    CTInline ctInline = fakeDrawing.getInlineList().get(0);
                    ctInline.getExtent().set(ctAnchor.getExtent());
                    ctInline.getDocPr().set(ctAnchor.getDocPr());
                    ctInline.getCNvGraphicFramePr().set(ctAnchor.getCNvGraphicFramePr());
                    ctInline.getGraphic().set(ctAnchor.getGraphic());
                    //drawing.getInlineList().add(ctInline);
                    CTInline[] ctInline1 = new CTInline[drawing.getInlineArray().length + 1];
                    for (int i = 1; i < ctInline1.length; i++) {
                        ctInline1[i] = drawing.getInlineArray()[i - 1];
                    }
                    ctInline1[0] = ctInline;
                    drawing.setInlineArray(ctInline1);

                    // ctAnchor.getExtent().setCx(0);
                    // ctAnchor.getExtent().setCy(0);
                    // drawing.getAnchorList().remove(ctAnchor);
                    // drawing.set(fakeDrawing);
                    System.out.println("替换好的 drawing为：  " + drawing.toString());
                    return;
//                    XmlCursor cursor = ctAnchor.newCursor();

//                    cursor.selectPath("./*");
//                    while (cursor.toNextSelection())
//                    {
//                        XmlObject xmlObject = cursor.getObject();
//                        if (xmlObject instanceof CTAnchor)
//                        {
//
//                        }
//                    }

                    //ctAnchor.getGraphic().toString();
                }
//                if(((ctAnchor.getGraphic()) instanceof XWPFPicture)&&(ctAnchor.getGraphic().equals(picture)))
//                {
//                    ctAnchor.getExtent().setCx(width);
//                }
            }

        }

    }

    public static void PoiSetWidth(XWPFPicture picture, int pic, long width, long EMUHeight) throws XmlException {
        System.out.println("pic" + pic + "  picture.getCTPicture()  :" + picture.getCTPicture());
        System.out.println("pic" + pic + "  picture.getDepth()  :" + picture.getDepth());
        System.out.println("pic" + pic + "  picture.getWidth()  :" + picture.getWidth());
        System.out.println("pic" + pic + "  picture.getPictureData()  :" + picture.getPictureData());
        System.out.println("pic" + pic + "  picture.getDescription()  :" + picture.getDescription());


        if ((null == picture.getCTPicture().getSpPr().getLn())) {
            picture.getCTPicture().getSpPr().addNewLn().setW(9500);
        } else {
            picture.getCTPicture().getSpPr().getLn().setW(9500);
        }

        if ((null != picture.getCTPicture().getSpPr().getLn().getSolidFill())) {
            //picture.getCTPicture().getSpPr().getLn().
            String solidFillStr =
                    "<a:srgbClr val=\"33A5FF\" xmlns:pic=\"http://schemas.openxmlformats.org/drawingml/2006/picture\" xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" xmlns:wpc=\"http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas\" xmlns:cx=\"http://schemas.microsoft.com/office/drawing/2014/chartex\" xmlns:cx1=\"http://schemas.microsoft.com/office/drawing/2015/9/8/chartex\" xmlns:cx2=\"http://schemas.microsoft.com/office/drawing/2015/10/21/chartex\" xmlns:cx3=\"http://schemas.microsoft.com/office/drawing/2016/5/9/chartex\" xmlns:cx4=\"http://schemas.microsoft.com/office/drawing/2016/5/10/chartex\" xmlns:cx5=\"http://schemas.microsoft.com/office/drawing/2016/5/11/chartex\" xmlns:cx6=\"http://schemas.microsoft.com/office/drawing/2016/5/12/chartex\" xmlns:cx7=\"http://schemas.microsoft.com/office/drawing/2016/5/13/chartex\" xmlns:cx8=\"http://schemas.microsoft.com/office/drawing/2016/5/14/chartex\" xmlns:mc=\"http://schemas.openxmlformats.org/markup-compatibility/2006\" xmlns:aink=\"http://schemas.microsoft.com/office/drawing/2016/ink\" xmlns:am3d=\"http://schemas.microsoft.com/office/drawing/2017/model3d\" xmlns:o=\"urn:schemas-microsoft-com:office:office\" xmlns:oel=\"http://schemas.microsoft.com/office/2019/extlst\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" xmlns:m=\"http://schemas.openxmlformats.org/officeDocument/2006/math\" xmlns:v=\"urn:schemas-microsoft-com:vml\" xmlns:wp14=\"http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing\" xmlns:wp=\"http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing\" xmlns:w10=\"urn:schemas-microsoft-com:office:word\" xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\" xmlns:w14=\"http://schemas.microsoft.com/office/word/2010/wordml\" xmlns:w15=\"http://schemas.microsoft.com/office/word/2012/wordml\" xmlns:w16cex=\"http://schemas.microsoft.com/office/word/2018/wordml/cex\" xmlns:w16cid=\"http://schemas.microsoft.com/office/word/2016/wordml/cid\" xmlns:w16=\"http://schemas.microsoft.com/office/word/2018/wordml\" xmlns:w16sdtdh=\"http://schemas.microsoft.com/office/word/2020/wordml/sdtdatahash\" xmlns:w16se=\"http://schemas.microsoft.com/office/word/2015/wordml/symex\" xmlns:wpg=\"http://schemas.microsoft.com/office/word/2010/wordprocessingGroup\" xmlns:wpi=\"http://schemas.microsoft.com/office/word/2010/wordprocessingInk\" xmlns:wne=\"http://schemas.microsoft.com/office/word/2006/wordml\" xmlns:wps=\"http://schemas.microsoft.com/office/word/2010/wordprocessingShape\"/>"

                    //+
                    // "                        <a:schemeClr val=\"accent1\">\n" +
                    // "                          <a:lumMod val=\"60000\"/>\n" +
                    // "                          <a:lumOff val=\"40000\"/>\n" +
                    // "                        </a:schemeClr>\n" +
                    //"                      </a:solidFill>\n" +
                    //"                    </a:ln>\n"
                    ;
            //String xml=  picture.getCTPicture().getSpPr().getLn().getSolidFill().xmlText();

            picture.getCTPicture().getSpPr().getLn().getSolidFill()
                    .set(XmlToken.Factory.parse(solidFillStr));


        }
        if ((null == picture.getCTPicture().getSpPr().getLn().getSolidFill())) {
            //picture.getCTPicture().getSpPr().getLn().
            picture.getCTPicture().getSpPr().getLn().addNewSolidFill();

        }
        //33A5FF: 51 165 255
        if ((null == picture.getCTPicture().getSpPr().getLn().getSolidFill().getSrgbClr())) {

//                                    picture.getCTPicture().getSpPr().getLn().getSolidFill().getSchemeClr().addNewRed();
            picture.getCTPicture().getSpPr().getLn().getSolidFill().addNewSrgbClr();
        }
        picture.getCTPicture().getSpPr().getLn().getSolidFill().getSrgbClr().
                setVal(new byte[]{(byte) (51), (byte) (165), (byte) (255)});
        picture.getCTPicture().getSpPr().getXfrm().getExt().setCx(width);  // 设置宽度
        picture.getCTPicture().getSpPr().getXfrm().getExt().setCy
                (EMUHeight);
//            picture.getCTPicture().getSpPr().getXfrm().getExt().setCy((long) 8 * 360000);  // 设置宽度
        if (picture.getCTPicture().getSpPr().getLn().isSetNoFill()) {
            picture.getCTPicture().getSpPr().getLn().unsetNoFill();
        }

    }

    public static void preMain(String[] args) throws InvalidFormatException {
        try {
            // 读取Word文档
            String srcFile = "D:/test_word/pureWord.docx";
            String desFile = "D:/test_word/modified_document.docx";
            FileInputStream fis = new FileInputStream(srcFile);
            FileOutputStream fos = new FileOutputStream(desFile);
            CustomXWPFDocument document = new CustomXWPFDocument(fis);
            //HWPFDocument document =new HWPFDocument(fis);

            int para = 1;
            int runCount = 1;
            int pic = 1;
            List<XWPFPicture> priPics = new ArrayList<XWPFPicture>();
            // 获取文档中的所有段落
            for (XWPFParagraph paragraph : document.getParagraphs()) {

//                XWPFParagraph currenPara= desDoc.createParagraph();
//                currenPara=paragraph;
                //paragraph.setBorderTop(Borders.valueOf(3));
                runCount = 1;
                System.out.println(para + " para.toString():    :" + paragraph.toString());
                System.out.println(para + " para  getStyle:" + paragraph.getStyle());
                System.out.println(para + " para.getText   :" + paragraph.getText());
                int pos = 0;
                // 获取段落中的所有Run
                for (XWPFRun run : paragraph.getRuns()) {
                    System.out.println(runCount + " run.text()   :" + run.text());
                    System.out.println("paragraph:" + para + "  run  :" + runCount + " run  有"
                            + run.getEmbeddedPictures().size() + "张图片");
//                    XWPFRun currenRun=currenPara.createRun();
//                    currenRun.setText(run.text());
//                    currenRun=run;
                    runCount++;
                    pic = 1;

                    //run.getEmbeddedPictures().removeAll(priPics);
                    // 获取Run中的所有Embedded Pictures
                    for (XWPFPicture picture : run.getEmbeddedPictures()) {
                        priPics.add(picture);
                        PoiSetWidth(picture, pic, (long) (14.3 * 360000), 8 * 360000);
                        CTInline inline = run.getCTR().getDrawingList().get(pic).getInlineArray(pic);
                        inline.getExtent().setCx((long) (14.3 * 360000));
                        //插入图片
//                        InputStream inputStream = new ByteArrayInputStream(decoderBytes);
//
//                        //为图片设置黑色边框
//                        XWPFPicture xwpfPicture = run.addPicture(inputStream, picture.getPictureData().getPictureType(), picture.getPictureData().getFileName(), Units.toEMU(picture.getWidth()), Units.toEMU(picture.getDepth()));
//                        xwpfPicture.getCTPicture().getSpPr().addNewLn().addNewSolidFill().addNewSchemeClr().setVal(STSchemeColorVal.Enum.forString("tx1"));
//                        inputStream.close();
//                        InputStream inputStream = new ByteArrayInputStream(decoderBytes);
//                        //为图片设置黑色边框
//                        XWPFPicture xwpfPicture = run.addPicture(inputStream, getPictureType(head), picture.getFileName(), Units.toEMU(picture.getWidth()), Units.toEMU(picture.getHeight()));
//                        xwpfPicture.getCTPicture().getSpPr().addNewLn().addNewSolidFill().addNewSchemeClr().setVal(STSchemeColorVal.Enum.forString("tx1"));

                        // 获取图片对象
                        XWPFPictureData pictureData = picture.getPictureData();
                        byte[] bytes = pictureData.getData();
                        double hight = Units.pixelToPoints(picture.getDepth()) * 2.54 / 1440.0 * 20.0;
                        // Pictures.ofBytes(bytes).sizeInCm(14.3,hight).create();
                        // Pictures.ofBytes(bytes).sizeInCm(14.3,picture.getDepth());
//                        // 修改图片的边框
                        // if(!picture.getCTPicture().getSpPr().isSetSolidFill())
                        String bid = document.addPictureData(picture.getPictureData().getData(), picture.getPictureData().getPictureType());
                        document.createPicture(paragraph, run, picture,
                                picture.getCTPicture().
                                        getBlipFill().getBlip().getEmbed(),
                                (int) picture.getCTPicture().getNvPicPr().getCNvPr().getId(),
                                (long) (14.3 * 360000),
                                (long) (8 * 360000),
                                //                           (int)picture.getWidth(),
                                //                          (int)picture.getDepth(),
                                19500,
                                "33A5FF", picture.getCTPicture().
                                        getBlipFill().getBlip().getEmbed());


//                            CTSRgbColor ctsRgbColor = CTSRgbColor.Factory.newInstance();
//                            ctsRgbColor.addNewBlue();
                        // CTLineProperties ctLine= new CTLineProperties() ;
//                           picture.getCTPicture().getSpPr().setLn(null);
//                           picture.getCTPicture().getSpPr().addNewLn().setW(19000);
//                            picture.getCTPicture().getSpPr().getLn().
//                                    addNewSolidFill().addNewSrgbClr();
//                            picture.getCTPicture().getSpPr().getLn().
//                                    getSolidFill().getSrgbClr().
//                                    setVal(new byte[]{(byte) (51),(byte)(165),(byte)(255)});
//
//                            // 设置边框颜色
////                       // picture.getCTPicture().getSpPr().addNewLn().setAlgn(STPenAlignment.Enum.forInt(1));  // 设置边框居中
                        float desiredWidthCm = 14.3f;
                        System.out.println("changed pic" + pic + "  picture.getCTPicture()  :" + picture.getCTPicture());
                        System.out.println("changed pic" + pic + "  picture.getDepth()  :" + picture.getDepth());
                        System.out.println("changed pic" + pic + "  picture.getWidth()  :" + picture.getWidth());
                        System.out.println("changed pic" + pic + "  picture.getPictureData()  :" + picture.getPictureData());
                        System.out.println("changed pic" + pic++ + "  picture.getDescription()  :" + picture.getDescription());
                    }
                    //run.getEmbeddedPictures().removeAll(priPics);
                }
                para++;
            }

            // 保存修改后的Word文档
            // FileOutputStream fos = new FileOutputStream("D:/test_word/modified_document.docx");
            document.write(fos);
            fos.flush();
            // 关闭资源
            fis.close();
            fos.close();
//删除图片
//            FileOutputStream delFos = new FileOutputStream("D:/test_word/complete.docx");
//            FileInputStream delFis = new FileInputStream(desFile);
//            CustomXWPFDocument desDoc= new CustomXWPFDocument(delFis);
//            System.out.println("一共："+desDoc.getParagraphs().size()+"个 paragraph");
//            int paraIndex=1;
//            for (XWPFParagraph par : desDoc.getParagraphs())
//            {
//                int pos = 0;
//                System.out.println("para"+paraIndex+++"一共："+par.getRuns().size()+"个 run");
//                while (pos < par.getRuns().size())
//                {
//                    XWPFRun run = par.getRuns().get(pos);
//                    double sumWidth=0f;
//                    System.out.println("run"+pos+"一共："+run.getEmbeddedPictures().size()+"个 embeded picture");
//                    for(XWPFPicture picture:run.getEmbeddedPictures())
//                    {
//                        sumWidth=(picture.getWidth()*12700);
//                        System.out.println(sumWidth);
//                        if (sumWidth!=(14.3*360000))
//                        {
//                            par.removeRun(pos);
//                            break;
//                        }
//                        else
//                        {
//                            pos++;
//                        }
//                    }
//
//                }
//            }
//            desDoc.write(delFos);
//            delFos.flush();
//            // 关闭资源
//            delFis.close();
//            delFos.close();
        } catch (IOException e) {
            e.printStackTrace();
        } catch (XmlException e) {
            throw new RuntimeException(e);
        }
    }

    public static void setAnchorAndInline(XWPFRun run, XWPFPicture picture, long width, long EMUHeight) throws XmlException {
        List<org.openxmlformats.schemas.wordprocessingml.x2006.main.CTDrawing> drawingList
                = run.getCTR().getDrawingList();
        for (org.openxmlformats.schemas.wordprocessingml.x2006.main.CTDrawing drawing : drawingList) {
            for (CTAnchor ctAnchor : drawing.getAnchorList()) {
                System.out.println(ctAnchor.toString());
                if ((ctAnchor.getGraphic().getGraphicData()).toString().
                        indexOf("blip r:embed=\"" + picture.getCTPicture().getBlipFill().getBlip().getEmbed() + "\"") > -1) {
                    System.out.println("图片 ：" + picture.getCTPicture().getBlipFill().getBlip().getEmbed()
                            + "的 CTAnchor被设置cx");
                    ctAnchor.getExtent().setCx(width);
                    ctAnchor.getExtent().setCy(EMUHeight);
                    System.out.println("CTAnchor被设置cx 后为 ：" + drawing.toString());
                    ctAnchor.getEffectExtent().setB("10000");
                    ctAnchor.getEffectExtent().setT("10000");
                    ctAnchor.getEffectExtent().setL("10000");
                    ctAnchor.getEffectExtent().setR("10000");
                    setAnchorToInline(run, picture, width, EMUHeight);
                    break;
                    //ctAnchor.getGraphic().toString();
                }
//                if(((ctAnchor.getGraphic()) instanceof XWPFPicture)&&(ctAnchor.getGraphic().equals(picture)))
//                {
//                    ctAnchor.getExtent().setCx(width);
//                }
            }

        }
        for (org.openxmlformats.schemas.wordprocessingml.x2006.main.CTDrawing drawing : drawingList) {

            for (CTInline ctInline : drawing.getInlineList()) {
                if ((ctInline.getGraphic().getGraphicData()).toString().indexOf
                        ("blip r:embed=\"" + picture.getCTPicture().getBlipFill().getBlip().getEmbed() + "\"") > -1) {

                    System.out.println("图片 ：" + picture.getCTPicture().getBlipFill().getBlip().getEmbed()
                            + "的 CTInLine被设置cx");
                    ctInline.getExtent().setCx(width);
                    ctInline.getExtent().setCy(EMUHeight);
                    System.out.println("CTInLine被设置cx 后为 ：" + drawing.toString());
                    ctInline.getEffectExtent().setB("10000");
                    ctInline.getEffectExtent().setT("10000");
                    ctInline.getEffectExtent().setL("10000");
                    ctInline.getEffectExtent().setR("10000");
                    // ctInline.getGraphic().toString());
                }
            }
        }
    }

    public static void collectModelStyles(HashMap<String, XWPFStyle> styleHashMap, XWPFTemplate modelTemp) {
        String[] lvls = {"0", "1", "2", "3", "4", "5", "6", "7", "8", "9", "10"};
        XWPFStyle styleNormal = modelTemp.getXWPFDocument().getStyles().
                getStyleWithName("Normal");
        //styleNormal.setStyleId("model" + styleNormal.getStyleId());
        styleHashMap.put("Normal", styleNormal);
        for (XWPFParagraph para : modelTemp.getXWPFDocument().getParagraphs()) {
            String lvl = getTitleLvl(modelTemp.getXWPFDocument(), para);
            if (para.getRuns() == null || para.getRuns().size() < 1) continue;
            if (para.getRuns().get(0).getEmbeddedPictures().size() > 0) {
                if (styleHashMap.get("picture") == null) {
                    XWPFStyle style = modelTemp.getXWPFDocument().getStyles().
                            getStyle(para.getStyleID());
                    CTString ctString = CTString.Factory.newInstance();
                    ctString.setVal(styleNormal.getCTStyle().getStyleId());
                   // style.getCTStyle().setBasedOn(ctString);

                    styleHashMap.put("picture", style);
                }
            } else if ((para.getStyleID() != null) && modelTemp.getXWPFDocument().getStyles().getStyle
                    (para.getStyleID()).getType().toString().equals("table")) {
                if (styleHashMap.get("table") == null) {
                    XWPFStyle style = modelTemp.getXWPFDocument().getStyles().
                            getStyle(para.getStyleID());
                    CTString ctString = CTString.Factory.newInstance();
                    ctString.setVal(styleNormal.getCTStyle().getStyleId());
                    //style.getCTStyle().setBasedOn(ctString);
                    styleHashMap.put("table", style);
                }
            }
//        else if(para.getDocument().getNumbering().getNums().size()>0)
//        {
//
//        }
            else {
                if ((lvl != null) && (lvl != "") && (styleHashMap.get(lvl) == null)) {
                    XWPFStyle style = modelTemp.getXWPFDocument().getStyles().
                            getStyle(para.getStyleID());
                    CTString ctString = CTString.Factory.newInstance();
                    ctString.setVal(styleNormal.getCTStyle().getStyleId());
                   // style.getCTStyle().setBasedOn(ctString);
                    styleHashMap.put(lvl, style);
                }
            }

        }
//    for(XWPFParagraph para:modelTemp.getXWPFDocument().getParagraphs())
//    {
//        XWPFStyle style=  modelTemp.getXWPFDocument().getStyles().
//                getStyle(para.getStyleID());
//        style.setStyleId("model"+style.getStyleId());
//    }
    }

    public static void main(String[] args) throws InvalidFormatException, XmlException, IOException {

        // 读取Word文档
//            String srcFile="F:\\2023双高建设\\终期个人负责\\过程性\\1.2.2.4省级研究课题：117项（20231128修改版本）.docx";
//            String desFile="F:\\2023双高建设\\终期个人负责\\过程性\\后\\1.2.2.4省级研究课题：117项（20231128修改版本）.docx";
        // String srcFile="D:\\test_word\\new.docx";
        //String desFile="D:\\test_word\\modifyFile.docx";
        String srcFileFolder = "D:\\modify_source\\";
        String desFile = "D:\\modify_source_target\\";
        String styleModelFile = "D:\\modify_source\\styles_model.docx";
        // HashMap<String,XWPFStyle>
        File desFileFolder = new File(desFile);
        desFileFolder.mkdir();
        File srcFileDir = new File(srcFileFolder);
        if (!srcFileDir.exists())
            srcFileDir.mkdir();
        File files[] = srcFileDir.listFiles();
        XWPFTemplate styleModel = XWPFTemplate.compile(styleModelFile);
        HashMap<String, XWPFStyle> styleHashMap = new HashMap<>();
        collectModelStyles(styleHashMap, styleModel);
        XWPFNumbering modelNumberbing = styleModel.getXWPFDocument().getNumbering();

        for (File srcFile : files) {
            if (srcFile.getName().equals("styles_model.docx")) {
                continue;
            }


            try {
                FileInputStream fis = new FileInputStream(srcFile);
                FileOutputStream fos = new FileOutputStream(desFile + srcFile.getName());
                // CustomXWPFDocument document = new CustomXWPFDocument(fis);
                //HWPFDocument document =new HWPFDocument(fis);
                XWPFTemplate xwpfTemplate = XWPFTemplate.compile(srcFile);
                //List<XWPFPicture> pictureList=
                XWPFDocument document = xwpfTemplate.getXWPFDocument();
                if(document.getNumbering()==null)document.createNumbering();
                for (XWPFNum n : modelNumberbing.getNums())
                {
                    document.getNumbering().addNum(n);
                }
                for (XWPFAbstractNum n : modelNumberbing.getAbstractNums())
                {
                    document.getNumbering().addAbstractNum(n);
                }

//             for (XWPFStyle style: styleModel.getXWPFDocument().getStyles().getUsedStyleList
//                     (styleModel.getXWPFDocument().getStyles().getStyleWithName("Normal")))
//             {
//                 document.getStyles().addStyle(style);
//             }
                //添加模板中的样式
                CTStyles ctStyles = styleModel.getXWPFDocument().getStyle();
                CTStyle[] ctArray = ctStyles.getStyleArray();
                //document.getStyles().addStyle(styleHashMap.get("Normal"));

                Map<String, String> styleMap = new HashMap<>();
                if (document.getStyles() == null) {
                    document.createStyles();
                }
                for (int styleId = 0; styleId < ctArray.length; styleId++)
                {

                    //  XWPFStyles styles=  styleModel.getXWPFDocument().getStyles();

                    XWPFStyle style = styleModel.getXWPFDocument().
                            getStyles().getStyle(ctArray[styleId].getStyleId());
                    if (style == null)
                    {
                        continue;
                    }
                    //document.getStyle().set(styleModel.getXWPFDocument().getStyle());
                    if(document.getStyles().getStyle(style.getStyleId())!=null)
                    {
                        //document.getStyles().addStyle(style);
                        document.getStyles().getStyle(style.getStyleId()).setStyle(style.getCTStyle());
                    }
                    else
                    {
                        document.getStyles().addStyle(style);
                    }
                    try
                    {
                        String le = ctArray[styleId].getPPr().getOutlineLvl().toString();
                        System.out.println(le);
                        if (styleMap.get(le) == null) {
                            styleMap.put(le, ctArray[styleId].getStyleId() + ",," + ctArray[styleId].getName());
                        }
                    } catch (Exception e) {
                        styleMap.put("text", ctArray[styleId].getStyleId() + ",," + ctArray[styleId].getName());
                    }
                }

                //xwpfTemplate.render();
//            for(int i=0;i< pictureList.size();i++)
//            {
//                pictureList.get(i)
//            }
                int para = 1;
                int runCount = 1;
                int pic = 1;
                List<XWPFPicture> priPics = new ArrayList<XWPFPicture>();
//                for(IBodyElement iBodyElement:document.getBodyElements())
//                {
//                   String string= iBodyElement.getElementType().name();
//                   System.out.println("iBodyElement.getElementType().name()= "+string);
//                   string = iBodyElement.getPartType().name();
//                    System.out.println("iBodyElement.getPartType().name()= "+string);
//                }
                // 获取文档中的所有段落
                for (XWPFParagraph paragraph : document.getParagraphs())
                {

                    if(paragraph.getText()!="")
                        paragraph.setStyle(document.getStyles().getStyle("3").getStyleId());

                    String paraType = "";
//                    if (paragraph.getDocument().getTables().size() > 0) {
//                        paraType = "table";
//                        if (document.getStyles().getStyle(styleHashMap.get("table").getStyleId()) != null) {
//                            document.getStyles().addStyle(styleHashMap.get("table"));
//                        }
//                        paragraph.setStyle(styleHashMap.get("table").getStyleId());
//                    } else if (paragraph.getDocument().getAllPictures().size() > 0) {
//                        paraType = "picture";
//                        if (document.getStyles().getStyle(styleHashMap.get("picture").getStyleId()) != null) {
//                            document.getStyles().addStyle(styleHashMap.get("picture"));
//                        }
//                        paragraph.setStyle(styleHashMap.get("picture").getStyleId());
//                    } else if (paragraph.getText().length() > 0) {
//                        String lvl = getTitleLvl(document, paragraph);
//                        if (lvl != null && lvl != "") {
//                            if (document.getStyles().getStyle(styleHashMap.get(lvl).getStyleId()) != null) {
//                                document.getStyles().addStyle(styleHashMap.get(lvl));
//                            }
//                            paragraph.setStyle(styleHashMap.get(lvl).getStyleId());
//                        } else
//                            paragraph.setStyle(styleHashMap.get("Normal").getStyleId());
//
//
//                    }


                    String eleType = String.valueOf(paragraph.getElementType());
                    String tempLevel = getTitleLvl(document, paragraph);
                    XWPFStyle style = document.getStyles().
                            getStyle(paragraph.getStyleID() == null ? "" : paragraph.getStyleID());
                    runCount = 1;
                    System.out.println(para + " para.toString():    :" + paragraph.toString());
                    System.out.println(para + " para  getStyle:" + paragraph.getStyle());
                    System.out.println(para + " para.getText   :" + paragraph.getText());
                    int pos = 0;
                    int pictureAmount = 0;
                    int amountOneLine = 0;
                    long sumWidth = 0l;
                    List<Long> picWidthArray = new ArrayList<Long>();
                    List<List<XWPFPicture>> pictureGroups = new ArrayList();
                    // 获取段落中的所有Run
                    for (XWPFRun run : paragraph.getRuns()) {
                        pictureAmount += run.getEmbeddedPictures().size();
                        for (int i = 0; i < run.getEmbeddedPictures().size(); i++) {
                            priPics.add(run.getEmbeddedPictures().get(i));
                            XWPFPicture picture = run.getEmbeddedPictures().get(i);
                            picWidthArray.add(picture.getCTPicture().getSpPr().getXfrm().getExt().getCx());
                            sumWidth += picture.getCTPicture().getSpPr().getXfrm().getExt().getCx();
                            if ((sumWidth > 16 * 360000) && (priPics.size() > 1)) {
                                priPics.remove(picture);
                                pictureGroups.add(priPics);
                                priPics = new ArrayList<XWPFPicture>();
                                priPics.add(picture);
                                sumWidth = picture.getCTPicture().getSpPr().getXfrm().getExt().getCx();
                            }
                        }

                    }
                    pictureGroups.add(priPics);
                    priPics = new ArrayList<XWPFPicture>();
//                for(List<XWPFPicture> list:pictureGroups)
//                {
//                    for(XWPFPicture picture:list)
//                    {
//                        float desiredWidthCm = 14.3f;//厘米
//                        PoiSetWidth(picture,pic,(long)(desiredWidthCm*360000/pictureAmount));
//                        System.out.println("changed pic"+pic+"  picture.getCTPicture()  :"+picture.getCTPicture());
//                        setAnchorAndInline(run,picture,(long)(desiredWidthCm*360000/pictureAmount));
//                    }
//                }
                    for (XWPFRun run : paragraph.getRuns()) {
                        System.out.println(runCount + " run.text()   :" + run.text());
                        System.out.println("paragraph:" + para + "  run  :" + runCount + " run  有"
                                + run.getEmbeddedPictures().size() + "张图片");
//                    XWPFRun currenRun=currenPara.createRun();
//                    currenRun.setText(run.text());
//                    currenRun=run;
                        runCount++;
                        pic = 1;

                        //run.getEmbeddedPictures().removeAll(priPics);
                        // 获取Run中的所有Embedded Pictures
                        for (XWPFPicture picture : run.getEmbeddedPictures()) {
                            // picture.getPictureData()
                            pictureAmount = 1;
                            float desiredWidthCm = 14.3f;//厘米
                            long desiredHeight = 0l;//EMU
                            //priPics.add(picture);
                            //if( picture.getCTPicture().getSpPr().getXfrm().getExt().getCx()/360000)


                            for (List<XWPFPicture> list : pictureGroups) {
                                if (list.indexOf(picture) > -1) {
                                    pictureAmount = list.size();
                                    for (XWPFPicture pp : list) {
                                        if (pp.getCTPicture().getSpPr().getXfrm().getExt().getCy() > desiredHeight) {
                                            desiredHeight = pp.getCTPicture().getSpPr().getXfrm().getExt().getCy();
                                        }
                                    }
                                    break;
                                }
                            }
                            // paragraph.
                            paragraph.setIndentationLeft(10);
                            paragraph.setIndentationRight(10);
                            paragraph.setFirstLineIndent(0);
                            paragraph.setAlignment(ParagraphAlignment.CENTER);
                            PoiSetWidth(picture, pic, (long) (desiredWidthCm * 360000 / pictureAmount), desiredHeight);
                            System.out.println("changed pic" + pic + "  picture.getCTPicture()  :" + picture.getCTPicture());
                            setAnchorAndInline(run, picture, (long) (desiredWidthCm * 360000 / pictureAmount), desiredHeight);


                            // 获取图片对象

                            // word标准布局的页边距
                            long LEFT_MARGIN = 1800L;
                            long RIGHT_MARGIN = 1800L;
                            long TOP_MARGIN = 1440L;
                            long BOTTOM_MARGIN = 1440L;
                            double hight = Units.pixelToPoints(picture.getDepth()) * 2.54 / 1440.0 * 20.0;
                            // Pictures.ofBytes(bytes).sizeInCm(14.3,hight).create();
                            // Pictures.ofBytes(bytes).sizeInCm(14.3,picture.getDepth());
//                        // 修改图片的边框
                            // if(!picture.getCTPicture().getSpPr().isSetSolidFill())
//                        String bid= document.addPictureData(picture.getPictureData().getData(),picture.getPictureData().getPictureType());
//                        document.createPicture(paragraph,run,picture,
//                                picture.getCTPicture().
//                                        getBlipFill().getBlip().getEmbed(),
//                                (int)picture.getCTPicture().getNvPicPr().getCNvPr().getId(),
//                                (long)(14.3*360000),
//                                (long)(8*360000),
//                                //                           (int)picture.getWidth(),
//                                //                          (int)picture.getDepth(),
//                                19500,
//                                "33A5FF",picture.getCTPicture().
//                                        getBlipFill().getBlip().getEmbed());


//                            CTSRgbColor ctsRgbColor = CTSRgbColor.Factory.newInstance();
//                            ctsRgbColor.addNewBlue();
                            // CTLineProperties ctLine= new CTLineProperties() ;
//                           picture.getCTPicture().getSpPr().setLn(null);
//                           picture.getCTPicture().getSpPr().addNewLn().setW(19000);
//                            picture.getCTPicture().getSpPr().getLn().
//                                    addNewSolidFill().addNewSrgbClr();
//                            picture.getCTPicture().getSpPr().getLn().
//                                    getSolidFill().getSrgbClr().
//                                    setVal(new byte[]{(byte) (51),(byte)(165),(byte)(255)});
//
//                            // 设置边框颜色
////                       // picture.getCTPicture().getSpPr().addNewLn().setAlgn(STPenAlignment.Enum.forInt(1));  // 设置边框居中

//                        System.out.println("changed pic"+pic+"  picture.getCTPicture()  :"+picture.getCTPicture());
//                        System.out.println("changed pic"+pic+"  picture.getDepth()  :"+picture.getDepth());
//                        System.out.println("changed pic"+pic+"  picture.getWidth()  :"+picture.getWidth());
//                        System.out.println("changed pic"+pic+"  picture.getPictureData()  :"+picture.getPictureData());
//                        System.out.println("changed pic"+pic+++"  picture.getDescription()  :"+picture.getDescription());
                        }
                        //run.getEmbeddedPictures().removeAll(priPics);
                    }
                    para++;
                }
                List<XWPFParagraph> paragraphLists = document.getParagraphs();
                for (XWPFParagraph p : paragraphLists) {
                    List<XWPFRun> xwpfRuns = p.getRuns();
                    for (XWPFRun r : xwpfRuns) {
                        List<CTDrawing> drawingList = r.getCTR().getDrawingList();
                        for (CTDrawing d : drawingList) {
                            List<CTAnchor> anchors = d.getAnchorList();
                            anchors.removeAll(anchors);
                        }
                    }
                }
                // 保存修改后的Word文档
                // FileOutputStream fos = new FileOutputStream("D:/test_word/modified_document.docx");
                //styleModel.getXWPFDocument().getDocument().set(xwpfTemplate.getXWPFDocument().getDocument());
                // styleModel.writeAndClose(fos);
                xwpfTemplate.writeAndClose(fos);

                //styleModel.writeAndClose(fos);
                fos.flush();
                // 关闭资源
                fis.close();
                fos.close();

            } catch (IOException | XmlException e) {
                e.printStackTrace();
            }
        }
        styleModel.close();
    }
}
