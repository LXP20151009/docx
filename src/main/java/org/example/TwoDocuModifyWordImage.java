package org.example;

import com.deepoove.poi.XWPFTemplate;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.util.Units;
import org.apache.poi.xwpf.usermodel.*;
import org.apache.xmlbeans.XmlException;
import org.apache.xmlbeans.XmlToken;
import org.openxmlformats.schemas.drawingml.x2006.main.impl.CTNonVisualDrawingPropsImpl;
import org.openxmlformats.schemas.drawingml.x2006.wordprocessingDrawing.CTAnchor;
import org.openxmlformats.schemas.drawingml.x2006.wordprocessingDrawing.CTInline;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTDrawing;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTPicture;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

public class TwoDocuModifyWordImage {

    public static void setAnchorToInline(XWPFRun run,XWPFPicture picture,long width,long EMUHeight) throws XmlException {
        List<CTDrawing> drawingList
                =  run.getCTR().getDrawingList();
        for(CTDrawing drawing:drawingList)
        {
            for(CTAnchor ctAnchor :drawing.getAnchorList())
            {
                if((ctAnchor.getGraphic().getGraphicData()).toString().
                        indexOf("blip r:embed=\""+picture.getCTPicture().getBlipFill().getBlip().getEmbed()+"\"")>-1)
                {
                    String picStr=picture.toString();
                    org.openxmlformats.schemas.drawingml.x2006.picture.CTPicture
                    ctPicture = picture.getCTPicture();
                    ctPicture.getSpPr().set(picture.getCTPicture().getSpPr());
                    ctPicture.getBlipFill().set(picture.getCTPicture().getBlipFill());
                    ctPicture.getNvPicPr().set(picture.getCTPicture().getSpPr());
                    XWPFPicture fakePic = CustomXWPFDocument.createAndReturnPicture
                            (run,picture,picture.getCTPicture().getBlipFill().getBlip().getEmbed(),
                                    (int)picture.getCTPicture().getNvPicPr().getCNvPr().getId(),
                                    (long)picture.getWidth(),(long)picture.getDepth(),
                                    9500,
                                    "33A5FF",
                                    picture.getCTPicture().getBlipFill().getBlip().getEmbed());
                    if(null==picture.getCTPicture().getNvPicPr().getCNvPr())
                    {
                        //picture.getCTPicture().getNvPicPr().getCNvPr()=new CTNonVisualDrawingPropsImpl()
                    }
                    ctPicture.getNvPicPr().getCNvPr().set(picture.getCTPicture().getNvPicPr().getCNvPr());
                    ctAnchor.getEffectExtent().setB("10000");
                    ctAnchor.getEffectExtent().setT("10000");
                    ctAnchor.getEffectExtent().setL("10000");
                    ctAnchor.getEffectExtent().setR("10000");
                    String inlineStr="<wp:inline distT=\"0\" distB=\"0\" distL=\"0\" distR=\"0\" wp14:anchorId=\"2764B781\" wp14:editId=\"3860602A\" xmlns:wpc=\"http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas\" xmlns:cx=\"http://schemas.microsoft.com/office/drawing/2014/chartex\" xmlns:cx1=\"http://schemas.microsoft.com/office/drawing/2015/9/8/chartex\" xmlns:cx2=\"http://schemas.microsoft.com/office/drawing/2015/10/21/chartex\" xmlns:cx3=\"http://schemas.microsoft.com/office/drawing/2016/5/9/chartex\" xmlns:cx4=\"http://schemas.microsoft.com/office/drawing/2016/5/10/chartex\" xmlns:cx5=\"http://schemas.microsoft.com/office/drawing/2016/5/11/chartex\" xmlns:cx6=\"http://schemas.microsoft.com/office/drawing/2016/5/12/chartex\" xmlns:cx7=\"http://schemas.microsoft.com/office/drawing/2016/5/13/chartex\" xmlns:cx8=\"http://schemas.microsoft.com/office/drawing/2016/5/14/chartex\" xmlns:mc=\"http://schemas.openxmlformats.org/markup-compatibility/2006\" xmlns:aink=\"http://schemas.microsoft.com/office/drawing/2016/ink\" xmlns:am3d=\"http://schemas.microsoft.com/office/drawing/2017/model3d\" xmlns:o=\"urn:schemas-microsoft-com:office:office\" xmlns:oel=\"http://schemas.microsoft.com/office/2019/extlst\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" xmlns:m=\"http://schemas.openxmlformats.org/officeDocument/2006/math\" xmlns:v=\"urn:schemas-microsoft-com:vml\" xmlns:wp14=\"http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing\" xmlns:wp=\"http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing\" xmlns:w10=\"urn:schemas-microsoft-com:office:word\" xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\" xmlns:w14=\"http://schemas.microsoft.com/office/word/2010/wordml\" xmlns:w15=\"http://schemas.microsoft.com/office/word/2012/wordml\" xmlns:w16cex=\"http://schemas.microsoft.com/office/word/2018/wordml/cex\" xmlns:w16cid=\"http://schemas.microsoft.com/office/word/2016/wordml/cid\" xmlns:w16=\"http://schemas.microsoft.com/office/word/2018/wordml\" xmlns:w16sdtdh=\"http://schemas.microsoft.com/office/word/2020/wordml/sdtdatahash\" xmlns:w16se=\"http://schemas.microsoft.com/office/word/2015/wordml/symex\" xmlns:wpg=\"http://schemas.microsoft.com/office/word/2010/wordprocessingGroup\" xmlns:wpi=\"http://schemas.microsoft.com/office/word/2010/wordprocessingInk\" xmlns:wne=\"http://schemas.microsoft.com/office/word/2006/wordml\" xmlns:wps=\"http://schemas.microsoft.com/office/word/2010/wordprocessingShape\">\n" +
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
                    CTDrawing fakeDrawing=//(CTDrawing)XmlToken.Factory.parse(inlineStr);
                            (CTDrawing) CTDrawing.Factory.parse(inlineStr);
//                            (CTDrawing) drawingList.get(drawingList.indexOf(drawing))
//                            .set(XmlToken.Factory.parse(inlineStr));
                    CTInline ctInline=fakeDrawing.getInlineList().get(0);
                    ctInline.getExtent().set(ctAnchor.getExtent());
                    ctInline.getDocPr().set(ctAnchor.getDocPr());
                    ctInline.getCNvGraphicFramePr().set(ctAnchor.getCNvGraphicFramePr());
                    ctInline.getGraphic().set(ctAnchor.getGraphic());
                    picture.getCTPicture().set(ctPicture);
                    picture= fakePic;
                    if(null==drawing.getInlineList())
                    {

                        drawing.setInlineArray(new CTInline[]{ctInline});
                    }
                    if(drawing.getInlineList().size()==0)
                            drawing.setInlineArray(new CTInline[]{ctInline});
                    drawing.getInlineList().get(0).set(ctInline);
//                    CTInline []ctInline1=new CTInline[drawing.getInlineArray().length+1];
//                    for(int i=1;i<ctInline1.length;i++)
//                    {
//                        ctInline1[i]=drawing.getInlineArray()[i-1];
//                    }
//                    ctInline1[0]=ctInline;
//                    drawing.setInlineArray(ctInline1);
                    drawing.getAnchorList().remove(ctAnchor);
//                    ctAnchor.getExtent().setCx(0);
//                    ctAnchor.getExtent().setCy(0);
                   // drawing.getAnchorList().remove(ctAnchor);
                   // drawing.set(fakeDrawing);
                    System.out.println("替换好的 drawing为：  "+drawing.toString());
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

    public static void PoiSetWidth(XWPFPicture picture ,int pic,long width,long EMUHeight ) throws XmlException
    {
        System.out.println("pic" + pic + "  picture.getCTPicture()  :" + picture.getCTPicture());
        System.out.println("pic" + pic + "  picture.getDepth()  :" + picture.getDepth());
        System.out.println("pic" + pic + "  picture.getWidth()  :" + picture.getWidth());
        System.out.println("pic" + pic + "  picture.getPictureData()  :" + picture.getPictureData());
        System.out.println("pic" + pic + "  picture.getDescription()  :" + picture.getDescription());


        if ((null == picture.getCTPicture().getSpPr().getLn()))
        {
            picture.getCTPicture().getSpPr().addNewLn().setW(9500);
        }
        else
        {
            picture.getCTPicture().getSpPr().getLn().setW(9500);
        }

        if ((null != picture.getCTPicture().getSpPr().getLn().getSolidFill()))
        {
            //picture.getCTPicture().getSpPr().getLn().
            String solidFillStr=
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
        if ((null == picture.getCTPicture().getSpPr().getLn().getSolidFill()))
        {
            //picture.getCTPicture().getSpPr().getLn().
            picture.getCTPicture().getSpPr().getLn().addNewSolidFill();

        }
        //33A5FF: 51 165 255
        if ((null == picture.getCTPicture().getSpPr().getLn().getSolidFill().getSrgbClr()))
        {

//                                    picture.getCTPicture().getSpPr().getLn().getSolidFill().getSchemeClr().addNewRed();
            picture.getCTPicture().getSpPr().getLn().getSolidFill().addNewSrgbClr();
        }
        picture.getCTPicture().getSpPr().getLn().getSolidFill().getSrgbClr().
                setVal(new byte[]{(byte) (51), (byte) (165), (byte) (255)});
        picture.getCTPicture().getSpPr().getXfrm().getExt().setCx(width);  // 设置宽度
        picture.getCTPicture().getSpPr().getXfrm().getExt().setCy
                (EMUHeight);
//            picture.getCTPicture().getSpPr().getXfrm().getExt().setCy((long) 8 * 360000);  // 设置宽度
        if(picture.getCTPicture().getSpPr().getLn().isSetNoFill())
        {
            picture.getCTPicture().getSpPr().getLn().unsetNoFill();
        }

    }
    public static void preMain(String[] args) throws InvalidFormatException
    {
        try
        {
            // 读取Word文档
            String srcFile="D:/test_word/pureWord.docx";
            String desFile="D:/test_word/modified_document.docx";
            FileInputStream fis = new FileInputStream(srcFile);
            FileOutputStream fos = new FileOutputStream(desFile);
            CustomXWPFDocument document = new CustomXWPFDocument(fis);
            //HWPFDocument document =new HWPFDocument(fis);

            int para=1;
            int  runCount=1;
            int  pic=1;
            List<XWPFPicture> priPics=new ArrayList<XWPFPicture>() ;
            // 获取文档中的所有段落
            for (XWPFParagraph paragraph : document.getParagraphs())
            {

//                XWPFParagraph currenPara= desDoc.createParagraph();
//                currenPara=paragraph;
                //paragraph.setBorderTop(Borders.valueOf(3));
                runCount=1;
                System.out.println(para+" para.toString():    :"+paragraph.toString());
                System.out.println(para+" para  getStyle:"+paragraph.getStyle());
                System.out.println(para+" para.getText   :"+paragraph.getText());
                int pos = 0;
                // 获取段落中的所有Run
                for (XWPFRun run : paragraph.getRuns())
                {
                    System.out.println(runCount+" run.text()   :"+run.text());
                    System.out.println("paragraph:"+para+"  run  :"+runCount+" run  有"
                            +run.getEmbeddedPictures().size()+"张图片");
//                    XWPFRun currenRun=currenPara.createRun();
//                    currenRun.setText(run.text());
//                    currenRun=run;
                    runCount++;
                    pic=1;

                    //run.getEmbeddedPictures().removeAll(priPics);
                    // 获取Run中的所有Embedded Pictures
                    for (XWPFPicture picture : run.getEmbeddedPictures())
                    {
                        priPics.add(picture);
                        PoiSetWidth(picture,pic,(long)(14.3*360000),8*360000);
                        CTInline inline= run.getCTR().getDrawingList().get(pic).getInlineArray(pic);
                        inline.getExtent().setCx((long)(14.3*360000));
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
                        double hight=  Units.pixelToPoints(picture.getDepth()) * 2.54 / 1440.0 * 20.0;
                        // Pictures.ofBytes(bytes).sizeInCm(14.3,hight).create();
                        // Pictures.ofBytes(bytes).sizeInCm(14.3,picture.getDepth());
//                        // 修改图片的边框
                        // if(!picture.getCTPicture().getSpPr().isSetSolidFill())
                        String bid= document.addPictureData(picture.getPictureData().getData(),picture.getPictureData().getPictureType());
                        document.createPicture(paragraph,run,picture,
                                picture.getCTPicture().
                                        getBlipFill().getBlip().getEmbed(),
                                (int)picture.getCTPicture().getNvPicPr().getCNvPr().getId(),
                                (long)(14.3*360000),
                                (long)(8*360000),
                                //                           (int)picture.getWidth(),
                                //                          (int)picture.getDepth(),
                                19500,
                                "33A5FF",picture.getCTPicture().
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
                        System.out.println("changed pic"+pic+"  picture.getCTPicture()  :"+picture.getCTPicture());
                        System.out.println("changed pic"+pic+"  picture.getDepth()  :"+picture.getDepth());
                        System.out.println("changed pic"+pic+"  picture.getWidth()  :"+picture.getWidth());
                        System.out.println("changed pic"+pic+"  picture.getPictureData()  :"+picture.getPictureData());
                        System.out.println("changed pic"+pic+++"  picture.getDescription()  :"+picture.getDescription());
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

        }
        catch (IOException e)
        {
            e.printStackTrace();
        } catch (XmlException e)
        {
            throw new RuntimeException(e);
        }
    }
    public static void setAnchorAndInline(XWPFRun run,XWPFPicture picture,long width,long EMUHeight) throws XmlException {
        List<CTDrawing> drawingList
                =  run.getCTR().getDrawingList();
        for(CTDrawing drawing:drawingList)
        {
            for(CTAnchor ctAnchor :drawing.getAnchorList())
            {
               System.out.println(ctAnchor.toString());
                if((ctAnchor.getGraphic().getGraphicData()).toString().
                        indexOf("blip r:embed=\""+picture.getCTPicture().getBlipFill().getBlip().getEmbed()+"\"")>-1)
                {
                    System.out.println("图片 ："+picture.getCTPicture().getBlipFill().getBlip().getEmbed()
                            +"的 CTAnchor被设置cx");
                    ctAnchor.getExtent().setCx(width);
                    ctAnchor.getExtent().setCy(EMUHeight);
                    System.out.println("CTAnchor被设置cx 后为 ："+drawing.toString());
                    ctAnchor.getEffectExtent().setB("10000");
                    ctAnchor.getEffectExtent().setT("10000");
                    ctAnchor.getEffectExtent().setL("10000");
                    ctAnchor.getEffectExtent().setR("10000");
                    setAnchorToInline(run,picture,width,EMUHeight);
                    break;
                    //ctAnchor.getGraphic().toString();
                }
//                if(((ctAnchor.getGraphic()) instanceof XWPFPicture)&&(ctAnchor.getGraphic().equals(picture)))
//                {
//                    ctAnchor.getExtent().setCx(width);
//                }
            }

        }
        for(CTDrawing drawing:drawingList)
        {
            for(CTInline ctInline :drawing.getInlineList())
            {
                if((ctInline.getGraphic().getGraphicData()).toString().indexOf
                        ("blip r:embed=\""+picture.getCTPicture().getBlipFill().getBlip().getEmbed()+"\"")>-1)
                {

                    System.out.println("图片 ："+picture.getCTPicture().getBlipFill().getBlip().getEmbed()
                            +"的 CTInLine被设置cx");
                    ctInline.getExtent().setCx(width);
                    ctInline.getExtent().setCy(EMUHeight);
                    System.out.println("CTInLine被设置cx 后为 ："+drawing.toString());
                    ctInline.getEffectExtent().setB("10000");
                    ctInline.getEffectExtent().setT("10000");
                    ctInline.getEffectExtent().setL("10000");
                    ctInline.getEffectExtent().setR("10000");
                    // ctInline.getGraphic().toString());
                }
            }
        }
    }
    public static void main(String[] args) throws InvalidFormatException, XmlException
    {
        try
        {
            // 读取Word文档
//            String srcFile="F:\\2023双高建设\\终期个人负责\\过程性\\1.2.2.4省级研究课题：117项（20231128修改版本）.docx";
//            String desFile="F:\\2023双高建设\\终期个人负责\\过程性\\后\\1.2.2.4省级研究课题：117项（20231128修改版本）.docx";
            String srcFile="D:\\test_word\\new.docx";
            String desFile="D:\\test_word\\modifyFile.docx";
            FileInputStream fis = new FileInputStream(srcFile);
            FileOutputStream fos = new FileOutputStream(desFile);
            // CustomXWPFDocument document = new CustomXWPFDocument(fis);
            //HWPFDocument document =new HWPFDocument(fis);
            XWPFTemplate  xwpfTemplate= XWPFTemplate.compile(srcFile);
            //List<XWPFPicture> pictureList=
            XWPFDocument document=        xwpfTemplate.getXWPFDocument();
//            for(int i=0;i< pictureList.size();i++)
//            {
//                pictureList.get(i)
//            }
            int para=1;
            int  runCount=1;
            int  pic=1;
            List<XWPFPicture> priPics=new ArrayList<XWPFPicture>() ;
            // 获取文档中的所有段落
            for (XWPFParagraph paragraph : document.getParagraphs())
            {



                runCount=1;
                System.out.println(para+" para.toString():    :"+paragraph.toString());
                System.out.println(para+" para  getStyle:"+paragraph.getStyle());
                System.out.println(para+" para.getText   :"+paragraph.getText());
                int pos = 0;
                int pictureAmount=0;
                int amountOneLine=0;
                long sumWidth=0l;
                List<Long> picWidthArray=new ArrayList<Long>();
                List<List<XWPFPicture>> pictureGroups =new ArrayList();
                // 获取段落中的所有Run
                for (XWPFRun run : paragraph.getRuns())
                {
                    pictureAmount += run.getEmbeddedPictures().size();
                    for (int i=0;i<run.getEmbeddedPictures().size();i++)
                    {
                        priPics.add(run.getEmbeddedPictures().get(i));
                        XWPFPicture picture = run.getEmbeddedPictures().get(i);
                        picWidthArray.add(picture.getCTPicture().getSpPr().getXfrm().getExt().getCx());
                        sumWidth+=picture.getCTPicture().getSpPr().getXfrm().getExt().getCx();
                        if(sumWidth>14.3*360000)
                        {
                            priPics.remove(picture);
                            pictureGroups.add(priPics);
                            priPics=new ArrayList<XWPFPicture>();
                            priPics.add(picture);
                            sumWidth=picture.getCTPicture().getSpPr().getXfrm().getExt().getCx();
                        }
                    }

                }
                pictureGroups.add(priPics);
                priPics=new ArrayList<XWPFPicture>();
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
                for (XWPFRun run : paragraph.getRuns())
                {
                    System.out.println(runCount+" run.text()   :"+run.text());
                    System.out.println("paragraph:"+para+"  run  :"+runCount+" run  有"
                            +run.getEmbeddedPictures().size()+"张图片");
//                    XWPFRun currenRun=currenPara.createRun();
//                    currenRun.setText(run.text());
//                    currenRun=run;
                    runCount++;
                    pic=1;

                    //run.getEmbeddedPictures().removeAll(priPics);
                    // 获取Run中的所有Embedded Pictures
                    for (XWPFPicture picture : run.getEmbeddedPictures())
                    {
                        // picture.getPictureData()
                        pictureAmount=1;
                        float desiredWidthCm = 14.3f;//厘米
                        long desiredHeight=0l;//EMU
                        for(List<XWPFPicture> list:pictureGroups)
                        {
                            if(list.indexOf(picture)>-1)
                            {
                                pictureAmount=list.size();
                                for(XWPFPicture pp:list)
                                {
                                    if (pp.getCTPicture().getSpPr().getXfrm().getExt().getCy()>desiredHeight)
                                    {
                                        desiredHeight=pp.getCTPicture().getSpPr().getXfrm().getExt().getCy();
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
                        PoiSetWidth(picture,pic,(long)(desiredWidthCm*360000/pictureAmount),desiredHeight);
                        System.out.println("changed pic"+pic+"  picture.getCTPicture()  :"+picture.getCTPicture());
                        setAnchorAndInline(run,picture,(long)(desiredWidthCm*360000/pictureAmount),desiredHeight);

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

                        // word标准布局的页边距
                        long LEFT_MARGIN = 1800L;
                        long RIGHT_MARGIN = 1800L;
                        long TOP_MARGIN = 1440L;
                        long BOTTOM_MARGIN = 1440L;

//                        CTSectPr sectPr = document.getDocument().getBody().getSectPr();
//                        CTPageMar pageMar = sectPr.getPgMar();
//                        pageMar.setLeft(BigInteger.valueOf(LEFT_MARGIN));
//                        pageMar.setRight(BigInteger.valueOf(RIGHT_MARGIN));
//                        pageMar.setTop(BigInteger.valueOf(TOP_MARGIN));
//                        pageMar.setBottom(BigInteger.valueOf(BOTTOM_MARGIN));
                        double hight=  Units.pixelToPoints(picture.getDepth()) * 2.54 / 1440.0 * 20.0;
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

            // 保存修改后的Word文档
            // FileOutputStream fos = new FileOutputStream("D:/test_word/modified_document.docx");
            //document.write(fos);
            xwpfTemplate.writeAndClose(fos);
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
        }
        catch (IOException | XmlException e)
        {
            e.printStackTrace();
        }
    }
}
