package org.example;

import com.deepoove.poi.XWPFTemplate;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.util.Units;
import org.apache.poi.xwpf.usermodel.*;
import org.apache.xmlbeans.XmlException;
import org.apache.xmlbeans.XmlObject;
import org.apache.xmlbeans.XmlToken;
import org.openxmlformats.schemas.drawingml.x2006.wordprocessingDrawing.CTInline;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

public class AModifyWordImage {

    public static void PoiSetWidth(XWPFPicture picture ,int pic,float width ) throws XmlException {
        System.out.println("pic" + pic + "  picture.getCTPicture()  :" + picture.getCTPicture());
        System.out.println("pic" + pic + "  picture.getDepth()  :" + picture.getDepth());
        System.out.println("pic" + pic + "  picture.getWidth()  :" + picture.getWidth());
        System.out.println("pic" + pic + "  picture.getPictureData()  :" + picture.getPictureData());
        System.out.println("pic" + pic + "  picture.getDescription()  :" + picture.getDescription());
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
        // Pictures.ofBytes(bytes).sizeInCm(14.3,hight).create();
        // Pictures.ofBytes(bytes).sizeInCm(14.3,picture.getDepth());
//                        // 修改图片的边框
        // if(!picture.getCTPicture().getSpPr().isSetSolidFill())

//                            document.createPicture(paragraph,run,picture,
//                                    picture.getCTPicture().
//                                            getBlipFill().getBlip().getEmbed(),
//                                    (int)picture.getCTPicture().getNvPicPr().getCNvPr().getId(),
//                                    (long)(14.3*360000),
//                                    (long)(8*360000),
////                                    (int)picture.getWidth(),
////                                    (int)picture.getDepth(),
//                                    19500,
//                                    "33A5FF","");
            pic++;

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

            if ((null == picture.getCTPicture().getSpPr().getLn()))
            {
                picture.getCTPicture().getSpPr().addNewLn().setW(9500);
            } else {
                picture.getCTPicture().getSpPr().getLn().setW(9500);
            }
        if ((null != picture.getCTPicture().getSpPr().getLn().getSolidFill()))
        {
            //picture.getCTPicture().getSpPr().getLn().
            String solidFillStr=
                    "<a:SrgbClr val=\"33A5FF\" xmlns:pic=\"http://schemas.openxmlformats.org/drawingml/2006/picture\" xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" xmlns:wpc=\"http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas\" xmlns:cx=\"http://schemas.microsoft.com/office/drawing/2014/chartex\" xmlns:cx1=\"http://schemas.microsoft.com/office/drawing/2015/9/8/chartex\" xmlns:cx2=\"http://schemas.microsoft.com/office/drawing/2015/10/21/chartex\" xmlns:cx3=\"http://schemas.microsoft.com/office/drawing/2016/5/9/chartex\" xmlns:cx4=\"http://schemas.microsoft.com/office/drawing/2016/5/10/chartex\" xmlns:cx5=\"http://schemas.microsoft.com/office/drawing/2016/5/11/chartex\" xmlns:cx6=\"http://schemas.microsoft.com/office/drawing/2016/5/12/chartex\" xmlns:cx7=\"http://schemas.microsoft.com/office/drawing/2016/5/13/chartex\" xmlns:cx8=\"http://schemas.microsoft.com/office/drawing/2016/5/14/chartex\" xmlns:mc=\"http://schemas.openxmlformats.org/markup-compatibility/2006\" xmlns:aink=\"http://schemas.microsoft.com/office/drawing/2016/ink\" xmlns:am3d=\"http://schemas.microsoft.com/office/drawing/2017/model3d\" xmlns:o=\"urn:schemas-microsoft-com:office:office\" xmlns:oel=\"http://schemas.microsoft.com/office/2019/extlst\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" xmlns:m=\"http://schemas.openxmlformats.org/officeDocument/2006/math\" xmlns:v=\"urn:schemas-microsoft-com:vml\" xmlns:wp14=\"http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing\" xmlns:wp=\"http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing\" xmlns:w10=\"urn:schemas-microsoft-com:office:word\" xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\" xmlns:w14=\"http://schemas.microsoft.com/office/word/2010/wordml\" xmlns:w15=\"http://schemas.microsoft.com/office/word/2012/wordml\" xmlns:w16cex=\"http://schemas.microsoft.com/office/word/2018/wordml/cex\" xmlns:w16cid=\"http://schemas.microsoft.com/office/word/2016/wordml/cid\" xmlns:w16=\"http://schemas.microsoft.com/office/word/2018/wordml\" xmlns:w16sdtdh=\"http://schemas.microsoft.com/office/word/2020/wordml/sdtdatahash\" xmlns:w16se=\"http://schemas.microsoft.com/office/word/2015/wordml/symex\" xmlns:wpg=\"http://schemas.microsoft.com/office/word/2010/wordprocessingGroup\" xmlns:wpi=\"http://schemas.microsoft.com/office/word/2010/wordprocessingInk\" xmlns:wne=\"http://schemas.microsoft.com/office/word/2006/wordml\" xmlns:wps=\"http://schemas.microsoft.com/office/word/2010/wordprocessingShape\"> </a:SrgbClr>"

                    //+
                    // "                        <a:schemeClr val=\"accent1\">\n" +
                    // "                          <a:lumMod val=\"60000\"/>\n" +
                    // "                          <a:lumOff val=\"40000\"/>\n" +
                    // "                        </a:schemeClr>\n" +
                    //"                      </a:solidFill>\n" +
                    //"                    </a:ln>\n"
                    ;
          String xml=  picture.getCTPicture().getSpPr().getLn().getSolidFill().xmlText();
            picture.getCTPicture().getSpPr().getLn().getSolidFill()
                    .set(XmlToken.Factory.parse(solidFillStr));


        }
            if ((null == picture.getCTPicture().getSpPr().getLn().getSolidFill())) {
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
            // setVal(new byte[]{(byte) (51^ 0xff),(byte)(165^ 0xff),(byte)(255^ 0xff)});

        //picture.getCTPicture().getSpPr().getLn().getSolidFill().set();

            picture.getCTPicture().getSpPr().getXfrm().getExt().setCx((long)( width * 360000));  // 设置宽度
//            picture.getCTPicture().getSpPr().getXfrm().getExt().setCy((long) 8 * 360000);  // 设置宽度
//                            System.out.println("是否等于："+  (picture.getCTPicture().getSpPr().getSolidFill()==
//

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
                        PoiSetWidth(picture,pic,14.3f);
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
        catch (IOException e)
        {
            e.printStackTrace();
        } catch (XmlException e) {
            throw new RuntimeException(e);
        }
    }
    public static void main(String[] args) throws InvalidFormatException
    {
        try
        {
            // 读取Word文档
            String srcFile="D:/test_word/realFile2.docx";
            String desFile="D:/test_word/modified_document.docx";
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

//                XWPFParagraph currenPara= desDoc.createParagraph();
//                currenPara=paragraph;
                //paragraph.setBorderTop(Borders.valueOf(3));
                runCount=1;
                System.out.println(para+" para.toString():    :"+paragraph.toString());
                System.out.println(para+" para  getStyle:"+paragraph.getStyle());
                System.out.println(para+" para.getText   :"+paragraph.getText());
                int pos = 0;
                int pictureAmount=0;
                // 获取段落中的所有Run
                for (XWPFRun run : paragraph.getRuns())
                {
                    pictureAmount += run.getEmbeddedPictures().size();
                }


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
                        float desiredWidthCm = 14.3f;
                        priPics.add(picture);
                        PoiSetWidth(picture,pic,(float)(desiredWidthCm/pictureAmount));
                        CTInline inline= run.getCTR().getDrawingList().get(pic-1).getInlineArray(pic-1);
                        inline.getExtent().setCx((long)(desiredWidthCm/pictureAmount*360000));
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
