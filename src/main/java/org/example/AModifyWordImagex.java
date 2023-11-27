package org.example;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.util.Units;
import org.apache.poi.xwpf.usermodel.*;
import org.apache.xmlbeans.XmlException;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTStyle;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTStyles;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;


public class AModifyWordImagex {
    /**
     * 查找文档样式值
     * @param document 文档类
     * @param styleName 样式名称
     * @return 样式值
     * @throws IOException
     * @throws XmlException
     */
    public static String getStyleValue(XWPFDocument document, String styleName) throws IOException, XmlException {
        if (styleName == null || styleName.length() == 0)
        {
            return null;
        }
        CTStyles styles = document.getStyle();
        CTStyle[] styleArray = styles.getStyleArray();
        for (CTStyle style : styleArray) {
            if (style.getName().getVal().equals(styleName)) {
                return style.getStyleId();
            }
        }
        return null;
    }
    public static void main(String[] args)
    {
        try {
            // 读取Word文档
            String oldSrcFile="D:/test_word/new.docx";
            String srcFile="D:/test_word/primaryFile.docx";
            String desFile="D:/test_word/modified_document.docx";
            FileInputStream fis = new FileInputStream(srcFile);
            FileOutputStream fos = new FileOutputStream(desFile);
            FileInputStream myfis = new FileInputStream(desFile);
            FileOutputStream myfos = new FileOutputStream(desFile);
            CustomXWPFDocument document = new CustomXWPFDocument(fis);
            CustomXWPFDocument myDocument = new CustomXWPFDocument();
            XWPFDocument testDocument = new XWPFDocument();

            int para=1;
            int  pc=1;
            int  pic=1;
            List<XWPFPicture> priPics=new ArrayList<XWPFPicture>() ;
            // 获取文档中的所有段落
            for (XWPFParagraph paragraph : document.getParagraphs())
            {
                pc=1;
               XWPFParagraph currenPara= myDocument.createParagraph();
               currenPara.setStyle(paragraph.getStyle());
               currenPara.setAlignment(paragraph.getAlignment());
               currenPara.setFontAlignment(paragraph.getFontAlignment());
               currenPara.setFirstLineIndent(paragraph.getFirstLineIndent());
               currenPara.setIndentFromLeft(paragraph.getIndentFromLeft());
                paragraph.setBorderTop(Borders.valueOf(3));
                System.out.println(para+" para.toString():    :"+paragraph.toString());
                System.out.println(para+" para  getStyle:"+paragraph.getStyle());
                System.out.println(para+++" para.getText   :"+paragraph.getText());
                int pos = 0;
                // 获取段落中的所有Run
                for (XWPFRun run : paragraph.getRuns())
                {
                    XWPFRun currenRun=currenPara.createRun();
                    currenRun.setText(run.text());
                    currenRun.setStyle(run.getStyle());
                    currenRun.setVerticalAlignment(String.valueOf(run.getVerticalAlignment()));
                    //currenRun.setFontSize(run.getFontSize());
                    currenRun.setFontFamily(run.getFontFamily());
                    currenRun.setBold(run.isBold());
//                    currenRun=run;
                    pic=1;
                    System.out.println(pc+" run.text()   :"+run.text());
                    //run.getEmbeddedPictures().removeAll(priPics);
                    // 获取Run中的所有Embedded Pictures

                    for (XWPFPicture picture : run.getEmbeddedPictures())
                    {

                        priPics.add(picture);
                       // picture.getPictureData().
                        System.out.println("pic"+pic+"  picture.getCTPicture()  :"+picture.getCTPicture());
                        System.out.println("pic"+pic+"  picture.getDepth()  :"+picture.getDepth());
                        System.out.println("pic"+pic+"  picture.getWidth()  :"+picture.getWidth());
                        System.out.println("pic"+pic+"  picture.getPictureData()  :"+picture.getPictureData());
                        System.out.println("pic"+pic+"  picture.getDescription()  :"+picture.getDescription());
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
                        byte[] picBytes = pictureData.getData();
//                        Pictures.ofBytes (picBytes,)
//                                .size(100, 120).create();
                        double hight=  Units.pixelToPoints(picture.getDepth()) * 2.54 / 1440.0 * 20.0;
                       // Pictures.ofBytes(bytes).sizeInCm(14.3,hight).create();
                       // Pictures.ofBytes(bytes).sizeInCm(14.3,picture.getDepth());
//                        // 修改图片的边框
                     // if(!picture.getCTPicture().getSpPr().isSetSolidFill())
                        {

                            String bid= myDocument.addPictureData(picture.getPictureData().getData(),picture.getPictureData().getPictureType());
                          //String bid=  myDocument.addPictureData(new FileInputStream("D:\\wechatfile.jpg"),picture.getPictureData().getPictureType());
                            myDocument.createPicture(paragraph,run,picture,
                                    picture.getCTPicture().
                                            getBlipFill().getBlip().getEmbed(),
                                    (int)picture.getCTPicture().getNvPicPr().getCNvPr().getId(),
                                    (long)(14.3*360000),
                                    (long)(8*360000),
//                                    (int)picture.getWidth(),
//                                    (int)picture.getDepth(),
                                    19500,
                                    "FF0000",bid);
                            //FF0000  33A5FF
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
//                            if ((null == picture.getCTPicture().getSpPr().getLn()))
//                            {
//                                picture.getCTPicture().getSpPr().addNewLn().setW(19000);
//                            } else
//                            {
//                                picture.getCTPicture().getSpPr().getLn().setW(19000);
//                            }
//                            if ((null == picture.getCTPicture().getSpPr().getLn().getSolidFill()))
//                            {
//                               //picture.getCTPicture().getSpPr().getLn().
//                                picture.getCTPicture().getSpPr().getLn().addNewSolidFill();
//                            }
//                            else
//                            {
//                                //33A5FF: 51 165 255
//                                if ((null == picture.getCTPicture().getSpPr().getLn().getSolidFill().getSrgbClr()))
//                                {
//                                    picture.getCTPicture().getSpPr().getLn().getSolidFill().addNewSrgbClr();
//                                }
//                                picture.getCTPicture().getSpPr().getLn().getSolidFill().getSrgbClr().
//                                        setVal(new byte[]{(byte) (51),(byte)(165),(byte)(255)});
//                                       // setVal(new byte[]{(byte) (51^ 0xff),(byte)(165^ 0xff),(byte)(255^ 0xff)});
//
//                            }

//                            System.out.println("是否等于："+  (picture.getCTPicture().getSpPr().getSolidFill()==
//                                    picture.getCTPicture().getSpPr().getLn().getSolidFill()));
                            // 设置边框宽度，单位为20分之1磅
                          //  picture.getCTPicture().getSpPr().getLn().addNewSolidFill().setSrgbClr(ctsRgbColor);
                           // picture.getCTPicture().getSpPr().getLn().addNewSolidFill().addNewSrgbClr().setVal(new byte[]{0,0,0});
//                            CTSRgbColor ctsRgbColor = CTSRgbColor.Factory.newInstance();
//                            ctsRgbColor.addNewBlue();
//                            picture.getCTPicture().getSpPr().addNewSolidFill().setSrgbClr(ctsRgbColor);
//                            // 设置边框颜色
////                       // picture.getCTPicture().getSpPr().addNewLn().setAlgn(STPenAlignment.Enum.forInt(1));  // 设置边框居中
                            float desiredWidthCm = 14.3f;
//                           double heightCm= Units.pixelToPoints(picture.getDepth()) * 2.54 / 1440.0 * 20.0;
                            int desiredWidthEMU = (int) (Units.toEMU(desiredWidthCm * 2.54f));  // 将厘米转换为EMU
                            //picture.getCTPicture().getSpPr().getXfrm().getExt().setCx((long) desiredWidthCm * 360000);  // 设置宽度
                           // picture.getCTPicture().getSpPr().getXfrm().getExt().setCy((long) picture.getDepth());  // 设置宽度
                            System.out.println("chenge pic"+pic+"  picture.getCTPicture()  :"+picture.getCTPicture());
                            System.out.println("chenge pic"+pic+"  picture.getDepth()  :"+picture.getDepth());
                            System.out.println("chenge pic"+pic+"  picture.getWidth()  :"+picture.getWidth());
                            System.out.println("chenge pic"+pic+"  picture.getPictureData()  :"+picture.getPictureData());
                            System.out.println("chenge pic"+pic+++"  picture.getDescription()  :"+picture.getDescription());

                        }
                    }
                    //run.getEmbeddedPictures().removeAll(priPics);
                }
            }

            // 保存修改后的Word文档
          // FileOutputStream fos = new FileOutputStream("D:/test_word/modified_document.docx");
            myDocument.write(fos);
            fos.flush();
            fos.close();
            fis.close();
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
//           delFis.close();
//            delFos.close();
            //myfos.flush();
            // 关闭资源
            //fis.close();
            //fos.close();
            //myfis.close();
            //myfos.close();

        } catch (IOException e)
        {
            e.printStackTrace();
        } catch (InvalidFormatException e) {
            throw new RuntimeException(e);
        }

    }
}
