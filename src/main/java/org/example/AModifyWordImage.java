package org.example;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.util.Units;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFPicture;
import org.apache.poi.xwpf.usermodel.XWPFPictureData;
import org.apache.poi.xwpf.usermodel.XWPFRun;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;


public class AModifyWordImage {

    public static void PoiSetWidth(XWPFPicture picture ,int pic ) {
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
        double hight = Units.pixelToPoints(picture.getDepth()) * 2.54 / 1440.0 * 20.0;
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
            if ((null == picture.getCTPicture().getSpPr().getLn())) {
                picture.getCTPicture().getSpPr().addNewLn().setW(19000);
            } else {
                picture.getCTPicture().getSpPr().getLn().setW(19000);
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


            if (null != picture.getCTPicture().getSpPr().getLn().getSolidFill().getSchemeClr()) {
                picture.getCTPicture().getSpPr().getLn().getSolidFill().getSchemeClr().set(null);
            }
            picture.getCTPicture().getSpPr().getXfrm().getExt().setCx((long) 14.3 * 360000);  // 设置宽度
            picture.getCTPicture().getSpPr().getXfrm().getExt().setCy((long) 8 * 360000);  // 设置宽度
//                            System.out.println("是否等于："+  (picture.getCTPicture().getSpPr().getSolidFill()==
//

    }
    public static void main(String[] args) throws InvalidFormatException {
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
                        PoiSetWidth(picture,pic);
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
        }
    }
}
