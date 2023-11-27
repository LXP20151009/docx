package org.example;

import com.deepoove.poi.data.Pictures;
import org.apache.poi.xwpf.usermodel.*;
import org.apache.poi.util.Units;
import org.openxmlformats.schemas.drawingml.x2006.main.STPenAlignment;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

public class ModifyWordImage {

    public static void main(String[] args) {
        try {
            // 读取Word文档
            FileInputStream fis = new FileInputStream("D:/test_word/test1.docx");
            XWPFDocument document = new XWPFDocument(fis);
            int  pc=1;
            // 获取文档中的所有段落
            for (XWPFParagraph paragraph : document.getParagraphs())
            {
                // 获取段落中的所有Run
                for (XWPFRun run : paragraph.getRuns())
                {

                    System.out.println(pc+++"    :"+run.text());
                    // 获取Run中的所有Embedded Pictures
                    for (XWPFPicture picture : run.getEmbeddedPictures())
                    {    // 获取图片对象
                        XWPFPictureData pictureData = picture.getPictureData();
                        byte[] bytes = pictureData.getData();
                        double hight=  Units.pixelToPoints(picture.getDepth()) * 2.54 / 1440.0 * 20.0;
                        //Pictures.ofBytes(bytes).sizeInCm(14.3,hight).create();
                       // Pictures.ofBytes(bytes).sizeInCm(14.3,picture.getDepth());
                        // 修改图片的边框
                        //picture.getCTPicture().getSpPr().addNewLn().setW(15);  // 设置边框宽度，单位为20分之1磅
                       // picture.getCTPicture().getSpPr().addNewLn().setAlgn(STPenAlignment.Enum.forInt(1));  // 设置边框居中
                        float desiredWidthCm = 14.3f;
                        double heightCm= Units.pixelToPoints(picture.getDepth()) * 2.54 / 1440.0 * 20.0;
                        int desiredWidthEMU = (int) (Units.toEMU(desiredWidthCm * 2.54f));  // 将厘米转换为EMU
                        picture.getCTPicture().getSpPr().getXfrm().getExt().setCx(desiredWidthEMU);  // 设置宽度

                    }
                }
            }

            // 保存修改后的Word文档
            FileOutputStream fos = new FileOutputStream("D:/test_word/modified_document.docx");
            document.write(fos);
            fos.flush();
            // 关闭资源
            fis.close();
            fos.close();

        } catch (IOException e) {
            e.printStackTrace();
        }
    }
}
