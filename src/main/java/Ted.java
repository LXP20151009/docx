import com.deepoove.poi.XWPFTemplate;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.HashMap;
import java.util.Map;

public class Ted {
    
    public static void main(String[] args) throws IOException {
        System.out.println("1123");
        //要写入模板的数据
        Map<String,Object> exampleData = new HashMap<>();
        exampleData.put("username","admin");
        exampleData.put("password","123456");
        FileInputStream inputStream = new FileInputStream("D:/test_word/example1.docx");
        XWPFTemplate template = XWPFTemplate.compile(inputStream).render(exampleData);
        //文件输出流
        FileOutputStream out = new FileOutputStream("D:/test_word/example1.docx");
        template.write(out);
        try {
            out.flush();
        } catch (IOException e) {
            throw new RuntimeException(e);
        }
        try {
            out.close();
        } catch (IOException e) {
            throw new RuntimeException(e);
        }
        template.close();
    }
}
