package org.example;

import org.apache.poi.xwpf.usermodel.*;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.*;

import javax.annotation.PostConstruct;
import javax.xml.bind.annotation.adapters.HexBinaryAdapter;
import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.math.BigInteger;
import java.time.LocalDateTime;
import java.util.Date;
import java.util.Timer;


public class CreateTestWordUtil {
	private static CreateTestWordUtil createJGWordUtil;
	int numLevel = 0; //编号
	String filePath;
	String filename;
	@PostConstruct
	public void init() {
		createJGWordUtil = this;
	}
 
	/**
	 * @param styles       样式
	 * @param strStyleId   标题id
	 * @param headingLevel 标题级别
	 * @param pointSize    字体大小（/2）
	 * @param hexColor     字体颜色
	 * @param typefaceName 字体名称（默认微软雅黑）
	 */
	public void createHeadingStyle(XWPFStyles styles, String strStyleId,
								   int headingLevel, int pointSize, String hexColor, String typefaceName) {
		//创建样式
		CTStyle ctStyle = CTStyle.Factory.newInstance();
		//设置id
		ctStyle.setStyleId(strStyleId);
 
		CTString styleName = CTString.Factory.newInstance();
		styleName.setVal(strStyleId);
		ctStyle.setName(styleName);
 
		CTDecimalNumber indentNumber = CTDecimalNumber.Factory.newInstance();
		indentNumber.setVal(BigInteger.valueOf(headingLevel));
 
		// 数字越低在格式栏中越突出
		ctStyle.setUiPriority(indentNumber);
 
		CTOnOff onoffnull = CTOnOff.Factory.newInstance();
		ctStyle.setUnhideWhenUsed(onoffnull);
 
		// 样式将显示在“格式”栏中
		ctStyle.setQFormat(onoffnull);
 
		// 样式定义给定级别的标题
		if (headingLevel != 0) {
			CTPPr ppr = CTPPr.Factory.newInstance();
			ppr.setOutlineLvl(indentNumber);
			ctStyle.setPPr((CTPPrGeneral) ppr);
		}
		XWPFStyle style = new XWPFStyle(ctStyle);
 
		CTHpsMeasure size = CTHpsMeasure.Factory.newInstance();
		size.setVal(new BigInteger(String.valueOf(pointSize)));
		CTHpsMeasure size2 = CTHpsMeasure.Factory.newInstance();
		size2.setVal(new BigInteger(String.valueOf(pointSize)));
 
		CTFonts fonts = CTFonts.Factory.newInstance();
		if (typefaceName == null || typefaceName.equals("")) typefaceName = "微软雅黑";
		fonts.setAscii(typefaceName);    //字体
 
		CTRPr rpr = CTRPr.Factory.newInstance();
		rpr.setRFontsArray(new CTFonts[]{fonts});
		//rpr.setSz(size);
		//rpr.setSzCs(size2);    //字体大小
 
		CTColor color = CTColor.Factory.newInstance();
		color.setVal(hexToBytes(hexColor));
		//rpr.setColor(color);    //字体颜色
		style.getCTStyle().setRPr(rpr);
		// is a null op if already defined
 
		style.setType(STStyleType.PARAGRAPH);
		styles.addStyle(style);
 
	}
 
	public void writeWordAQJG() {
		// 文档生成方法
		XWPFDocument doc = new XWPFDocument();
 
		XWPFStyles xwpfStyles = doc.createStyles();
		CTFonts fonts = CTFonts.Factory.newInstance();
		fonts.setAscii("微软雅黑");
		fonts.setEastAsia("微软雅黑");
		fonts.setHAnsi("微软雅黑");
		xwpfStyles.setDefaultFonts(fonts);
		createHeadingStyle(xwpfStyles, "标题 1", 1, 32, "000000", "微软雅黑");
		createHeadingStyle(xwpfStyles, "标题 2", 2, 28, "000000", "微软雅黑");
		createHeadingStyle(xwpfStyles, "正文", 0, 24, "000000", "微软雅黑");
 
		XWPFParagraph xwpfParagraphtop = doc.createParagraph(); // 创建段落
		xwpfParagraphtop.setFontAlignment(2);
		xwpfParagraphtop.setStyle("标题 1");
		XWPFRun xwpfRuntop = xwpfParagraphtop.createRun(); // 创建段落文本
		xwpfRuntop.setText(String.format("标题")); // 设置文本
//		xwpfRuntop.setFontFamily("微软雅黑");
		xwpfRuntop.setBold(true);
		xwpfRuntop.setFontSize(24);
//		xwpfRuntop.addBreak();// 换行
		xwpfRuntop.addTab();
 
		XWPFParagraph xwpfParagraphtop1 = doc.createParagraph(); // 创建段落
		xwpfParagraphtop1.setFontAlignment(3);
		xwpfParagraphtop1.setStyle("正文");
		XWPFRun xwpfRuntop1 = xwpfParagraphtop1.createRun(); // 创建段落文本
		xwpfRuntop1.setText("- abcd"); // 设置文本
//		xwpfRuntop1.setFontSize(12);
//		xwpfRuntop1.addBreak();// 换行
		xwpfRuntop1.addTab();
 
 
		FileOutputStream out = null; // 创建输出流
		try {
			//需要的配置项
			writeItemBGGS(doc);
			writeItemNGDWQK(doc);
			writeItemNGZCQK(doc);
			writeItemDDZLQK(doc);
 
			if (System.getProperty("os.name").toLowerCase().contains("linux")) {
				filePath = "/usr/local/createfile/weekly/";
			} else {
				filePath = "D:\\hian\\createfile\\weekly\\";
			}
 
			filename = LocalDateTime.now().getYear() + "年" + LocalDateTime.now().getMonth().getValue() + "月-";// + .now().getTime();
 
			File file = new File(filePath + filename + ".docx");
			if (!file.exists()) {
				file.getParentFile().mkdirs();
				file.createNewFile();
			}
			out = new FileOutputStream(file);
			doc.write(out);
		} catch (Exception e) {
			e.printStackTrace();
		} finally {
			if (out != null) {
				try {
					doc.close();
					out.close();
 
				} catch (IOException e) {
					e.printStackTrace();
				}
			}
		}
	}
 
 
	public void writeItemBGGS(XWPFDocument doc) {
		//编号等级
		CTAbstractNum cTAbstractNum = CTAbstractNum.Factory.newInstance();
		cTAbstractNum.setAbstractNumId(BigInteger.valueOf(10));
 
		CTLvl cTLvl = cTAbstractNum.addNewLvl();
		cTLvl.addNewNumFmt().setVal(STNumberFormat.DECIMAL);
		cTLvl.addNewLvlText().setVal("%1.");
		cTLvl.addNewStart().setVal(BigInteger.valueOf(1));
 
		XWPFAbstractNum abstractNum = new XWPFAbstractNum(cTAbstractNum);
		XWPFNumbering numbering = doc.createNumbering();
		BigInteger abstractNumID = numbering.addAbstractNum(abstractNum);
		BigInteger numID = numbering.addNum(abstractNumID);
 
		XWPFParagraph xwpfParagraphtext = doc.createParagraph(); // 创建段落
		xwpfParagraphtext.setAlignment(ParagraphAlignment.LEFT);
		xwpfParagraphtext.setNumID(numID);
		xwpfParagraphtext.setStyle("标题 1");
		XWPFRun xwpfRuntext = xwpfParagraphtext.createRun(); // 创建段落文本
		xwpfRuntext.setText("标题");
		xwpfRuntext.setBold(true);
 
		XWPFParagraph xwpfParagraphtext1 = doc.createParagraph(); // 创建段落
		XWPFRun xwpfRuntext1 = xwpfParagraphtext1.createRun(); // 创建段落文本
		xwpfRuntext1.setStyle("正文");
		xwpfRuntext1.setText("abcd");
		xwpfRuntext1.addBreak();// 换行
 
		numLevel++;
	}
 
 
	public void writeItemNGDWQK(XWPFDocument doc) {
		//编号等级
		CTAbstractNum cTAbstractNum = CTAbstractNum.Factory.newInstance();
		cTAbstractNum.setAbstractNumId(BigInteger.valueOf(10));
 
 
		CTLvl cTLvl = cTAbstractNum.addNewLvl();
		cTLvl.addNewNumFmt().setVal(STNumberFormat.DECIMAL);
		cTLvl.addNewLvlText().setVal("%1.");
		cTLvl.addNewStart().setVal(BigInteger.valueOf(1));
 
		XWPFAbstractNum abstractNum = new XWPFAbstractNum(cTAbstractNum);
		XWPFNumbering numbering = doc.createNumbering();
		BigInteger abstractNumID = numbering.addAbstractNum(abstractNum);
		BigInteger numID = numbering.addNum(abstractNumID);
 
		XWPFParagraph xwpfParagraphtext = doc.createParagraph(); // 创建段落
		xwpfParagraphtext.setAlignment(ParagraphAlignment.LEFT);
		xwpfParagraphtext.setNumID(numID);
		xwpfParagraphtext.setStyle("标题 1");
		XWPFRun xwpfRuntext = xwpfParagraphtext.createRun(); // 创建段落文本
		xwpfRuntext.setText("标题");
		xwpfRuntext.setBold(true);
 
		XWPFParagraph xwpfParagraphtext1 = doc.createParagraph(); // 创建段落
		XWPFRun xwpfRuntext1 = xwpfParagraphtext1.createRun(); // 创建段落文本
		xwpfRuntext1.setStyle("正文");
		xwpfRuntext1.setText("abcd");
 
		xwpfRuntext1.addBreak();// 换行
 
 
		numLevel++;
	}
 
 
	public void writeItemNGZCQK(XWPFDocument doc) {
		//编号等级
		CTAbstractNum cTAbstractNum = CTAbstractNum.Factory.newInstance();
		cTAbstractNum.setAbstractNumId(BigInteger.valueOf(10));
 
		CTLvl cTLvl = cTAbstractNum.addNewLvl();
		cTLvl.addNewNumFmt().setVal(STNumberFormat.DECIMAL);
		cTLvl.addNewLvlText().setVal("%1.");
		cTLvl.addNewStart().setVal(BigInteger.valueOf(1));
 
		XWPFAbstractNum abstractNum = new XWPFAbstractNum(cTAbstractNum);
		XWPFNumbering numbering = doc.createNumbering();
		BigInteger abstractNumID = numbering.addAbstractNum(abstractNum);
		BigInteger numID = numbering.addNum(abstractNumID);
 
		XWPFParagraph xwpfParagraphtext = doc.createParagraph(); // 创建段落
		xwpfParagraphtext.setAlignment(ParagraphAlignment.LEFT);
		xwpfParagraphtext.setNumID(numID);
		xwpfParagraphtext.setStyle("标题 1");
		XWPFRun xwpfRuntext = xwpfParagraphtext.createRun(); // 创建段落文本
		xwpfRuntext.setText("标题");
		xwpfRuntext.setBold(true);
 
		numLevel++;
 
		writeItemZCRRQK(doc);
		writeItemGFXZC(doc);
		writeItemXTHGQK(doc);
	}
 
	public void writeItemDDZLQK(XWPFDocument doc) {
		//编号等级
		CTAbstractNum cTAbstractNum = CTAbstractNum.Factory.newInstance();
		cTAbstractNum.setAbstractNumId(BigInteger.valueOf(10));
 
		CTLvl cTLvl = cTAbstractNum.addNewLvl();
		cTLvl.addNewNumFmt().setVal(STNumberFormat.DECIMAL);
		cTLvl.addNewLvlText().setVal("%1.");
		cTLvl.addNewStart().setVal(BigInteger.valueOf(1));
 
		XWPFAbstractNum abstractNum = new XWPFAbstractNum(cTAbstractNum);
		XWPFNumbering numbering = doc.createNumbering();
		BigInteger abstractNumID = numbering.addAbstractNum(abstractNum);
		BigInteger numID = numbering.addNum(abstractNumID);
 
		XWPFParagraph xwpfParagraphtext = doc.createParagraph(); // 创建段落
		xwpfParagraphtext.setAlignment(ParagraphAlignment.LEFT);
		xwpfParagraphtext.setNumID(numID);
		xwpfParagraphtext.setStyle("标题 1");
		XWPFRun xwpfRuntext = xwpfParagraphtext.createRun(); // 创建段落文本
		xwpfRuntext.setText("标题");
		xwpfRuntext.setBold(true);
 
		numLevel++;
		writeItemZLLXTJ(doc);
		writeItemZLXYTJ(doc);
	}
 
	public void writeItemZCRRQK(XWPFDocument doc) {
		//编号等级
		CTAbstractNum cTAbstractNum = CTAbstractNum.Factory.newInstance();
		cTAbstractNum.setAbstractNumId(BigInteger.valueOf(13));
 
		CTLvl cTLvl = cTAbstractNum.addNewLvl();
		cTLvl.addNewNumFmt().setVal(STNumberFormat.DECIMAL);
 
		cTLvl.addNewLvlText().setVal(numLevel + ".%1");
		cTLvl.addNewStart().setVal(BigInteger.valueOf(1));
 
		XWPFAbstractNum abstractNum = new XWPFAbstractNum(cTAbstractNum);
		XWPFNumbering numbering = doc.createNumbering();
		BigInteger abstractNumID = numbering.addAbstractNum(abstractNum);
		BigInteger numID = numbering.addNum(abstractNumID);
 
		XWPFParagraph xwpfParagraphtext = doc.createParagraph(); // 创建标题段落
		xwpfParagraphtext.setNumID(numID);
		xwpfParagraphtext.setStyle("标题 2");
		XWPFRun xwpfRuntext = xwpfParagraphtext.createRun(); // 创建段落文本
		xwpfRuntext.setText("标题");
		xwpfRuntext.setBold(true);
 
		XWPFParagraph xwpfParagraphtext1 = doc.createParagraph(); // 创建段落
		XWPFRun xwpfRuntext1 = xwpfParagraphtext1.createRun(); // 创建段落文本
		xwpfRuntext1.setStyle("正文");
		xwpfRuntext1.setText(String.format("abcd："));
//		xwpfRuntext1.addTab();
//		xwpfRuntext1.addBreak();// 换行
 
	}
 
	public void writeItemGFXZC(XWPFDocument doc) {
		//编号等级
		CTAbstractNum cTAbstractNum = CTAbstractNum.Factory.newInstance();
		cTAbstractNum.setAbstractNumId(BigInteger.valueOf(13));
 
		CTLvl cTLvl = cTAbstractNum.addNewLvl();
		cTLvl.addNewNumFmt().setVal(STNumberFormat.DECIMAL);
 
		cTLvl.addNewLvlText().setVal(numLevel + ".%1");
		cTLvl.addNewStart().setVal(BigInteger.valueOf(1));
 
		XWPFAbstractNum abstractNum = new XWPFAbstractNum(cTAbstractNum);
		XWPFNumbering numbering = doc.createNumbering();
		BigInteger abstractNumID = numbering.addAbstractNum(abstractNum);
		BigInteger numID = numbering.addNum(abstractNumID);
 
		XWPFParagraph xwpfParagraphtext = doc.createParagraph(); // 创建标题段落
		xwpfParagraphtext.setNumID(numID);
		xwpfParagraphtext.setStyle("标题 2");
		XWPFRun xwpfRuntext = xwpfParagraphtext.createRun(); // 创建段落文本
		xwpfRuntext.setText("标题");
		xwpfRuntext.setBold(true);
 
	}
 
	public void writeItemXTHGQK(XWPFDocument doc) {
		//编号等级
		CTAbstractNum cTAbstractNum = CTAbstractNum.Factory.newInstance();
		cTAbstractNum.setAbstractNumId(BigInteger.valueOf(13));
 
		CTLvl cTLvl = cTAbstractNum.addNewLvl();
		cTLvl.addNewNumFmt().setVal(STNumberFormat.DECIMAL);
 
		cTLvl.addNewLvlText().setVal(numLevel + ".%1");
		cTLvl.addNewStart().setVal(BigInteger.valueOf(1));
 
		XWPFAbstractNum abstractNum = new XWPFAbstractNum(cTAbstractNum);
		XWPFNumbering numbering = doc.createNumbering();
		BigInteger abstractNumID = numbering.addAbstractNum(abstractNum);
		BigInteger numID = numbering.addNum(abstractNumID);
 
		XWPFParagraph xwpfParagraphtext = doc.createParagraph(); // 创建标题段落
		xwpfParagraphtext.setNumID(numID);
		xwpfParagraphtext.setStyle("标题 2");
		XWPFRun xwpfRuntext = xwpfParagraphtext.createRun(); // 创建段落文本
		xwpfRuntext.setText("标题");
		xwpfRuntext.setBold(true);
 
	}
 
	public void writeItemZLLXTJ(XWPFDocument doc) {
		//编号等级
		CTAbstractNum cTAbstractNum = CTAbstractNum.Factory.newInstance();
		cTAbstractNum.setAbstractNumId(BigInteger.valueOf(14));
 
		CTLvl cTLvl = cTAbstractNum.addNewLvl();
		cTLvl.addNewNumFmt().setVal(STNumberFormat.DECIMAL);
 
		cTLvl.addNewLvlText().setVal(numLevel + ".%1");
		cTLvl.addNewStart().setVal(BigInteger.valueOf(1));
 
		XWPFAbstractNum abstractNum = new XWPFAbstractNum(cTAbstractNum);
		XWPFNumbering numbering = doc.createNumbering();
		BigInteger abstractNumID = numbering.addAbstractNum(abstractNum);
		BigInteger numID = numbering.addNum(abstractNumID);
 
		XWPFParagraph xwpfParagraphtext = doc.createParagraph(); // 创建标题段落
		xwpfParagraphtext.setNumID(numID);
		xwpfParagraphtext.setStyle("标题 2");
		XWPFRun xwpfRuntext = xwpfParagraphtext.createRun(); // 创建段落文本
		xwpfRuntext.setText("标题");
		xwpfRuntext.setBold(true);
 
	}
 
	public void writeItemZLXYTJ(XWPFDocument doc) {
		//编号等级
		CTAbstractNum cTAbstractNum = CTAbstractNum.Factory.newInstance();
		cTAbstractNum.setAbstractNumId(BigInteger.valueOf(14));
 
		CTLvl cTLvl = cTAbstractNum.addNewLvl();
		cTLvl.addNewNumFmt().setVal(STNumberFormat.DECIMAL);
 
		cTLvl.addNewLvlText().setVal(numLevel + ".%1");
		cTLvl.addNewStart().setVal(BigInteger.valueOf(1));
 
		XWPFAbstractNum abstractNum = new XWPFAbstractNum(cTAbstractNum);
		XWPFNumbering numbering = doc.createNumbering();
		BigInteger abstractNumID = numbering.addAbstractNum(abstractNum);
		BigInteger numID = numbering.addNum(abstractNumID);
 
		XWPFParagraph xwpfParagraphtext = doc.createParagraph(); // 创建标题段落
		xwpfParagraphtext.setNumID(numID);
		xwpfParagraphtext.setStyle("标题 2");
		XWPFRun xwpfRuntext = xwpfParagraphtext.createRun(); // 创建段落文本
		xwpfRuntext.setText("标题");
		xwpfRuntext.setBold(true);
 
		XWPFParagraph xwpfParagraphtext1 = doc.createParagraph(); // 创建段落
		XWPFRun xwpfRuntext1 = xwpfParagraphtext1.createRun(); // 创建段落文本
		xwpfRuntext1.setStyle("正文");
		xwpfRuntext1.setText(String.format("abc"));
 
	}
 
	public static byte[] hexToBytes(String hexString) {
		HexBinaryAdapter adapter = new HexBinaryAdapter();
		byte[] bytes = adapter.unmarshal(hexString);
		return bytes;
	}
 
}