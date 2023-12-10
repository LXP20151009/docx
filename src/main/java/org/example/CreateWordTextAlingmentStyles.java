package org.example;

import java.io.*;

import org.apache.poi.xwpf.usermodel.*;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.*;
import java.lang.reflect.Field;

public class CreateWordTextAlingmentStyles {

 private static XWPFStyle createNamedStyle(XWPFStyles styles, STStyleType.Enum styleType, String styleId)
 {
  if (styles == null || styleId == null) return null;
  XWPFStyle style = styles.getStyle(styleId);
  if (style == null)
  {
   CTStyle ctStyle = CTStyle.Factory.newInstance();
   ctStyle.addNewName().setVal(styleId);
   ctStyle.setCustomStyle("");
   style = new XWPFStyle(ctStyle, styles);
   style.setType(styleType);
   style.setStyleId(styleId);
   styles.addStyle(style);
  }
  return style;
 }

 private static void applyTextAlignment(XWPFStyle style, TextAlignment value) throws Exception {
  if (style == null || value == null) return;

  Field _ctStyles = XWPFStyles.class.getDeclaredField("ctStyles");
  _ctStyles.setAccessible(true);
  CTStyles ctStyles = (CTStyles)_ctStyles.get(style.getStyles());

  for (CTStyle ctStyle : ctStyles.getStyleList()) {
   if (ctStyle.getStyleId().equals(style.getStyleId())) {
    CTPPr ppr = (CTPPr) ctStyle.getPPr();
    if (ppr == null) ppr = (CTPPr) ctStyle.addNewPPr();
    CTTextAlignment ctTextAlignment = ppr.getTextAlignment(); 
    if (ctTextAlignment == null) ctTextAlignment = ppr.addNewTextAlignment();
    if (value == TextAlignment.AUTO) {
     ctTextAlignment.setVal(STTextAlignment.AUTO);
    } else if (value == TextAlignment.BASELINE) {
     ctTextAlignment.setVal(STTextAlignment.BASELINE);
    } else if (value == TextAlignment.BOTTOM) {
     ctTextAlignment.setVal(STTextAlignment.BOTTOM);
    } else if (value == TextAlignment.CENTER) {
     ctTextAlignment.setVal(STTextAlignment.CENTER);
    } else if (value == TextAlignment.TOP) {
     ctTextAlignment.setVal(STTextAlignment.TOP);
    }
    style.setStyle(ctStyle);
   }
  }
 }

 private static void applyJustification(XWPFStyle style, ParagraphAlignment value) throws Exception {
  if (style == null || value == null) return;

  Field _ctStyles = XWPFStyles.class.getDeclaredField("ctStyles");
  _ctStyles.setAccessible(true);
  CTStyles ctStyles = (CTStyles)_ctStyles.get(style.getStyles());

  for (CTStyle ctStyle : ctStyles.getStyleList()) {
   if (ctStyle.getStyleId().equals(style.getStyleId())) {
    CTPPr ppr = (CTPPr) ctStyle.getPPr(); if (ppr == null) ppr = (CTPPr) ctStyle.addNewPPr();
    CTJc jc = ppr.getJc(); if (jc == null) jc = ppr.addNewJc();
    if (value == ParagraphAlignment.BOTH) {
     jc.setVal(STJc.BOTH);
    } else if (value == ParagraphAlignment.CENTER) {
     jc.setVal(STJc.CENTER);
    } else if (value == ParagraphAlignment.DISTRIBUTE) {
     jc.setVal(STJc.DISTRIBUTE);
    } else if (value == ParagraphAlignment.HIGH_KASHIDA) {
     jc.setVal(STJc.HIGH_KASHIDA);
    } else if (value == ParagraphAlignment.LEFT) {
     jc.setVal(STJc.LEFT);
    } else if (value == ParagraphAlignment.LOW_KASHIDA) {
     jc.setVal(STJc.LOW_KASHIDA);
    } else if (value == ParagraphAlignment.MEDIUM_KASHIDA) {
     jc.setVal(STJc.MEDIUM_KASHIDA);
    } else if (value == ParagraphAlignment.NUM_TAB) {
     jc.setVal(STJc.NUM_TAB);
    } else if (value == ParagraphAlignment.RIGHT) {
     jc.setVal(STJc.RIGHT);
    } else if (value == ParagraphAlignment.THAI_DISTRIBUTE) {
     jc.setVal(STJc.THAI_DISTRIBUTE);
    }
    style.setStyle(ctStyle);
   }
  }
 }

 public static void main(String[] args) throws Exception {

  XWPFDocument document = new XWPFDocument();
  XWPFParagraph paragraph = null;
  XWPFRun run = null;

  XWPFStyles styles = document.createStyles();

  XWPFStyle style = createNamedStyle(styles, STStyleType.PARAGRAPH, "TextAlignmentAUTO");
  applyTextAlignment(style, TextAlignment.AUTO);
  paragraph = document.createParagraph();
  paragraph.setStyle(style.getStyleId());
  run = paragraph.createRun();
  run.setText("TextAlignment.AUTO");
  run.setFontSize(8);
  run = paragraph.createRun();
  run.setText("Bigger text");
  run.setFontSize(30);

  style = createNamedStyle(styles, STStyleType.PARAGRAPH, "TextAlignmentBASELINECentered");
  applyJustification(style, ParagraphAlignment.CENTER);
  applyTextAlignment(style, TextAlignment.BASELINE);
  paragraph = document.createParagraph();
  paragraph.setStyle(style.getStyleId());
  run = paragraph.createRun();
  run.setText("TextAlignment.BASELINE");
  run.setFontSize(8);
  run = paragraph.createRun();
  run.setText("Bigger text");
  run.setFontSize(30);

  style = createNamedStyle(styles, STStyleType.PARAGRAPH, "TextAlignmentBOTTOMRight");
  applyJustification(style, ParagraphAlignment.RIGHT);
  applyTextAlignment(style, TextAlignment.BOTTOM);
  paragraph = document.createParagraph();
  paragraph.setStyle(style.getStyleId());
  run = paragraph.createRun();
  run.setText("TextAlignment.BOTTOM");
  run.setFontSize(8);
  run = paragraph.createRun();
  run.setText("Bigger text");
  run.setFontSize(30);

  style = createNamedStyle(styles, STStyleType.PARAGRAPH, "TextAlignmentCENTERBoth");
  applyJustification(style, ParagraphAlignment.BOTH);
  applyTextAlignment(style, TextAlignment.CENTER);
  paragraph = document.createParagraph();
  paragraph.setStyle(style.getStyleId());
  run = paragraph.createRun();
  run.setText("TextAlignment.CENTER");
  run.setFontSize(8);
  run = paragraph.createRun();
  run.setText("Bigger text");
  run.setFontSize(30);

  style = createNamedStyle(styles, STStyleType.PARAGRAPH, "TextAlignmentTOPLeft");
  applyJustification(style, ParagraphAlignment.LEFT);
  applyTextAlignment(style, TextAlignment.TOP);
  paragraph = document.createParagraph();
  paragraph.setStyle(style.getStyleId());
  run = paragraph.createRun();
  run.setText("TextAlignment.TOP");
  run.setFontSize(8);
  run = paragraph.createRun();
  run.setText("Bigger text");
  run.setFontSize(30);

  FileOutputStream out = new FileOutputStream("./CreateWordTextAlingmentStyles.docx");
  document.write(out);
  out.close();
  document.close();

 }
}