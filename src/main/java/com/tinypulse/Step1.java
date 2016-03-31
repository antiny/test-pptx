package com.tinypulse;

import org.apache.poi.xslf.usermodel.*;

import java.io.FileInputStream;

/**
 * Created by antran on 3/29/16.
 */
public class Step1 {
    public static void main(String[] args) throws Exception {

//        XMLSlideShow ppt = new XMLSlideShow(new FileInputStream("/Users/antran/workspace/pptx-assets/sample-shape.pptx"));
        XMLSlideShow ppt = new XMLSlideShow(new FileInputStream("/Users/antran/workspace/pptx-assets/template1.pptx"));


        // first see what slide layouts are available by default
        System.out.println("Available slide layouts:");
        for(XSLFSlideMaster master : ppt.getSlideMasters()){
            System.out.println("Master:" + master.toString());
            for(XSLFSlideLayout layout : master.getSlideLayouts()){
                System.out.println(layout.getType());
            }
        }


        for(XSLFSlide slide : ppt.getSlides()){
            System.out.println("Title: " + slide.getTitle() + " =================");

            for(XSLFShape shape : slide.getShapes()){
                if(shape instanceof XSLFTextShape) {
                    XSLFTextShape tsh = (XSLFTextShape)shape;
                    System.out.println("  shape: " + tsh);
//                    tsh.get
                    for(XSLFTextParagraph p : tsh){
                        System.out.println("Paragraph level: " + p.getIndentLevel());
                        for(XSLFTextRun r : p){
                            System.out.println(r.getRawText());
                            System.out.println("  bold: " + r.isBold());
                            System.out.println("  italic: " + r.isItalic());
                            System.out.println("  underline: " + r.isUnderlined());
                            System.out.println("  font.family: " + r.getFontFamily());
                            System.out.println("  font.size: " + r.getFontSize());
                            System.out.println("  font.color: " + r.getFontColor());
                        }
                    }
                }
            }
        }
    }
}
