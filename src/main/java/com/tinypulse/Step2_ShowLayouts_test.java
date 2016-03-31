package com.tinypulse;

import org.apache.poi.xslf.usermodel.*;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.util.List;

/**
 * Created by antran on 3/29/16.
 */
public class Step2_ShowLayouts_test {
    public static void main(String[] args) throws Exception{
        XMLSlideShow ppt = new XMLSlideShow(new FileInputStream("/Users/antran/workspace/pptx-assets/template1.pptx"));

        // title slide
        List<XSLFSlide> slides = ppt.getSlides();

        final XSLFSlide introSlide = slides.get(0);
        introSlide.getPlaceholder(0).setText("TINYpusle Engagement Report");
        introSlide.getPlaceholder(1).setText("April 2016");

        final XSLFSlide summarySlide = slides.get(1);
        summarySlide.getPlaceholder(0).setText("Do you feel that your manager has clearly defined your roles and responsibilities and how it contributes to the success of the organization?");


//        // title and content
//        XSLFSlideLayout titleBodyLayout = defaultMaster.getLayout(SlideLayout.TITLE_AND_CONTENT);
//        XSLFSlide slide2 = ppt.createSlide(titleBodyLayout);
//
//        XSLFTextShape title2 = slide2.getPlaceholder(0);
//        title2.setText("Second Title");
//
//        XSLFTextShape body2 = slide2.getPlaceholder(1);
//        body2.clearText(); // unset any existing text
//        body2.addNewTextParagraph().addNewTextRun().setText("First paragraph");
//        body2.addNewTextParagraph().addNewTextRun().setText("Second paragraph");
//        body2.addNewTextParagraph().addNewTextRun().setText("Third paragraph");



        FileOutputStream out = new FileOutputStream("step2.pptx");
        ppt.write(out);
        out.close();

    }
}
