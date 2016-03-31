package com.tinypulse;

import org.apache.poi.POIXMLDocumentPart;
import org.apache.poi.xslf.usermodel.XMLSlideShow;
import org.apache.poi.xslf.usermodel.XSLFChart;
import org.apache.poi.xslf.usermodel.XSLFSlide;
import org.apache.poi.xslf.usermodel.XSLFTextShape;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTChart;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTPlotArea;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.util.List;

/**
 * Created by antran on 3/29/16.
 */
public class TestReport {
    public static void main(String[] args) throws Exception{
        XMLSlideShow ppt = new XMLSlideShow(new FileInputStream("/Users/antran/workspace/pptx-assets/Presentation3.pptx"));

        // title slide
        List<XSLFSlide> slides = ppt.getSlides();

        final XSLFSlide introSlide = slides.get(0);
        final XSLFTextShape[] placeholders = introSlide.getPlaceholders();
        introSlide.getPlaceholder(0).setText("Engagement Report");
        introSlide.getPlaceholder(1).setText("April 2016");

        // find chart in the slide
        final XSLFSlide slide = slides.get(1);
        XSLFChart chart = null;
        for(POIXMLDocumentPart part : slide.getRelations()){
            if(part instanceof XSLFChart){
                chart = (XSLFChart) part;
                break;
            }
        }

        if(chart == null) throw new IllegalStateException("chart not found in the template");

        // embedded Excel workbook that holds the chart data
        POIXMLDocumentPart xlsPart = chart.getRelations().get(0);
        XSSFWorkbook wb = new XSSFWorkbook();

        CTChart ctChart = chart.getCTChart();
        CTPlotArea plotArea = ctChart.getPlotArea();
        


//        final XSLFSlide summarySlide = slides.get(1);
//        summarySlide.getPlaceholder(0).setText("Do you feel that your manager has clearly defined your roles and responsibilities and how it contributes to the success of the organization?");


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

        FileOutputStream out = new FileOutputStream("engagement-report.pptx");
        ppt.write(out);
        out.close();

    }
}
