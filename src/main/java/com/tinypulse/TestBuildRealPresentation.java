package com.tinypulse;

import org.apache.poi.sl.usermodel.PictureData;
import org.apache.poi.util.IOUtils;
import org.apache.poi.xslf.usermodel.*;

import java.awt.geom.Rectangle2D;
import java.io.*;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

public class TestBuildRealPresentation {
    public static void main(String[] args) throws IOException {
        XMLSlideShow ppt = new XMLSlideShow(new FileInputStream("src/main/resources/template1.pptx"));
        final List<XSLFSlide> slides = ppt.getSlides();

        final XSLFSlide introSlide = slides.get(0);
        populateIntroSlide(introSlide);

        final XSLFSlide summarySlide = slides.get(1);
        populateSummarySlide(summarySlide);

        final XSLFSlide textSlide = slides.get(2);
//        debugShapes(textSlide);
        populateHighlightSlide(textSlide);

        save(ppt, "engagement-report.pptx");
    }

    private static Map<String, XSLFTextShape> placeholdersByName(XSLFSlide textSlide) {
        final XSLFTextShape[] placeholders = textSlide.getPlaceholders();

        final Map<String, XSLFTextShape> placeholdersByName = new HashMap<String, XSLFTextShape>();
        for (int i = 0; i < placeholders.length; i++) {
            placeholdersByName.put(placeholders[i].getShapeName(), placeholders[i]);
        }

        return placeholdersByName;
    }

    private static void populateHighlightSlide(XSLFSlide textSlide) {
        final Map<String, XSLFTextShape> placeholdersByName = placeholdersByName(textSlide);
        placeholdersByName.get("date").setText("Mar 27th, 2016");

        populateTrainingSection(placeholdersByName);
        populateManagementSection(placeholdersByName);
        populateBenefitSection(placeholdersByName);
        populateTeamSection(placeholdersByName);
    }

    private static void populateTeamSection(Map<String, XSLFTextShape> placeholdersByName) {
        placeholdersByName.get("team_res_count").setText("36 responses");

        placeholdersByName.get("team_res1_score").setText("10");
        placeholdersByName.get("team_res1_text").setText("\"We were recently purchased by Microsoft. I feel so blessed that the 365 team has made us feel at home from day one\"");

        placeholdersByName.get("team_res2_score").setText("9");
        placeholdersByName.get("team_res2_text").setText("\"The new additions have really lightened the load. Good coworkers makes work so much better\"");

        placeholdersByName.get("team_res3_score").setText("9");
        placeholdersByName.get("team_res3_text").setText("\"I’m grateful for the flexibility to switch teams here. Variety is the spice of life, and I feel more inclined to stay with the company now that I have expos...\"");
    }

    private static void populateBenefitSection(Map<String, XSLFTextShape> placeholdersByName) {
        placeholdersByName.get("benefit_res_count").setText("36 responses | 10 Virtual Suggestions");

        placeholdersByName.get("benefit_res1_score").setText("10");
        placeholdersByName.get("benefit_res1_text").setText("\"I think it’s awesome the way the company has been focused on wellness and balance. I’ve been enjoying my classes at the Pro Club\"");

        placeholdersByName.get("benefit_res2_score").setText("9");
        placeholdersByName.get("benefit_res2_text").setText("\"I’m so appreciative of my team’s support during my eight weeks of maternity leave.\"");

        placeholdersByName.get("benefit_res3_score").setText("9");
        placeholdersByName.get("benefit_res3_text").setText("\"The new John Howie restaurant on the Redmond campus is a phenomenal deal at $15.\"");
    }

    private static void populateManagementSection(Map<String, XSLFTextShape> placeholdersByName) {
        placeholdersByName.get("man_res_count").setText("23 responses");

        placeholdersByName.get("man_res1_score").setText("9");
        placeholdersByName.get("man_res1_text").setText("\"Our management team is very open & respectful. I feel like recent changes have made it very easy for me to voice my opinion \"");

        placeholdersByName.get("man_res2_score").setText("8");
        placeholdersByName.get("man_res2_text").setText("\"Brenna has been a huge help in my onboarding. Lucky to have such a generous manager!\"");

        placeholdersByName.get("man_res3_score").setText("8");
        placeholdersByName.get("man_res3_text").setText("\"It’s not often you find a manager who is truly self-critical. Because of Jim, our team is constantly learning and improving.\"");
    }

    private static void populateTrainingSection(Map<String, XSLFTextShape> placeholdersByName) {
        placeholdersByName.get("training_res_count").setText("38 responses | 2 Virtual Suggestions");

        placeholdersByName.get("training_res1_score").setText("10");
        placeholdersByName.get("training_res1_text").setText("\"Love the new onboarding programs.  I feel my new team members have a better sense about our customer base\"");

        placeholdersByName.get("training_res2_score").setText("9");
        placeholdersByName.get("training_res2_text").setText("\"I think it’s great management has committed to the annual education fund.  I can’t wait to attend the upcoming MLDS Conference\"");

        placeholdersByName.get("training_res3_score").setText("9");
        placeholdersByName.get("training_res3_text").setText("\"It’s awesome that our management team is investing in Precision Questioning training.  This will definitely help us get honest feedback as we evaluate Product.\"");
    }

    private static void populateSummarySlide(XSLFSlide summarySlide) throws IOException {
        final Map<String, XSLFTextShape> placeholdersByName = placeholdersByName(summarySlide);

        populateSummaryText(placeholdersByName);
        populateSummaryImages(summarySlide);
    }

    private static void populateSummaryImages(XSLFSlide summarySlide) throws IOException {
        final XSLFTextShape barchart = summarySlide.getPlaceholder(0);
        populateImage(barchart, "src/main/resources/barchart.png");

        final XSLFTextShape happiness = summarySlide.getPlaceholder(1);
        populateImage(happiness, "src/main/resources/happiness-trend.png");

        final XSLFTextShape replies = summarySlide.getPlaceholder(2);
        populateImage(replies, "src/main/resources/replies.png");

        final XSLFTextShape benchmark = summarySlide.getPlaceholder(3);
        populateImage(benchmark, "src/main/resources/benchmark.png");
    }

    private static void populateSummaryText(Map<String, XSLFTextShape> placeholdersByName) {
        placeholdersByName.get("question").setText("On a scale from 1 to 10, how happy are you at work?");
        placeholdersByName.get("#responses").setText("99");
        placeholdersByName.get("#cheers").setText("123");
        placeholdersByName.get("#vss").setText("12");
        placeholdersByName.get("date").setText("Mar 27th, 2016");
    }

    private static void debugShapes(XSLFSlide slide) {
        // http://stackoverflow.com/questions/35721547/how-to-add-image-to-image-placeholder-added-in-pptx-using-apache-poi-api
        // read all shapes i.e place holder in array.
        List<XSLFShape> shapes = slide.getShapes();
        System.out.println("=== Shapes");
        for (int i = 0; i < shapes.size(); i++) {
            System.out.println(shapes.get(i).getShapeName());
        }

        System.out.println("=== Placeholder");
        XSLFTextShape[] placeholders = slide.getPlaceholders();
        for (int i = 0; i < placeholders.length; i++) {
            System.out.println(placeholders[i].getShapeName());
        }
    }

    private static void populateIntroSlide(XSLFSlide introSlide) throws IOException {
        introSlide.getPlaceholder(1).setText("Engagement Report");
        introSlide.getPlaceholder(0).setText("April 2016");

        XSLFShape logoPlaceholder = introSlide.getPlaceholder(2);
        Rectangle2D anchor = logoPlaceholder.getAnchor();

        byte[] pictureData = IOUtils.toByteArray(new FileInputStream("src/main/resources/Microsoft.png"));
        XSLFPictureData xslfPictureData = introSlide.getSlideShow().addPicture(pictureData, PictureData.PictureType.PNG);
        XSLFPictureShape picture = introSlide.createPicture(xslfPictureData);
        introSlide.removeShape(logoPlaceholder);

        picture.setAnchor(anchor);
    }

    private static void populateImage(XSLFShape placeholder, String imagePath) throws IOException {
        final XSLFSlide slide = (XSLFSlide) placeholder.getSheet();
        final XMLSlideShow slideShow = slide.getSlideShow();
        final Rectangle2D anchor = placeholder.getAnchor();

        byte[] pictureData = IOUtils.toByteArray(new FileInputStream(imagePath));
        XSLFPictureData xslfPictureData = slideShow.addPicture(pictureData, PictureData.PictureType.PNG);
        XSLFPictureShape picture = slide.createPicture(xslfPictureData);
        slide.removeShape(placeholder);

        picture.setAnchor(anchor);
    }

    public static void save(XMLSlideShow ppt, String filename) throws IOException {
        FileOutputStream out = new FileOutputStream(filename);
        ppt.write(out);
        out.close();
        ppt.close();
    }
}
