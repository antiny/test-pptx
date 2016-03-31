package com.tinypulse;

import org.apache.poi.sl.usermodel.PictureData;
import org.apache.poi.sl.usermodel.Placeholder;
import org.apache.poi.util.IOUtils;
import org.apache.poi.xslf.usermodel.*;

import java.awt.*;
import java.io.*;

public class TestMain {

    public static void main(String[] args) throws IOException {
        XMLSlideShow ppt = new XMLSlideShow();

        // general data
        XSLFSlideMaster defaultMaster = ppt.getSlideMasters().get(0);
        String filepath = "/Users/antran/workspace/pptx-assets/tinypulse_logo.png";
        XSLFPictureData logoIndex = addLogo(ppt, filepath);

        // slide 1
        // title - subtitle
        XSLFSlide slide = buildIntroSlide(ppt, defaultMaster);
        setLogo(slide, logoIndex);

        // slide 2
        // a big graph
        XSLFSlide slide2 = buildSummarySlide(ppt, defaultMaster);
        setLogo(slide2, logoIndex);

        // slide 3
        // box with comments
        XSLFSlide slide3 = buildCheersSlide(ppt, defaultMaster);
        setLogo(slide3, logoIndex);

        // slide 4
        // title with underlined text
        XSLFSlide slide4 = buildTitleSlide(ppt, defaultMaster);
        setLogo(slide4, logoIndex);

        save(ppt, "test-slide.pptx");
    }

    private static XSLFSlide buildTitleSlide(XMLSlideShow ppt, XSLFSlideMaster defaultMaster) {
        XSLFSlideLayout titleLayout = defaultMaster.getLayout(SlideLayout.BLANK);
        XSLFSlide slide = ppt.createSlide(titleLayout);

        XSLFAutoShape aTitle = slide.createAutoShape();
        aTitle.setAnchor(new Rectangle(new Point(50, 50), new Dimension(200, 50)));
        XSLFTextRun engagement = aTitle.setText("Engagement");
        engagement.setFontColor(hex2Rgb("#64d172"));
        engagement.setBold(true);
        engagement.setFontFamily("Aria");

        XSLFAutoShape aLine = slide.createAutoShape();
        aLine.setAnchor(new Rectangle(new Point(55, 90), new Dimension(200, 4)));
        aLine.setFillColor(hex2Rgb("#64d172"));

        return slide;
    }

    private static XSLFSlide buildCheersSlide(XMLSlideShow ppt, XSLFSlideMaster defaultMaster) {
        XSLFSlideLayout titleLayout = defaultMaster.getLayout(SlideLayout.BLANK);
        XSLFSlide slide = ppt.createSlide(titleLayout);

        final Point cheer1TopLeft = new Point(25, 100);
        addCheerBox(slide, cheer1TopLeft, "9", "My manager is great and has always made sure I'm headed in the right direction.");

        final Point cheer2TopLeft = nextCheerBox(cheer1TopLeft);
        addCheerBox(slide, cheer2TopLeft, "10", "I'm glad that we get to use our own offering to track my progress.");

        return slide;
    }

    private static Point nextCheerBox(Point point) {
        final Point nextPoint = (Point) point.clone();
        nextPoint.translate(0, 55 + 40);
        return nextPoint;
    }

    private static void addCheerBox(XSLFSlide slide, Point boxTopLeft, String boxContent, String cheerContent) {
        XSLFAutoShape rankingBox = slide.createAutoShape();

        final Dimension boxDimension = new Dimension(55, 55);
        rankingBox.setAnchor(new Rectangle(boxTopLeft, boxDimension));
        rankingBox.setFillColor(hex2Rgb("#2a6d9e"));
        rankingBox.setHorizontalCentered(true);

        XSLFTextParagraph xslfTextRuns = rankingBox.addNewTextParagraph();
        XSLFTextRun xslfTextRun = xslfTextRuns.addNewTextRun();
        xslfTextRun.setText(boxContent);
        xslfTextRun.setFontColor(Color.WHITE);

        final Point cheerTopLeft = (Point) boxTopLeft.clone();
        cheerTopLeft.translate((int) boxDimension.getWidth() + 5, 0);
        final Dimension cheerDimension = new Dimension(600, 65);
        final XSLFTextBox cheerBox = slide.createTextBox();
        cheerBox.setAnchor(new Rectangle(cheerTopLeft, cheerDimension));
        cheerBox.setLineColor(Color.white);
        cheerBox.setText(cheerContent);
        cheerBox.setFillColor(hex2Rgb("#e2e6ea"));
    }

    public static Color hex2Rgb(String colorStr) {
        return new Color(
                Integer.valueOf(colorStr.substring(1, 3), 16),
                Integer.valueOf(colorStr.substring(3, 5), 16),
                Integer.valueOf(colorStr.substring(5, 7), 16));
    }

    public static XSLFSlide buildSummarySlide(XMLSlideShow ppt, XSLFSlideMaster defaultMaster) throws IOException {
        XSLFSlideLayout titleLayout = defaultMaster.getLayout(SlideLayout.BLANK);
        XSLFSlide slide = ppt.createSlide(titleLayout);

        File img = new File("/Users/antran/workspace/pptx-assets/graph.png");
        byte[] data = IOUtils.toByteArray(new FileInputStream(img));
        XSLFPictureData pictureIndex = ppt.addPicture(data, PictureData.PictureType.PNG);

        XSLFPictureShape picture = slide.createPicture(pictureIndex);
        picture.setAnchor(new Rectangle(55, 250, 569, 200));

        XSLFTextShape titleShape = slide.createTextBox();
        titleShape.setPlaceholder(Placeholder.CENTERED_TITLE);
        titleShape.setAnchor(new Rectangle(50, 70, 600, 100));
        titleShape.setHorizontalCentered(true);
        XSLFTextRun xslfTextRun = titleShape.setText("Do you feel that your manager has clearly defined your roles and responsibilities and how it contributes to the success of the organization?");
        xslfTextRun.setFontColor(hex2Rgb("#214a7b"));
        xslfTextRun.setFontSize(Double.valueOf(28));

        return slide;
    }

    public static XSLFSlide buildIntroSlide(XMLSlideShow ppt, XSLFSlideMaster defaultMaster) {

        XSLFSlideLayout titleLayout = defaultMaster.getLayout(SlideLayout.TITLE);
        XSLFSlide slide = ppt.createSlide(titleLayout);

        XSLFTextShape title = slide.getPlaceholder(0);
        title.setText("TINYpulse Engagement Report");

        XSLFTextShape subtitle = slide.getPlaceholder(1);
        subtitle.setText("April 2016");

        return slide;
    }

    public static XSLFPictureData addLogo(XMLSlideShow ppt, String filepath) throws IOException {
        File img = new File(filepath);
        byte[] data = IOUtils.toByteArray(new FileInputStream(img));
        XSLFPictureData pictureIndex = ppt.addPicture(data, PictureData.PictureType.PNG);
        return pictureIndex;
    }

    public static void setLogo(XSLFSlide slide, XSLFPictureData logoIndex) {
        XSLFPictureShape picture = slide.createPicture(logoIndex);
        picture.setAnchor(new Rectangle(550, 10, 155, 25));
    }

    public static void save(XMLSlideShow ppt, String filename) throws IOException {
        FileOutputStream out = new FileOutputStream(filename);
        ppt.write(out);
        out.close();
        ppt.close();
    }
}
