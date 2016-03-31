package com.tinypulse;

import java.awt.Rectangle;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.sl.usermodel.Placeholder;
import org.apache.poi.xslf.usermodel.XMLSlideShow;
import org.apache.poi.xslf.usermodel.XSLFSlide;
import org.apache.poi.xslf.usermodel.XSLFTextShape;

/**
 * http://svn.apache.org/repos/asf/poi/trunk/src/examples/src/org/apache/poi/xslf/usermodel/Tutorial3.java
 */
public class Tutorial3_SimpleTitle {
    public static void main(String[] args) throws IOException{
        XMLSlideShow ppt = new XMLSlideShow();

        XSLFSlide slide = ppt.createSlide();

        XSLFTextShape titleShape = slide.createTextBox();
        titleShape.setPlaceholder(Placeholder.TITLE);
        titleShape.setText("This is a slide title");
        titleShape.setAnchor(new Rectangle(50, 50, 400, 100));

        FileOutputStream out = new FileOutputStream("title.pptx");
        ppt.write(out);
        out.close();

        ppt.close();
    }
}
