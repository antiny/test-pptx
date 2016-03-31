package com.tinypulse;

import org.apache.poi.sl.usermodel.PictureData;
import org.apache.poi.util.IOUtils;
import org.apache.poi.xslf.usermodel.XMLSlideShow;
import org.apache.poi.xslf.usermodel.XSLFPictureData;
import org.apache.poi.xslf.usermodel.XSLFPictureShape;
import org.apache.poi.xslf.usermodel.XSLFSlide;

import java.awt.*;
import java.awt.geom.Rectangle2D;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

/**
 * http://svn.apache.org/repos/asf/poi/trunk/src/examples/src/org/apache/poi/xslf/usermodel/Tutorial5.java
 */
public class Tutorial5_InsertImage {
    public static void main(String[] args) throws IOException {
        XMLSlideShow ppt = new XMLSlideShow();

        XSLFSlide slide = ppt.createSlide();
        File img = new File("/Users/antran/workspace/pptx-assets/tinypulse_logo.png");
        byte[] data = IOUtils.toByteArray(new FileInputStream(img));
        XSLFPictureData pictureIndex = ppt.addPicture(data, PictureData.PictureType.PNG);

        XSLFPictureShape picture = slide.createPicture(pictureIndex);
        picture.setAnchor(new Rectangle(200, 200, 200, 50));

        FileOutputStream out = new FileOutputStream("images.pptx");
        ppt.write(out);
        out.close();

        ppt.close();
    }
}
