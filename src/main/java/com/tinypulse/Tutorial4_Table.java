package com.tinypulse;

import java.awt.*;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.sl.usermodel.Placeholder;
import org.apache.poi.sl.usermodel.TableCell;
import org.apache.poi.sl.usermodel.TextParagraph;
import org.apache.poi.xslf.usermodel.*;

/**
 * http://svn.apache.org/repos/asf/poi/trunk/src/examples/src/org/apache/poi/xslf/usermodel/Tutorial4.java
 */
public class Tutorial4_Table {
    public static void main(String[] args) throws IOException{
        XMLSlideShow ppt = new XMLSlideShow();

        // XSLFSlide#createSlide() with no arguments creates a blank slide
        XSLFSlide slide = ppt.createSlide();

        XSLFTable tbl = slide.createTable();
        tbl.setAnchor(new Rectangle(50, 50, 450, 300));

        int numColumns = 3;
        int numRows = 5;
        XSLFTableRow headerRow = tbl.addRow();
        headerRow.setHeight(50);
        // header
        for(int i = 0; i < numColumns; i++) {
            XSLFTableCell th = headerRow.addCell();
            XSLFTextParagraph p = th.addNewTextParagraph();
            p.setTextAlign(TextParagraph.TextAlign.CENTER);
            XSLFTextRun r = p.addNewTextRun();
            r.setText("Header " + (i+1));
            r.setBold(true);
            r.setFontColor(Color.white);
            th.setFillColor(new Color(79, 129, 189));
            th.setBorderWidth(TableCell.BorderEdge.bottom, 2.0);
            th.setBorderColor(TableCell.BorderEdge.bottom, Color.white);

            tbl.setColumnWidth(i, 150);  // all columns are equally sized
        }

        // rows

        for(int rownum = 0; rownum < numRows; rownum ++){
            XSLFTableRow tr = tbl.addRow();
            tr.setHeight(50);
            // header
            for(int i = 0; i < numColumns; i++) {
                XSLFTableCell cell = tr.addCell();
                XSLFTextParagraph p = cell.addNewTextParagraph();
                XSLFTextRun r = p.addNewTextRun();

                r.setText("Cell " + (i+1));
                if(rownum % 2 == 0)
                    cell.setFillColor(new Color(208, 216, 232));
                else
                    cell.setFillColor(new Color(233, 247, 244));

            }

        }


        FileOutputStream out = new FileOutputStream("table.pptx");
        ppt.write(out);
        out.close();

        ppt.close();
    }
}
