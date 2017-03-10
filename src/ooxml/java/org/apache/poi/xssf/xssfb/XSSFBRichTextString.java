package org.apache.poi.xssf.xssfb;

import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.xssf.usermodel.XSSFRichTextString;

/**
 * Wrapper class around String so that we can use it in Comment.
 * Nothing has been implemented yet except for {@link #getString()}.
 */
public class XSSFBRichTextString extends XSSFRichTextString {
    private final String string;

    public XSSFBRichTextString(String string) {
        this.string = string;
    }

    @Override
    public void applyFont(int startIndex, int endIndex, short fontIndex) {

    }

    @Override
    public void applyFont(int startIndex, int endIndex, Font font) {

    }

    @Override
    public void applyFont(Font font) {

    }

    @Override
    public void clearFormatting() {

    }

    @Override
    public String getString() {
        return string;
    }

    @Override
    public int length() {
        return string.length();
    }

    @Override
    public int numFormattingRuns() {
        return 0;
    }

    @Override
    public int getIndexOfFormattingRun(int index) {
        return 0;
    }

    @Override
    public void applyFont(short fontIndex) {

    }
}
