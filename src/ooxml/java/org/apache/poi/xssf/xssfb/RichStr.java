package org.apache.poi.xssf.xssfb;


import org.apache.poi.ooxmlb.POIXMLBException;

public class RichStr {


    public static RichStr build(byte[] bytes, int offset) throws POIXMLBException {
        byte first = bytes[offset];
        boolean dwSizeStrRunExists = (first >> 7 & 1) == 1;//first bit == 1?
        boolean phoneticExists = (first >> 6 & 1) == 1;//second bit == 1?
        StringBuilder sb = new StringBuilder();

        int read = XSSFBUtils.readXLWideString(bytes, offset+1, sb);
        //TODO: parse phonetic strings.
        return new RichStr(sb.toString(), "");
    }

    private final String string;
    private final String phoneticString;

    public RichStr(String string, String phoneticString) {
        this.string = string;
        this.phoneticString = phoneticString;
    }

    public String getString() {
        return string;
    }
}
