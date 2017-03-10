package org.apache.poi.xssf.xssfb;

public class XSSFBHeaderFooters {

    public static XSSFBHeaderFooters parse(byte[] data) {
        boolean diffOddEven = false;
        boolean diffFirst = false;
        boolean scaleWDoc = false;
        boolean alignMargins = false;

        int offset = 2;
        XSSFBHeaderFooters xssfbHeaderFooter = new XSSFBHeaderFooters();
        xssfbHeaderFooter.header = new XSSFBHeaderFooter("header", true);
        xssfbHeaderFooter.footer = new XSSFBHeaderFooter("footer", false);
        xssfbHeaderFooter.headerEven = new XSSFBHeaderFooter("evenHeader", true);
        xssfbHeaderFooter.footerEven = new XSSFBHeaderFooter("evenFooter", false);
        xssfbHeaderFooter.headerFirst = new XSSFBHeaderFooter("firstHeader", true);
        xssfbHeaderFooter.footerFirst = new XSSFBHeaderFooter("firstFooter", false);
        offset += readHeaderFooter(data, offset, xssfbHeaderFooter.header);
        offset += readHeaderFooter(data, offset, xssfbHeaderFooter.footer);
        offset += readHeaderFooter(data, offset, xssfbHeaderFooter.headerEven);
        offset += readHeaderFooter(data, offset, xssfbHeaderFooter.footerEven);
        offset += readHeaderFooter(data, offset, xssfbHeaderFooter.headerFirst);
        readHeaderFooter(data, offset, xssfbHeaderFooter.footerFirst);
        return xssfbHeaderFooter;
    }

    private static int readHeaderFooter(byte[] data, int offset, XSSFBHeaderFooter headerFooter) {
        if (offset + 4 >= data.length) {
            return 0;
        }
        StringBuilder sb = new StringBuilder();
        int bytesRead = XSSFBUtils.readXLNullableWideString(data, offset, sb);
        headerFooter.setRawString(sb.toString());
        return bytesRead;
    }

    private XSSFBHeaderFooter header;
    private XSSFBHeaderFooter footer;
    private XSSFBHeaderFooter headerEven;
    private XSSFBHeaderFooter footerEven;
    private XSSFBHeaderFooter headerFirst;
    private XSSFBHeaderFooter footerFirst;

    public XSSFBHeaderFooter getHeader() {
        return header;
    }

    public XSSFBHeaderFooter getFooter() {
        return footer;
    }

    public XSSFBHeaderFooter getHeaderEven() {
        return headerEven;
    }

    public XSSFBHeaderFooter getFooterEven() {
        return footerEven;
    }

    public XSSFBHeaderFooter getHeaderFirst() {
        return headerFirst;
    }

    public XSSFBHeaderFooter getFooterFirst() {
        return footerFirst;
    }
}
