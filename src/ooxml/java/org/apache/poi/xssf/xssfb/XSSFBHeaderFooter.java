package org.apache.poi.xssf.xssfb;

import org.apache.poi.xssf.usermodel.helpers.HeaderFooterHelper;

public class XSSFBHeaderFooter {
    private final String headerFooterTypeLabel;
    private final boolean isHeader;
    private String rawString;
    private HeaderFooterHelper headerFooterHelper = new HeaderFooterHelper();


    public XSSFBHeaderFooter(String headerFooterTypeLabel, boolean isHeader) {
        this.headerFooterTypeLabel = headerFooterTypeLabel;
        this.isHeader = isHeader;
    }

    public String getHeaderFooterTypeLabel() {
        return headerFooterTypeLabel;
    }

    public String getRawString() {
        return rawString;
    }

    public String getString() {
        StringBuilder sb = new StringBuilder();
        String left = headerFooterHelper.getLeftSection(rawString);
        String center = headerFooterHelper.getCenterSection(rawString);
        String right = headerFooterHelper.getRightSection(rawString);
        if (left != null && left.length() > 0) {
            sb.append(left);
        }
        if (center != null && center.length() > 0) {
            if (sb.length() > 0) {
                sb.append(" ");
            }
            sb.append(center);
        }
        if (right != null && right.length() > 0) {
            if (sb.length() > 0) {
                sb.append(" ");
            }
            sb.append(right);
        }
        return sb.toString();
    }

    public void setRawString(String rawString) {
        this.rawString = rawString;
    }

    public boolean isHeader() {
        return isHeader;
    }

}
