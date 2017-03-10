package org.apache.poi.xssf.binary;


import org.apache.poi.ss.usermodel.ClientAnchor;
import org.apache.poi.ss.usermodel.RichTextString;
import org.apache.poi.ss.util.CellAddress;
import org.apache.poi.xssf.usermodel.XSSFComment;

public class XSSFBComment extends XSSFComment {

    private final CellAddress cellAddress;
    private final String author;
    private final XSSFBRichTextString comment;
    private boolean visible = true;

    public XSSFBComment(CellAddress cellAddress, String author, String comment) {
        super(null, null, null);
        this.cellAddress = cellAddress;
        this.author = author;
        this.comment = new XSSFBRichTextString(comment);
    }

    @Override
    public void setVisible(boolean visible) {
        throw new IllegalArgumentException("XSSFBComment is read only.");
    }

    @Override
    public boolean isVisible() {
        return visible;
    }

    @Override
    public CellAddress getAddress() {
        return cellAddress;
    }

    @Override
    public void setAddress(CellAddress addr) {
        throw new IllegalArgumentException("XSSFBComment is read only");
    }

    @Override
    public void setAddress(int row, int col) {
        throw new IllegalArgumentException("XSSFBComment is read only");

    }

    @Override
    public int getRow() {
        return cellAddress.getRow();
    }

    @Override
    public void setRow(int row) {
        throw new IllegalArgumentException("XSSFBComment is read only");
    }

    @Override
    public int getColumn() {
        return cellAddress.getColumn();
    }

    @Override
    public void setColumn(int col) {
        throw new IllegalArgumentException("XSSFBComment is read only");
    }

    @Override
    public String getAuthor() {
        return author;
    }

    @Override
    public void setAuthor(String author) {
        throw new IllegalArgumentException("XSSFBComment is read only");
    }

    @Override
    public XSSFBRichTextString getString() {
        return comment;
    }

    @Override
    public void setString(RichTextString string) {
        throw new IllegalArgumentException("XSSFBComment is read only");
    }

    @Override
    public ClientAnchor getClientAnchor() {
        return null;
    }
}
