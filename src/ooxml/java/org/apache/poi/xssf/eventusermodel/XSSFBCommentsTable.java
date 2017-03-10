package org.apache.poi.xssf.eventusermodel;

import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.Comparator;
import java.util.LinkedList;
import java.util.List;
import java.util.Map;
import java.util.Queue;
import java.util.TreeMap;

import org.apache.poi.ss.util.CellAddress;
import org.apache.poi.util.LittleEndian;
import org.apache.poi.xssf.binary.BinaryParseException;
import org.apache.poi.xssf.binary.RichStr;
import org.apache.poi.xssf.binary.XSSFBComment;
import org.apache.poi.xssf.binary.XSSFBUtils;
import org.apache.poi.xssf.binary.XSSFBinaryRecordType;

public class XSSFBCommentsTable extends BinaryReader {

    Map<CellAddress, XSSFBComment> comments = new TreeMap<CellAddress, XSSFBComment>(new CellAddressComparator());//String is the cellAddress A1
    Queue<CellAddress> commentAddresses = new LinkedList<CellAddress>();
    List<String> authors = new ArrayList<String>();

    //these are all used only during the, and they are mutable!
    int authorId = -1;
    CellAddress cellAddress = null;
    String comment = null;
    StringBuilder authorBuffer = new StringBuilder();


    public XSSFBCommentsTable(InputStream is) throws IOException {
        super(is);
        parse();
        commentAddresses.addAll(comments.keySet());
    }

    @Override
    public void handleRecord(int id, byte[] data) throws BinaryParseException {
        XSSFBinaryRecordType recordType = XSSFBinaryRecordType.lookup(id);
        System.out.println("COMMENTS TABLE: " + id + " : " + recordType + " : " + data.length);
        switch (recordType) {
            case BrtBeginComment:
                int offset = 0;
                authorId = XSSFBUtils.castToInt(LittleEndian.getUInt(data)); offset += LittleEndian.INT_SIZE;
                int rowFirst = XSSFBUtils.castToInt(LittleEndian.getUInt(data, offset)); offset += LittleEndian.INT_SIZE;
                int rowLast = XSSFBUtils.castToInt(LittleEndian.getUInt(data, offset)); offset += LittleEndian.INT_SIZE;
                int colFirst = XSSFBUtils.castToInt(LittleEndian.getUInt(data, offset)); offset += LittleEndian.INT_SIZE;
                int colLast = XSSFBUtils.castToInt(LittleEndian.getUInt(data, offset));
                //for strict parsing; confirm that rowFirst==rowLast and colFirst==colLats (2.4.28)
                cellAddress = new CellAddress(colFirst, rowFirst);
                break;
            case BrtCommentText:
                RichStr richStr = RichStr.build(data, 0);
                comment = richStr.getString();
                break;
            case BrtEndComment:
                comments.put(cellAddress, new XSSFBComment(cellAddress, authors.get(authorId), comment));
                authorId = -1;
                cellAddress = null;
                break;
            case BrtCommentAuthor:
                authorBuffer.setLength(0);
                XSSFBUtils.readXLWideString(data, 0, authorBuffer);
                authors.add(authorBuffer.toString());
                break;
        }
    }


    public Queue<CellAddress> getAddresses() {
        return commentAddresses;
    }

    public XSSFBComment get(CellAddress cellAddress) {
        if (cellAddress == null) {
            return null;
        }
        return comments.get(cellAddress);
    }

    private final static class CellAddressComparator implements Comparator<CellAddress> {

        @Override
        public int compare(CellAddress o1, CellAddress o2) {
            if (o1.getRow() < o2.getRow()) {
                return -1;
            } else if (o1.getRow() > o2.getRow()) {
                return 1;
            }
            if (o1.getColumn() < o2.getColumn()) {
                return -1;
            } else if (o1.getColumn() > o2.getColumn()) {
                return 1;
            }
            return 0;
        }
    }
}
