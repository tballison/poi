package org.apache.poi.xssf.eventusermodel;

import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.List;
import java.util.SortedMap;
import java.util.TreeMap;

import org.apache.poi.POIXMLException;
import org.apache.poi.ss.usermodel.BuiltinFormats;
import org.apache.poi.xssf.binary.BinaryParseException;
import org.apache.poi.xssf.binary.XSSFBUtils;
import org.apache.poi.xssf.binary.XSSFBinaryRecordType;

public class XSSFBStylesTable extends BinaryReader {

    private final SortedMap<Short, String> numberFormats = new TreeMap<Short,String>();
    private final List<Short> styleIds = new ArrayList<Short>();

    boolean inCellXFS = false;
    boolean inFmts = false;
    public XSSFBStylesTable(InputStream is) throws IOException {
        super(is);
        parse();
    }

    public String getNumberFormatString(int idx) {
        String fmt = null;
        if (numberFormats.containsKey(styleIds.get((short)idx))) {
            return numberFormats.get(styleIds.get((short)idx));
        }

        if(fmt == null) fmt = BuiltinFormats.getBuiltinFormat(styleIds.get((short)idx));
        return fmt;

    }

    @Override
    public void handleRecord(int recordType, byte[] data) throws BinaryParseException {
        XSSFBinaryRecordType type = XSSFBinaryRecordType.BRtBeginSst.lookup(recordType);
        switch (type) {
            case BrtBeginCellXFs:
                inCellXFS = true;
                break;
            case BrtEndCellXFs:
                inCellXFS = false;
            case BrtXf:
                if (inCellXFS) {
                    handleBrtXFInCellXF(data);
                }
                break;
            case BrtBeginFmts:
                inFmts = true;
                break;
            case BrtEndFmts:
                inFmts = false;
                break;
            case BrtFmt:
                if (inFmts) {
                    handleFormat(data);
                }
                break;

        }
    }

    private void handleFormat(byte[] data) {
        int ifmt = data[0] & 0xFF;
        if (ifmt > Short.MAX_VALUE) {
            throw new POIXMLException("Format id must be a short");
        }
        if (ifmt < 0) {
            throw new POIXMLException("Format id must be > 0");
        }
        StringBuilder sb = new StringBuilder();
        XSSFBUtils.readXLWideString(data, 2, sb);
        String fmt = sb.toString();
        numberFormats.put((short)ifmt, fmt);
    }

    private void handleBrtXFInCellXF(byte[] data) {
        int ifmtOffset = 2;
        //int ifmtLength = 2;

        //numFmtId in xml terms
        int ifmt = data[ifmtOffset] & 0xFF;//the second byte is ignored
        styleIds.add((short)ifmt);
    }
}
