package org.apache.poi.xssf.eventusermodel;


import java.io.InputStream;

import org.apache.poi.ss.util.CellAddress;
import org.apache.poi.util.LittleEndian;
import org.apache.poi.xssf.binary.BinaryParseException;
import org.apache.poi.xssf.binary.XSSFBinaryCell;
import org.apache.poi.xssf.binary.XSSFBinaryRecordType;
import org.apache.poi.xssf.usermodel.XSSFRichTextString;

public class XSSFSheetBinaryHandler extends BinaryReader {
    private final ReadOnlyBinarySharedStringsTable stringsTable;
    private final XSSFSheetXMLHandler.SheetContentsHandler handler;
    int currentRow = 0;
    int rows = 0;

    private final XSSFBinaryCell cellBuffer = new XSSFBinaryCell();
    public XSSFSheetBinaryHandler(InputStream is, XSSFSheetXMLHandler.SheetContentsHandler handler, ReadOnlyBinarySharedStringsTable stringsTable) {
        super(is);
        this.handler = handler;
        this.stringsTable = stringsTable;
    }

    @Override
    public void handleRecord(int recordType, byte[] data) throws BinaryParseException {
        XSSFBinaryRecordType type = XSSFBinaryRecordType.BRtBeginSst.lookup(recordType);
        switch(type) {
            case BrtRowHdr:
                if (rows++ > 0) {
                    handler.endRow(currentRow);
                }
                long rw = LittleEndian.getUInt(data, 0);
                if (rw > 0x00100000L) {//could make sure this is larger than currentRow, according to spec?
                    throw new BinaryParseException("Row number beyond allowable range: "+rw);
                }
                currentRow = (int)rw;
                handler.startRow(currentRow);
                break;
            case BrtCellIsst:
                handleBrtCellIsst(data);
                break;
            case BrtCellRk:
                handleCellRk(data);
                break;
        }
    }

    private void handleCellRk(byte[] data) {
        XSSFBinaryCell.parse(data, 0, cellBuffer);
        long val = LittleEndian.getUInt(data, XSSFBinaryCell.length);
        CellAddress ca = new CellAddress(currentRow, cellBuffer.getColNum());
        handler.cell(ca.formatAsString(), Long.toString(val), null);

    }

    private void handleBrtCellIsst(byte[] data) {
        XSSFBinaryCell.parse(data, 0, cellBuffer);
        long idx = LittleEndian.getUInt(data, XSSFBinaryCell.length);
        //check for out of range, buffer overflow

        XSSFRichTextString rtss = new XSSFRichTextString(stringsTable.getEntryAt((int)idx));
        CellAddress ca = new CellAddress(currentRow, cellBuffer.getColNum());
        handler.cell(ca.formatAsString(), rtss.getString(), null);
    }

}
