package org.apache.poi.xssf.eventusermodel;


import java.io.InputStream;
import java.util.Queue;

import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.util.CellAddress;
import org.apache.poi.util.LittleEndian;
import org.apache.poi.xssf.binary.BinaryParseException;
import org.apache.poi.xssf.binary.XSSFBComment;
import org.apache.poi.xssf.binary.XSSFBinaryCell;
import org.apache.poi.xssf.binary.XSSFBinaryRecordType;
import org.apache.poi.xssf.usermodel.XSSFRichTextString;

public class XSSFSheetBinaryHandler extends BinaryReader {

    private final static int CHECK_ALL_ROWS = -1;

    private final ReadOnlyBinarySharedStringsTable stringsTable;
    private final XSSFSheetXMLHandler.SheetContentsHandler handler;
    private final XSSFBStylesTable styles;
    private final XSSFBCommentsTable comments;
    private final DataFormatter dataFormatter;
    private final boolean formulasNotResults;


    private int formatIndex = -1;
    private String formatString = "";

    int lastRow = 0;
    int currentRow = 0;
    int rows = 0;
    byte[] rkBuffer = new byte[8];

    private final XSSFBinaryCell cellBuffer = new XSSFBinaryCell();
    public XSSFSheetBinaryHandler(InputStream is,
                                  XSSFBStylesTable styles,
                                  XSSFBCommentsTable comments,
                                  ReadOnlyBinarySharedStringsTable strings,
                                  XSSFSheetXMLHandler.SheetContentsHandler sheetContentsHandler,
                                  DataFormatter dataFormatter,
                                  boolean formulasNotResults) {
        super(is);
        this.styles = styles;
        this.comments = comments;
        this.stringsTable = strings;
        this.handler = sheetContentsHandler;
        this.dataFormatter = dataFormatter;
        this.formulasNotResults = formulasNotResults;
    }

    @Override
    public void handleRecord(int id, byte[] data) throws BinaryParseException {
        XSSFBinaryRecordType type = XSSFBinaryRecordType.BRtBeginSst.lookup(id);
        System.out.println(id + " : " + type + " : " + data.length);

        switch(type) {
            case BrtRowHdr:
                long rw = LittleEndian.getUInt(data, 0);
                if (rw > 0x00100000L) {//could make sure this is larger than currentRow, according to spec?
                    throw new BinaryParseException("Row number beyond allowable range: "+rw);
                }
                currentRow = (int)rw;
                if (rows++ > 0) {
                    handler.endRow(lastRow);
                    lastRow = currentRow;
                }
                checkMissedComments(currentRow);
                handler.startRow(currentRow);
                break;
            case BrtCellIsst:
                handleBrtCellIsst(data);
                break;
            case BrtCellRk:
                handleCellRk(data);
                break;
            case BrtCellReal:
                handleCellReal(data);
                break;
            case BrtEndSheetData:
                checkMissedComments(CHECK_ALL_ROWS);
                break;
        }
    }

    private void handleCellReal(byte[] data) {
        XSSFBinaryCell.parse(data, 0, currentRow, cellBuffer);
        checkMissedComments(currentRow, cellBuffer.getColNum());

        double val = LittleEndian.getDouble(data, XSSFBinaryCell.length);
        String formatString = styles.getNumberFormatString(cellBuffer.getStyleIdx());
        String formattedVal = dataFormatter.formatRawCellContents(val, cellBuffer.getStyleIdx(), formatString);
        CellAddress cellAddress = new CellAddress(currentRow, cellBuffer.getColNum());
        handler.cell(cellAddress.formatAsString(), formattedVal, comments.get(cellAddress));

    }

    private void handleCellRk(byte[] data) {
        XSSFBinaryCell.parse(data, 0, currentRow, cellBuffer);
        checkMissedComments(currentRow, cellBuffer.getColNum());
        double val = rkNumber(data, XSSFBinaryCell.length);
        String formatString = styles.getNumberFormatString(cellBuffer.getStyleIdx());
        String formattedVal = dataFormatter.formatRawCellContents(val, cellBuffer.getStyleIdx(), formatString);
        CellAddress cellAddress = new CellAddress(currentRow, cellBuffer.getColNum());
        handler.cell(cellAddress.formatAsString(), formattedVal, comments.get(cellAddress));
    }

    private void handleBrtCellIsst(byte[] data) {
        XSSFBinaryCell.parse(data, 0, 0, cellBuffer);
        checkMissedComments(currentRow, cellBuffer.getColNum());
        long idx = LittleEndian.getUInt(data, XSSFBinaryCell.length);
        //check for out of range, buffer overflow

        XSSFRichTextString rtss = new XSSFRichTextString(stringsTable.getEntryAt((int)idx));
        CellAddress cellAddress = new CellAddress(currentRow, cellBuffer.getColNum());
        handler.cell(cellAddress.formatAsString(), rtss.getString(), comments.get(cellAddress));
    }

    //at start of next cell or end of row, return the cellAddress if it equals currentRow and col
    private void checkMissedComments(int currentRow, int colNum) {
        if (comments == null) {
            return;
        }
        Queue<CellAddress> queue = comments.getAddresses();
        while (queue.size() > 0) {
            CellAddress cellAddress = queue.peek();
            if (cellAddress.getRow() == currentRow && cellAddress.getColumn() < colNum) {
                cellAddress = queue.remove();
                dumpEmptyCellComment(cellAddress, comments.get(cellAddress));
            } else if (cellAddress.getRow() == currentRow && cellAddress.getColumn() == colNum) {
                queue.remove();
                return;
            } else if (cellAddress.getRow() == currentRow && cellAddress.getColumn() > colNum) {
                return;
            } else if (cellAddress.getRow() > currentRow) {
                return;
            }
        }
    }

    //check for anything from rows before
    private void checkMissedComments(int currentRow) {
        if (comments == null) {
            return;
        }
        Queue<CellAddress> queue = comments.getAddresses();
        int lastInterpolatedRow = -1;
        while (queue.size() > 0) {
            CellAddress cellAddress = queue.peek();
            if (currentRow == CHECK_ALL_ROWS || cellAddress.getRow() < currentRow) {
                cellAddress = queue.remove();
                if (cellAddress.getRow() != lastInterpolatedRow) {
                    if (lastInterpolatedRow > -1) {
                        handler.endRow(lastInterpolatedRow);
                    }
                    handler.startRow(cellAddress.getRow());
                }
                dumpEmptyCellComment(cellAddress, comments.get(cellAddress));
            } else {
                break;
            }
        }

        if (lastInterpolatedRow > -1) {
            handler.endRow(lastInterpolatedRow);
        }
    }

    private void dumpEmptyCellComment(CellAddress cellAddress, XSSFBComment comment) {
        handler.cell(cellAddress.formatAsString(), null, comment);
    }



    private double rkNumber(byte[] data, int offset) {
        //see 2.5.122 for this abomination
        byte b0 = data[offset];
        String s = Integer.toString(b0, 2);
        boolean numDivBy100 = ((b0 >> 0 & 1) == 1); // else as is
        boolean floatingPoint = ((b0 >> 1 & 1) == 0); // else signed integer

        //unset highest 2 bits
        b0 &= ~(1<<0);
        b0 &= ~(1<<1);

        rkBuffer[4] = b0;
        for (int i = 1; i < 4; i++) {
            rkBuffer[i+4] = data[offset+i];
        }
        double d = 0.0;
        if (floatingPoint) {
            d = LittleEndian.getDouble(rkBuffer);
        } else {
            d = LittleEndian.getInt(rkBuffer);
        }
        d = (numDivBy100) ? d = (d/100) : d;
        return d;
    }
}
