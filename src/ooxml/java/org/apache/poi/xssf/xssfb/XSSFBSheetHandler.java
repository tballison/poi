/* ====================================================================
   Licensed to the Apache Software Foundation (ASF) under one or more
   contributor license agreements.  See the NOTICE file distributed with
   this work for additional information regarding copyright ownership.
   The ASF licenses this file to You under the Apache License, Version 2.0
   (the "License"); you may not use this file except in compliance with
   the License.  You may obtain a copy of the License at

       http://www.apache.org/licenses/LICENSE-2.0

   Unless required by applicable law or agreed to in writing, software
   distributed under the License is distributed on an "AS IS" BASIS,
   WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
   See the License for the specific language governing permissions and
   limitations under the License.
==================================================================== */

package org.apache.poi.xssf.xssfb;


import java.io.InputStream;
import java.util.Queue;

import org.apache.poi.ooxmlb.OOXMLBParser;
import org.apache.poi.ooxmlb.POIXMLBException;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.util.CellAddress;
import org.apache.poi.util.LittleEndian;
import org.apache.poi.xssf.eventusermodel.XSSFSheetXMLHandler;
import org.apache.poi.xssf.usermodel.XSSFRichTextString;

public class XSSFBSheetHandler extends OOXMLBParser {

    private final static int CHECK_ALL_ROWS = -1;

    private final ReadOnlyBinarySharedStringsTable stringsTable;
    private final XSSFSheetXMLHandler.SheetContentsHandler handler;
    private final XSSFBStylesTable styles;
    private final XSSFBCommentsTable comments;
    private final DataFormatter dataFormatter;
    private final boolean formulasNotResults;//TODO: implement this

    int lastEndedRow = -1;
    int lastStartedRow = -1;
    int currentRow = 0;
    byte[] rkBuffer = new byte[8];
    StringBuilder xlWideStringBuffer = new StringBuilder();

    private final XSSFBCell cellBuffer = new XSSFBCell();
    public XSSFBSheetHandler(InputStream is,
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
    public void handleRecord(int id, byte[] data) throws POIXMLBException {
        XSSFBRecordType type = XSSFBRecordType.BrtBeginSst.lookup(id);

        switch(type) {
            case BrtRowHdr:
                long rw = LittleEndian.getUInt(data, 0);
                if (rw > 0x00100000L) {//could make sure this is larger than currentRow, according to spec?
                    throw new POIXMLBException("Row number beyond allowable range: "+rw);
                }
                currentRow = (int)rw;
                checkMissedComments(currentRow);
                startRow(currentRow);
                break;
            case BrtCellIsst:
                handleBrtCellIsst(data);
                break;
            case BrtCellSt: //TODO: needs test
                handleCellSt(data);
                break;
            case BrtCellRk:
                handleCellRk(data);
                break;
            case BrtCellReal:
                handleCellReal(data);
                break;
            case BrtCellBool:
                handleBoolean(data);
                break;
            case BrtCellError:
                handleCellError(data);
                break;
            case BrtCellBlank:
                beforeCellValue(data);//read cell info and check for missing comments
                break;
            case BrtFmlaString:
                handleFmlaString(data);
                break;
            case BrtFmlaNum:
                handleFmlaNum(data);
                break;
            case BrtFmlaError:
                handleFmlaError(data);
                break;
                //TODO: All the PCDI and PCDIA
            case BrtEndSheetData:
                checkMissedComments(CHECK_ALL_ROWS);
                endRow(lastStartedRow);
                break;
            case BrtBeginHeaderFooter:
                handleHeaderFooter(data);
                break;
        }
    }


    private void beforeCellValue(byte[] data) {
        XSSFBCell.parse(data, 0, currentRow, cellBuffer);
        checkMissedComments(currentRow, cellBuffer.getColNum());
    }

    private void handleCellValue(String formattedValue) {
        CellAddress cellAddress = new CellAddress(currentRow, cellBuffer.getColNum());
        XSSFBComment comment = null;
        if (comments != null) {
            comment = comments.get(cellAddress);
        }
        handler.cell(cellAddress.formatAsString(), formattedValue, comment);
    }

    private void handleFmlaNum(byte[] data) {
        beforeCellValue(data);
        //xNum
        double val = LittleEndian.getDouble(data, XSSFBCell.length);
        String formatString = styles.getNumberFormatString(cellBuffer.getStyleIdx());
        String formattedVal = dataFormatter.formatRawCellContents(val, cellBuffer.getStyleIdx(), formatString);
        handleCellValue(formattedVal);
    }

    private void handleCellSt(byte[] data) {
        beforeCellValue(data);
        xlWideStringBuffer.setLength(0);
        XSSFBUtils.readXLWideString(data, XSSFBCell.length, xlWideStringBuffer);
        handleCellValue(xlWideStringBuffer.toString());
    }

    private void handleFmlaString(byte[] data) {
        beforeCellValue(data);
        xlWideStringBuffer.setLength(0);
        XSSFBUtils.readXLWideString(data, XSSFBCell.length, xlWideStringBuffer);
        handleCellValue(xlWideStringBuffer.toString());
    }

    private void handleCellError(byte[] data) {
        beforeCellValue(data);
        //TODO, read byte to figure out the type of error
        handleCellValue("ERROR");
    }

    private void handleFmlaError(byte[] data) {
        beforeCellValue(data);
        //TODO, read byte to figure out the type of error
        handleCellValue("ERROR");
    }

    private void handleBoolean(byte[] data) {
        beforeCellValue(data);
        String formattedVal = (data[XSSFBCell.length] == 1) ? "TRUE" : "FALSE";
        handleCellValue(formattedVal);
    }

    private void handleCellReal(byte[] data) {
        beforeCellValue(data);
        //xNum
        double val = LittleEndian.getDouble(data, XSSFBCell.length);
        String formatString = styles.getNumberFormatString(cellBuffer.getStyleIdx());
        String formattedVal = dataFormatter.formatRawCellContents(val, cellBuffer.getStyleIdx(), formatString);
        handleCellValue(formattedVal);
    }

    private void handleCellRk(byte[] data) {
        beforeCellValue(data);
        double val = rkNumber(data, XSSFBCell.length);
        String formatString = styles.getNumberFormatString(cellBuffer.getStyleIdx());
        String formattedVal = dataFormatter.formatRawCellContents(val, cellBuffer.getStyleIdx(), formatString);
        handleCellValue(formattedVal);
    }

    private void handleBrtCellIsst(byte[] data) {
        beforeCellValue(data);
        long idx = LittleEndian.getUInt(data, XSSFBCell.length);
        //check for out of range, buffer overflow

        XSSFRichTextString rtss = new XSSFRichTextString(stringsTable.getEntryAt((int)idx));
        handleCellValue(rtss.getString());
    }


    private void handleHeaderFooter(byte[] data) {
        XSSFBHeaderFooters headerFooter = XSSFBHeaderFooters.parse(data);
        outputHeaderFooter(headerFooter.getHeader());
        outputHeaderFooter(headerFooter.getFooter());
        outputHeaderFooter(headerFooter.getHeaderEven());
        outputHeaderFooter(headerFooter.getFooterEven());
        outputHeaderFooter(headerFooter.getHeaderFirst());
        outputHeaderFooter(headerFooter.getFooterFirst());
    }

    private void outputHeaderFooter(XSSFBHeaderFooter headerFooter) {
        String text = headerFooter.getString();
        if (text != null && text.trim().length() > 0) {
            handler.headerFooter(text, headerFooter.isHeader(), headerFooter.getHeaderFooterTypeLabel());
        }
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
                    startRow(cellAddress.getRow());
                }
                dumpEmptyCellComment(cellAddress, comments.get(cellAddress));
                lastInterpolatedRow = cellAddress.getRow();
            } else {
                break;
            }
        }

    }

    private void startRow(int row) {
        if (row == lastStartedRow) {
            return;
        }

        if (lastStartedRow != lastEndedRow) {
            endRow(lastStartedRow);
        }
        handler.startRow(row);
        lastStartedRow = row;
    }

    private void endRow(int row) {
        if (lastEndedRow == row) {
            return;
        }
        handler.endRow(row);
        lastEndedRow = row;
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
