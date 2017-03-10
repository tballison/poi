package org.apache.poi.xssf.binary;

import org.apache.poi.ss.util.CellReference;
import org.apache.poi.util.LittleEndian;

public class XSSFBinaryCell {
    public static int length = 8;

    /**
     *
     * @param data
     * @param offset
     * @param currentRow 0-offset based current row count
     * @param cell
     */
    public static void parse(byte[] data, int offset,  int currentRow, XSSFBinaryCell cell) {
        long colNum = LittleEndian.getUInt(data, offset); offset += LittleEndian.INT_SIZE;
        int styleIdx = LittleEndian.get24BitInt(data, offset); offset += 3;
        //TODO: range checking
        boolean showPhonetic = false;//TODO: fill this out
        cell.reset(currentRow, (int)colNum, styleIdx, showPhonetic);
    }

    private int rowNum;
    private int colNum;
    private int styleIdx;
    private boolean showPhonetic;

    public void reset(int rowNum, int colNum, int styleIdx, boolean showPhonetic) {
        this.rowNum = rowNum;
        this.colNum = colNum;
        this.styleIdx = styleIdx;
        this.showPhonetic = showPhonetic;
    }

    public int getColNum() {
        return colNum;
    }

    public String formatAddressAsString() {
        return CellReference.convertNumToColString(colNum)+(rowNum+1);
    }

    public int getStyleIdx() {
        return styleIdx;
    }
}
