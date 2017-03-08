package org.apache.poi.xssf.binary;

import org.apache.poi.util.LittleEndian;

/**
 * Created by TALLISON on 3/8/2017.
 */
public class XSSFBinaryCell {
    public static int length = 8;

    public static void parse(byte[] data, int offset, XSSFBinaryCell cell) {
        long colNum = LittleEndian.getUInt(data, offset); offset += LittleEndian.INT_SIZE;
        int styleIdx = LittleEndian.get24BitInt(data, offset); offset += 3;
        //TODO: range checking
        boolean showPhonetic = false;//TODO: fill this out

        cell.reset((int)colNum, styleIdx, showPhonetic);
    }

    private int colNum;
    private int styleIdx;
    private boolean showPhonetic;
    public void reset(int colNum, int styleIdx, boolean showPhonetic) {
        this.colNum = colNum;
        this.styleIdx = styleIdx;
        this.showPhonetic = showPhonetic;
    }

    public int getColNum() {
        return colNum;
    }
}
