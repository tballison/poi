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

import org.apache.poi.ss.util.CellReference;
import org.apache.poi.util.LittleEndian;

public class XSSFBCell {
    public static int length = 8;

    /**
     *
     * @param data
     * @param offset
     * @param currentRow 0-offset based current row count
     * @param cell
     */
    public static void parse(byte[] data, int offset,  int currentRow, XSSFBCell cell) {
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
