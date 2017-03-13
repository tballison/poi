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

import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.List;
import java.util.SortedMap;
import java.util.TreeMap;

import org.apache.poi.POIXMLException;
import org.apache.poi.ss.usermodel.BuiltinFormats;
import org.apache.poi.util.Internal;

@Internal
public class XSSFBStylesTable extends XSSFBParser {

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
    public void handleRecord(int recordType, byte[] data) throws XSSFBParseException {
        XSSFBRecordType type = XSSFBRecordType.BrtBeginSst.lookup(recordType);
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
