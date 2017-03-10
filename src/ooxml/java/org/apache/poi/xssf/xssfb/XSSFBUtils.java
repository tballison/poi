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


import java.nio.charset.Charset;

import org.apache.poi.POIXMLException;
import org.apache.poi.ooxmlb.POIXMLBException;
import org.apache.poi.util.LittleEndian;

public class XSSFBUtils {

    public static int readXLNullableWideString(byte[] data, int offset, StringBuilder sb) throws POIXMLBException {
        long numChars = LittleEndian.getUInt(data, offset);
        if (numChars < 0) {
            throw new POIXMLBException("too few chars to read");
        } else if (numChars == 0xFFFFFFFFL) { //this means null value (2.5.166), do not read any bytes!!!
            return 0;
        } else if (numChars > 0xFFFFFFFFL) {
            throw new POIXMLBException("too many chars to read");
        }

        int numBytes = 2*(int)numChars;
        offset += 4;
        if (offset+numBytes > data.length) {
            throw new POIXMLBException("trying to read beyond data length");
        }
        sb.append(new String(data, offset, numBytes, Charset.forName("UTF-16LE")));
        numBytes+=4;
        return numBytes;
    }


    public static int readXLWideString(byte[] data, int offset, StringBuilder sb) throws POIXMLBException {
        long numChars = LittleEndian.getUInt(data, offset);
        if (numChars < 0) {
            throw new POIXMLBException("too few chars to read");
        } else if (numChars > 0xFFFFFFFFL) {
            throw new POIXMLBException("too many chars to read");
        }
        int numBytes = 2*(int)numChars;
        offset += 4;
        if (offset+numBytes > data.length) {
            throw new POIXMLBException("trying to read beyond data length");
        }
        sb.append(new String(data, offset, numBytes, Charset.forName("UTF-16LE")));
        numBytes+=4;
        return numBytes;
    }

    public static int castToInt(long val) {
        if (val < Integer.MAX_VALUE && val > Integer.MIN_VALUE) {
            return (int)val;
        }
        throw new POIXMLException("val ("+val+") can't be cast to int");
    }

    public static short castToShort(int val) {
        if (val < Short.MAX_VALUE && val > Short.MIN_VALUE) {
            return (short)val;
        }
        throw new POIXMLException("val ("+val+") can't be cast to short");

    }
}
