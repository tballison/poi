package org.apache.poi.xssf.binary;


import java.nio.charset.Charset;

import org.apache.poi.POIXMLException;
import org.apache.poi.util.LittleEndian;

public class XSSFBUtils {

    public static int readXLWideString(byte[] data, int offset, StringBuilder sb) throws BinaryParseException {
        long numChars = LittleEndian.getUInt(data, offset);
        if (numChars < 0) {
            throw new BinaryParseException("too few chars to read");
        } else if (numChars > 0xFFFFFFFFL) {
            throw new BinaryParseException("too many chars to read");
        }
        int numBytes = 2*(int)numChars;
        offset += 4;
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
