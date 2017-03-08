package org.apache.poi.xssf.binary;

import java.nio.charset.Charset;

import org.apache.poi.util.LittleEndian;

public class XLWideString {
    public static int read(byte[] data, int offset, StringBuilder sb) throws BinaryParseException {
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
}
