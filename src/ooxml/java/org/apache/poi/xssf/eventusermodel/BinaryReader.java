package org.apache.poi.xssf.eventusermodel;

import java.io.IOException;
import java.io.InputStream;

import org.apache.poi.util.LittleEndianInputStream;
import org.apache.poi.xssf.binary.BinaryParseException;

abstract class BinaryReader {

    private final LittleEndianInputStream is;

    public BinaryReader(InputStream is) {
        this.is = new LittleEndianInputStream(is);
    }

    public void parse() throws IOException {

        while (true) {
            int bInt = is.read();
            if (bInt == -1) {
                return;
            }
            readNext((byte) bInt);
        }
    }

    private void readNext(byte b1) throws IOException {
        int recordId = 0;

        //if highest bit == 1
        if ((b1 >> 7 & 1) == 1) {
            byte b2 = is.readByte();
            b1 &= ~(1<<7); //unset highest bit
            b2 &= ~(1<<7); //unset highest bit (if it exists?)
            recordId = (128*(int)b2)+(int)b1;
        } else {
            recordId = (int)b1;
        }

        long recordLength = 0;
        int i = 0;
        boolean halt = false;
        int multiplier = 1;
        while (i < 4 && ! halt) {
            byte b = is.readByte();
            halt = (b >> 7 & 1) == 0; //if highest bit !=1 then continue
            b &= ~(1<<7);
            int lenToAdd = multiplier *(int)b;
            recordLength += lenToAdd;
            multiplier *= 128;
            i++;

        }

        byte[] buff = new byte[(int)recordLength];
        is.readFully(buff);
        handleRecord(recordId, buff);

    }

    abstract public void handleRecord(int recordType, byte[] bytes) throws BinaryParseException;

}
