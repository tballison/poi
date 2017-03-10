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

package org.apache.poi.ooxmlb;

import java.io.IOException;
import java.io.InputStream;

import org.apache.poi.util.Internal;
import org.apache.poi.util.LittleEndianInputStream;

/**
 * Experimental parser for Microsoft's ooxml xssfb format
 */
@Internal
public abstract class OOXMLBParser {

    private final LittleEndianInputStream is;

    public OOXMLBParser(InputStream is) {
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
        while (i < 4 && ! halt) {
            byte b = is.readByte();
            halt = (b >> 7 & 1) == 0; //if highest bit !=1 then continue
            b &= ~(1<<7);
            recordLength += (int)b << (i*7); //multiply by 128^i
            i++;

        }

        byte[] buff = new byte[(int)recordLength];
        is.readFully(buff);
        handleRecord(recordId, buff);

    }

    abstract public void handleRecord(int recordType, byte[] data) throws POIXMLBException;

}
