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

package org.apache.poi.xssf.eventusermodel;

import static org.junit.Assert.assertEquals;

import java.io.InputStream;
import java.util.List;
import java.util.regex.Pattern;

import org.apache.poi.POIDataSamples;
import org.apache.poi.ooxmlb.OOXMLBParser;
import org.apache.poi.ooxmlb.POIXMLBException;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.openxml4j.opc.PackagePart;
import org.apache.poi.xssf.xssfb.XSSFBRichStr;
import org.apache.poi.xssf.xssfb.XSSFBRecordType;
import org.junit.Test;

public class TestBinaryReader {

    private static POIDataSamples _ssTests = POIDataSamples.getSpreadSheetInstance();

    @Test
    public void testXLSBSST() throws Exception {

        OPCPackage pkg = OPCPackage.open(_ssTests.openResourceAsStream("51519.xlsb"));
        List<PackagePart> parts = pkg.getPartsByName(Pattern.compile("/xl/sharedStrings.bin"));
        assertEquals(1, parts.size());


        DebugBinaryReader reader = new DebugBinaryReader(parts.get(0).getInputStream());
        reader.parse();


    }

    @Test
    public void testOneOff() throws Exception {
        for (int i = 0; i < 300; i++) {
            byte b = (byte)i;
            if ( (b >> 7 & 1) == 1) {
                System.out.println("B1: " + i + " : " +  Integer.toString(b&0xff, 2));
            } else {
                System.out.println("B0: " + i + " : " +  Integer.toString(b&0xff, 2));
            }

        }
    }
    private static class DebugBinaryReader extends OOXMLBParser {

        DebugBinaryReader(InputStream is) {
            super(is);
        }

        @Override
        public void handleRecord(int recordType, byte[] bytes) throws POIXMLBException {
            if (recordType == XSSFBRecordType.BrtSstItem.getId()) {
                XSSFBRichStr rstr = XSSFBRichStr.build(bytes, 0);
            }

        }
    }
}
