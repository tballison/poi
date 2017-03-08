package org.apache.poi.xssf.eventusermodel;

import static org.junit.Assert.assertEquals;

import java.io.InputStream;
import java.util.List;
import java.util.regex.Pattern;

import org.apache.poi.POIDataSamples;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.openxml4j.opc.PackagePart;
import org.apache.poi.xssf.binary.BinaryParseException;
import org.apache.poi.xssf.binary.RichStr;
import org.apache.poi.xssf.binary.XSSFBinaryRecordType;
import org.junit.Test;

/**
 * Created by TALLISON on 3/8/2017.
 */
public class TestBinaryReader {
    static {
        System.setProperty("POI.testdata.path", "C:/users/tallison/Idea Projects/poi-github/test-data");
    }
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
    private static class DebugBinaryReader extends BinaryReader {

        DebugBinaryReader(InputStream is) {
            super(is);
        }

        @Override
        public void handleRecord(int recordType, byte[] bytes) throws BinaryParseException {
            if (recordType == XSSFBinaryRecordType.BRtSstItem.getId()) {
                RichStr rstr = RichStr.build(bytes, 0);
            }

        }
    }
}
