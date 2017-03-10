package org.apache.poi.xssf.eventusermodel;

import static org.junit.Assert.assertEquals;

import java.util.List;
import java.util.regex.Pattern;

import org.apache.poi.POIDataSamples;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.openxml4j.opc.PackagePart;
import org.apache.poi.xssf.xssfb.ReadOnlyBinarySharedStringsTable;
import org.junit.Test;

public class TestReadOnlyBinarySharedStringsTable {

    static {
        System.setProperty("POI.testdata.path", "C:/users/tallison/Idea Projects/poi-github/test-data");
    }
    private static POIDataSamples _ssTests = POIDataSamples.getSpreadSheetInstance();

    @Test
    public void testBasic() throws Exception {

        OPCPackage pkg = OPCPackage.open(_ssTests.openResourceAsStream("51519.xlsb"));
        List<PackagePart> parts = pkg.getPartsByName(Pattern.compile("/xl/sharedStrings.bin"));
        assertEquals(1, parts.size());

        ReadOnlyBinarySharedStringsTable rtbl = new ReadOnlyBinarySharedStringsTable(parts.get(0));
        List<String> strings = rtbl.getItems();
        assertEquals(49, strings.size());

        assertEquals("\u30B3\u30E1\u30F3\u30C8", rtbl.getEntryAt(0));
        assertEquals("\u65E5\u672C\u30AA\u30E9\u30AF\u30EB", rtbl.getEntryAt(3));
        assertEquals(55, rtbl.getCount());
        assertEquals(49, rtbl.getUniqueCount());

        //TODO: add in tests for phonetic runs

    }


}
