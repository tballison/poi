package org.apache.poi.xssf.eventusermodel;

import static org.junit.Assert.assertNotNull;

import java.io.InputStream;
import java.util.Iterator;

import org.apache.poi.POIDataSamples;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.util.SAXHelper;
import org.apache.poi.xssf.usermodel.XSSFComment;
import org.junit.Test;
import org.xml.sax.InputSource;
import org.xml.sax.XMLReader;

public class TestXSSFBinaryReader {

    static {
        System.setProperty("POI.testdata.path", "C:/users/tallison/Idea Projects/poi-github/test-data");
    }
    private static POIDataSamples _ssTests = POIDataSamples.getSpreadSheetInstance();

    @Test
    public void testBasic() throws Exception {
        OPCPackage pkg = OPCPackage.open(_ssTests.openResourceAsStream("51519.xlsb"));

        XSSFBinaryReader r = new XSSFBinaryReader(pkg);

        assertNotNull(r.getWorkbookData());
        assertNotNull(r.getSharedStringsData());
        assertNotNull(r.getStylesData());
        ReadOnlyBinarySharedStringsTable sst = new ReadOnlyBinarySharedStringsTable(pkg);
        Iterator<InputStream> it = r.getSheetsData();
        while (it.hasNext()) {
            InputStream is = it.next();
            XSSFSheetBinaryHandler sheetHandler = new XSSFSheetBinaryHandler(is, new DebugSheetHandler(), sst);
            sheetHandler.parse();

        }

    }

    @Test
    public void testRegular() throws Exception {
        OPCPackage pkg = OPCPackage.open(_ssTests.openResourceAsStream("51519.xlsx"));

        XSSFReader r = new XSSFReader(pkg);

        assertNotNull(r.getWorkbookData());
        assertNotNull(r.getSharedStringsData());
        assertNotNull(r.getStylesData());
        ReadOnlySharedStringsTable sst = new ReadOnlySharedStringsTable(pkg);
        Iterator<InputStream> it = r.getSheetsData();
        while (it.hasNext()) {
            InputStream is = it.next();
            XSSFSheetXMLHandler sheetHandler =
                    new XSSFSheetXMLHandler(null, sst, new DebugSheetHandler(), true);
            XMLReader sheetParser = SAXHelper.newXMLReader();
            sheetParser.setContentHandler(sheetHandler);
            sheetParser.parse(new InputSource(is));


        }

    }


    private class DebugSheetHandler implements XSSFSheetXMLHandler.SheetContentsHandler {

        @Override
        public void startRow(int rowNum) {
            System.out.println("starting row "+rowNum);
        }

        @Override
        public void endRow(int rowNum) {
            System.out.println("ending row "+rowNum);
        }

        @Override
        public void cell(String cellReference, String formattedValue, XSSFComment comment) {
            System.out.println(cellReference + " : " + formattedValue);
        }

        @Override
        public void headerFooter(String text, boolean isHeader, String tagName) {

        }
    }
}
