package org.apache.poi.xssf.eventusermodel;

import java.io.IOException;
import java.io.InputStream;
import java.util.Iterator;
import java.util.LinkedList;
import java.util.List;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.openxml4j.exceptions.OpenXML4JException;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.openxml4j.opc.PackagePart;
import org.apache.poi.util.LittleEndian;
import org.apache.poi.xssf.binary.BinaryParseException;
import org.apache.poi.xssf.binary.XLWideString;
import org.apache.poi.xssf.binary.XSSFBinaryRecordType;

public class XSSFBinaryReader extends XSSFReader {
    /**
     * Creates a new XSSFReader, for the given package
     *
     * @param pkg
     */
    public XSSFBinaryReader(OPCPackage pkg) throws IOException, OpenXML4JException {
        super(pkg);
    }

    /**
     * Returns an Iterator which will let you get at all the
     *  different Sheets in turn.
     * Each sheet's InputStream is only opened when fetched
     *  from the Iterator. It's up to you to close the
     *  InputStreams when done with each one.
     */
    @Override
    public Iterator<InputStream> getSheetsData() throws IOException, InvalidFormatException {
        return new SheetIterator(workbookPart);
    }

    public static class SheetIterator extends XSSFReader.SheetIterator {

        /**
         * Construct a new SheetIterator
         *
         * @param wb package part holding workbook.xml
         */
        private SheetIterator(PackagePart wb) throws IOException {
            super(wb);
        }

        Iterator<XSSFSheetRef> createSheetIteratorFromWB(PackagePart wb) throws IOException {
            SheetRefLoader sheetRefLoader = new SheetRefLoader(wb.getInputStream());
            sheetRefLoader.parse();
            return sheetRefLoader.getSheets().iterator();
        }

    }

    private static class SheetRefLoader extends BinaryReader {
        List<XSSFSheetRef> sheets = new LinkedList<XSSFSheetRef>();

        public SheetRefLoader(InputStream is) {
            super(is);
        }

        @Override
        public void handleRecord(int recordType, byte[] data) throws BinaryParseException {
            if (recordType == XSSFBinaryRecordType.BRtBundleSh.getId()) {
                addWorksheet(data);
            }
        }

        private void addWorksheet(byte[] data) {
            int offset = 0;
            //this is the sheet state #2.5.142
            long hsShtat = LittleEndian.getUInt(data, offset); offset += LittleEndian.INT_SIZE;

            long iTabID = LittleEndian.getUInt(data, offset); offset += LittleEndian.INT_SIZE;
            //according to #2.4.304
            if (iTabID < 1 || iTabID > 0x0000FFFFL) {
                throw new BinaryParseException("table id out of range: "+iTabID);
            }
            StringBuilder sb = new StringBuilder();
            offset += XLWideString.read(data, offset, sb);
            String relId = sb.toString();
            sb.setLength(0);
            XLWideString.read(data, offset, sb);
            String name = sb.toString();
            if (relId != null && relId.trim().length() > 0) {
                sheets.add(new XSSFSheetRef(relId, name));
            }
        }

        List<XSSFSheetRef> getSheets() {
            return sheets;
        }
    }
}