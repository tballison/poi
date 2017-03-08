package org.apache.poi.xssf.eventusermodel;


import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.List;

import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.openxml4j.opc.PackagePart;
import org.apache.poi.util.LittleEndian;
import org.apache.poi.xssf.binary.BinaryParseException;
import org.apache.poi.xssf.binary.RichStr;
import org.apache.poi.xssf.binary.XSSFBinaryRecordType;
import org.apache.poi.xssf.usermodel.XSSFRelation;
import org.xml.sax.SAXException;

public class ReadOnlyBinarySharedStringsTable {

    /**
     * An integer representing the total count of strings in the workbook. This count does not
     * include any numbers, it counts only the total of text strings in the workbook.
     */
    private int count;

    /**
     * An integer representing the total count of unique strings in the Shared String Table.
     * A string is unique even if it is a copy of another string, but has different formatting applied
     * at the character level.
     */
    private int uniqueCount;

    /**
     * The shared strings table.
     */
    private List<String> strings = new ArrayList<String>();

    /**
     * @param pkg The {@link OPCPackage} to use as basis for the shared-strings table.
     * @throws IOException If reading the data from the package fails.
     * @throws SAXException if parsing the XML data fails.
     */
    public ReadOnlyBinarySharedStringsTable(OPCPackage pkg)
            throws IOException, SAXException {
        ArrayList<PackagePart> parts =
                pkg.getPartsByContentType(XSSFRelation.SHARED_STRINGS_BINARY.getContentType());

        // Some workbooks have no shared strings table.
        if (parts.size() > 0) {
            PackagePart sstPart = parts.get(0);

            readFrom(sstPart.getInputStream());
        }
    }

    /**
     * Like POIXMLDocumentPart constructor
     *
     * @since POI 3.14-Beta3
     */
    public ReadOnlyBinarySharedStringsTable(PackagePart part) throws IOException, SAXException {
        readFrom(part.getInputStream());
    }

    private void readFrom(InputStream inputStream) throws IOException {
        SSTBinaryReader reader = new SSTBinaryReader(inputStream);
        reader.parse();
    }

    public List<String> getItems() {
        return strings;
    }

    public String getEntryAt(int i) {
        return strings.get(i);
    }

    /**
     * Return an integer representing the total count of strings in the workbook. This count does not
     * include any numbers, it counts only the total of text strings in the workbook.
     *
     * @return the total count of strings in the workbook
     */
    public int getCount() {
        return this.count;
    }

    /**
     * Returns an integer representing the total count of unique strings in the Shared String Table.
     * A string is unique even if it is a copy of another string, but has different formatting applied
     * at the character level.
     *
     * @return the total count of unique strings in the workbook
     */
    public int getUniqueCount() {
        return this.uniqueCount;
    }

    private class SSTBinaryReader extends BinaryReader {

        SSTBinaryReader(InputStream is) {
            super(is);
        }

        @Override
        public void handleRecord(int recordType, byte[] data) throws BinaryParseException {
            XSSFBinaryRecordType type = XSSFBinaryRecordType.BRtBeginSst.lookup(recordType);

            switch (type) {
                case BRtSstItem :
                    RichStr rstr = RichStr.build(data, 0);
                    strings.add(rstr.getString());
                    break;
                case BRtBeginSst:
                    count = (int)LittleEndian.getUInt(data,0);
                    uniqueCount = (int)LittleEndian.getUInt(data, 4);
                    break;
            }

        }
    }

}
