package org.apache.poi.xssf.eventusermodel;

import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.LinkedList;
import java.util.List;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.openxml4j.exceptions.OpenXML4JException;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.openxml4j.opc.PackagePart;
import org.apache.poi.openxml4j.opc.PackagePartName;
import org.apache.poi.openxml4j.opc.PackageRelationship;
import org.apache.poi.openxml4j.opc.PackageRelationshipCollection;
import org.apache.poi.openxml4j.opc.PackagingURIHelper;
import org.apache.poi.util.LittleEndian;
import org.apache.poi.xssf.binary.BinaryParseException;
import org.apache.poi.xssf.binary.XSSFBUtils;
import org.apache.poi.xssf.binary.XSSFBinaryRecordType;
import org.apache.poi.xssf.usermodel.XSSFRelation;

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

    public XSSFBStylesTable getXSSFBStylesTable() throws IOException {
        ArrayList<PackagePart> parts = pkg.getPartsByContentType( XSSFRelation.STYLES_BINARY.getContentType());
        if(parts.size() == 0) return null;

        // Create the Styles Table, and associate the Themes if present
        return new XSSFBStylesTable(parts.get(0).getInputStream());

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

        public XSSFBCommentsTable getXSSFBSheetComments() {
            PackagePart sheetPkg = getSheetPart();

            // Do we have a comments relationship? (Only ever one if so)
            try {
                PackageRelationshipCollection commentsList =
                        sheetPkg.getRelationshipsByType(XSSFRelation.SHEET_COMMENTS.getRelation());
                if (commentsList.size() > 0) {
                    PackageRelationship comments = commentsList.getRelationship(0);
                    PackagePartName commentsName = PackagingURIHelper.createPartName(comments.getTargetURI());
                    PackagePart commentsPart = sheetPkg.getPackage().getPart(commentsName);
                    return new XSSFBCommentsTable(commentsPart.getInputStream());
                }
            } catch (InvalidFormatException e) {
                return null;
            } catch (IOException e) {
                return null;
            }
            return null;
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
            offset += XSSFBUtils.readXLWideString(data, offset, sb);
            String relId = sb.toString();
            sb.setLength(0);
            XSSFBUtils.readXLWideString(data, offset, sb);
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