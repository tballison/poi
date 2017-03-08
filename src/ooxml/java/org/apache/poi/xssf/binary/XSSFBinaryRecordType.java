package org.apache.poi.xssf.binary;

public enum XSSFBinaryRecordType {

    BrtCellBlank(1),
    BrtCellRk(2),
    BrtCellError(3),
    BrtCellBool(4),
    BrtCellReal(5),
    BrtCellSt(6),
    BrtCellIsst(7),
    BrtFmlaString(8),
    BrtFmlaNum(9),
    BrtFmlaBool(10),
    BrtFmlError(11),
    BrtRowHdr(0),
    BrtCellRString(62),
    BrtBeginSheet(129),
    BrtWsProp(147),
    BrtWsDim(148),
    BrtColInfo(60),
    BrtBeginSheetData(145),
    BrtEndSheetData(146),
    BRtSstItem(19),   //stored strings item
    BRtBeginSst(159), //stored strings begin sst
    BRtEndSst(160),   //stored strings end sst

    BRtBundleSh(156), //defines worksheet in wb part
    Unimplemented(-1);
    ;

    private final int id;

    XSSFBinaryRecordType(int id) {
        this.id = id;
    }

    public int getId() {
        return id;
    }

    public XSSFBinaryRecordType lookup(int id) {
        for (XSSFBinaryRecordType r : XSSFBinaryRecordType.values()) {
            if (r.id == id) {
                return r;
            }
        }
        return Unimplemented;
    }

}
