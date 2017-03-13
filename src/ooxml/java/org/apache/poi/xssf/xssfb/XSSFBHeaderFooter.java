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

package org.apache.poi.xssf.xssfb;

import org.apache.poi.util.Internal;
import org.apache.poi.xssf.usermodel.helpers.HeaderFooterHelper;

@Internal
class XSSFBHeaderFooter {
    private final String headerFooterTypeLabel;
    private final boolean isHeader;
    private String rawString;
    private HeaderFooterHelper headerFooterHelper = new HeaderFooterHelper();


    public XSSFBHeaderFooter(String headerFooterTypeLabel, boolean isHeader) {
        this.headerFooterTypeLabel = headerFooterTypeLabel;
        this.isHeader = isHeader;
    }

    public String getHeaderFooterTypeLabel() {
        return headerFooterTypeLabel;
    }

    public String getRawString() {
        return rawString;
    }

    public String getString() {
        StringBuilder sb = new StringBuilder();
        String left = headerFooterHelper.getLeftSection(rawString);
        String center = headerFooterHelper.getCenterSection(rawString);
        String right = headerFooterHelper.getRightSection(rawString);
        if (left != null && left.length() > 0) {
            sb.append(left);
        }
        if (center != null && center.length() > 0) {
            if (sb.length() > 0) {
                sb.append(" ");
            }
            sb.append(center);
        }
        if (right != null && right.length() > 0) {
            if (sb.length() > 0) {
                sb.append(" ");
            }
            sb.append(right);
        }
        return sb.toString();
    }

    public void setRawString(String rawString) {
        this.rawString = rawString;
    }

    public boolean isHeader() {
        return isHeader;
    }

}
