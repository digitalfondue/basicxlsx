# basicxlsx: when you only need to write simple xlsx

[![Build Status](https://travis-ci.org/digitalfondue/basicxlsx.svg?branch=master)](https://travis-ci.org/digitalfondue/basicxlsx) [![Coverage Status](https://coveralls.io/repos/github/digitalfondue/basicxlsx/badge.svg?branch=master)](https://coveralls.io/github/digitalfondue/basicxlsx?branch=master)

## Why

Apache POI, being a complete solution, is quite heavyweight.
This library provide only the minimal amount of functionality to
write xslx files.

## TODO:

 - support autoSizeColumn https://poi.apache.org/apidocs/org/apache/poi/xssf/usermodel/XSSFSheet.html#autoSizeColumn-int-
    - https://github.com/apache/poi/blob/trunk/src/ooxml/java/org/apache/poi/xssf/streaming/SXSSFSheet.java#L1571
    - https://github.com/apache/poi/blob/trunk/src/java/org/apache/poi/ss/util/SheetUtil.java#L120
    - https://github.com/dtjohnson/xlsx-populate/issues/26#issuecomment-288796920
    - https://metacpan.org/pod/Spreadsheet::WriteExcel::Examples#Example:-autofit.pl
 - support other column type
    - [x] boolean
    - [ ] date (wip, add LocalDate support, add Zoned* variant too?)
    - [ ] a specific subset of formula (sum, ?)
 - add test (WIP)
 - support some styling
    - [x] bold
    - [x] italic
    - [x] size
    - [x] font name
    - [x] font color (currently supporting RGB, but it seems that ARGB is possible)
    - [x] font underline
    - [x] font strikeout
    - [ ] alignment 
    - [ ] cell background
    - [ ] cell border
    - [x] number and date formatting
    - see https://xlsxwriter.readthedocs.io/format.html

## Resources/examples about xlsx

- http://officeopenxml.com/anatomyofOOXML-xlsx.php
- https://social.technet.microsoft.com/wiki/contents/articles/19601.powershell-generate-real-excel-xlsx-files-without-excel.aspx
- https://github.com/mk-j/PHP_XLSXWriter
- https://github.com/jmcnamara/XlsxWriter

## Notes

### License format
- `mvn com.mycila:license-maven-plugin:format`

### Check updates
- `mvn versions:display-dependency-updates`
- `mvn versions:display-plugin-updates`