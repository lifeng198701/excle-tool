package org.wuwz.poi.core;

import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.openxml4j.exceptions.OpenXML4JException;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.ss.usermodel.BuiltinFormats;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.xssf.eventusermodel.XSSFReader;
import org.apache.poi.xssf.model.SharedStringsTable;
import org.apache.poi.xssf.model.StylesTable;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFRichTextString;
import org.wuwz.poi.hanlder.ReadHandler;
import org.xml.sax.Attributes;
import org.xml.sax.InputSource;
import org.xml.sax.SAXException;
import org.xml.sax.XMLReader;
import org.xml.sax.helpers.DefaultHandler;
import org.xml.sax.helpers.XMLReaderFactory;

public class XlsxReader
        extends DefaultHandler
{
    private SharedStringsTable mSharedStringsTable;
    private String mLastContents;
    private boolean mNextIsString;
    private int mSheetIndex = -1;
    private int mCurrentRowIndex = 0;
    private int mCurrentColumnIndex = 0;
    private boolean mIsTElement;
    private ReadHandler mReadHandler;
    private short mFormatIndex;
    private String mFormatString;
    private StylesTable mStylesTable;
    private CellValueType mNextDataType = CellValueType.STRING;
    private final DataFormatter mFormatter = new DataFormatter();
    private List<String> mRowData = new ArrayList();
    private String mPreviousRef = null;
    private String mCurrentRef = null;
    private String mMaxRef = null;
    private String mEmptyCellValue = null;

    public XlsxReader(ReadHandler handler)
    {
        this.mReadHandler = handler;
    }

    public XlsxReader setEmptyCellValue(String ecv)
    {
        this.mEmptyCellValue = ecv;
        return this;
    }

    public void process(String fileName)
            throws Exception
    {
        POIUtils.checkExcelFile(fileName);
        processAll(OPCPackage.open(fileName));
    }

    private void processAll(OPCPackage pkg)
            throws IOException, OpenXML4JException, InvalidFormatException, SAXException
    {
        XSSFReader xssfReader = new XSSFReader(pkg);
        this.mStylesTable = xssfReader.getStylesTable();
        SharedStringsTable sst = xssfReader.getSharedStringsTable();
        XMLReader parser = fetchSheetParser(sst);
        Iterator<InputStream> sheets = xssfReader.getSheetsData();
        while (sheets.hasNext())
        {
            this.mCurrentRowIndex = 0;
            this.mSheetIndex += 1;
            InputStream sheet = (InputStream)sheets.next();
            InputSource sheetSource = new InputSource(sheet);
            parser.parse(sheetSource);
            sheet.close();
        }
        pkg.close();
    }

    public void process(InputStream is, String fileName)
            throws Exception
    {
        POIUtils.checkExcelFile(fileName);
        processAll(OPCPackage.open(is));
    }

    public void process(String fileName, int sheetIndex)
            throws Exception
    {
        POIUtils.checkExcelFile(fileName);
        processBySheet(sheetIndex, OPCPackage.open(fileName));
    }

    private void processBySheet(int sheetIndex, OPCPackage pkg)
            throws IOException, OpenXML4JException, InvalidFormatException, SAXException
    {
        XSSFReader r = new XSSFReader(pkg);
        SharedStringsTable sst = r.getSharedStringsTable();

        XMLReader parser = fetchSheetParser(sst);



        InputStream sheet = r.getSheet("rId" + (sheetIndex + 1));
        this.mSheetIndex += 1;
        InputSource sheetSource = new InputSource(sheet);
        parser.parse(sheetSource);
        sheet.close();
        pkg.close();
    }

    public void process(InputStream is, String fileName, int sheetIndex)
            throws Exception
    {
        POIUtils.checkExcelFile(fileName);
        processBySheet(sheetIndex, OPCPackage.open(is));
    }

    private XMLReader fetchSheetParser(SharedStringsTable sst)
            throws SAXException
    {
        XMLReader parser = XMLReaderFactory.createXMLReader("org.apache.xerces.parsers.SAXParser");
        this.mSharedStringsTable = sst;
        parser.setContentHandler(this);
        return parser;
    }

    static enum CellValueType
    {
        BOOL,  ERROR,  FORMULA,  INLINESTR,  STRING,  NUMBER,  DATE,  NULL;

        private CellValueType() {}
    }

    public void setNextDataType(Attributes attributes)
    {
        this.mNextDataType = CellValueType.STRING;
        this.mFormatIndex = -1;
        this.mFormatString = null;
        String cellType = attributes.getValue("t");
        String cellStyleStr = attributes.getValue("s");
        if ("b".equals(cellType)) {
            this.mNextDataType = CellValueType.BOOL;
        } else if ("e".equals(cellType)) {
            this.mNextDataType = CellValueType.ERROR;
        } else if ("inlineStr".equals(cellType)) {
            this.mNextDataType = CellValueType.INLINESTR;
        } else if ("s".equals(cellType)) {
            this.mNextDataType = CellValueType.STRING;
        } else if ("str".equals(cellType)) {
            this.mNextDataType = CellValueType.FORMULA;
        }
        if (cellStyleStr != null)
        {
            int styleIndex = Integer.parseInt(cellStyleStr);
            XSSFCellStyle style = this.mStylesTable.getStyleAt(styleIndex);
            this.mFormatIndex = style.getDataFormat();
            this.mFormatString = style.getDataFormatString();
            if (this.mFormatString == null)
            {
                this.mNextDataType = CellValueType.NULL;
                this.mFormatString = BuiltinFormats.getBuiltinFormat(this.mFormatIndex);
            }
        }
    }

    public String getDataValue(String value, String newValue)
    {
        switch (this.mNextDataType.ordinal())
        {
            case 1:
                char first = value.charAt(0);
                newValue = first == '0' ? "FALSE" : "TRUE";
                break;
            case 2:
                newValue = "\"ERROR:" + value.toString() + '"';
                break;
            case 3:
                newValue = '"' + value.toString() + '"';
                break;
            case 4:
                newValue = new XSSFRichTextString(value.toString()).toString();
                break;
            case 5:
                newValue = String.valueOf(value);
                break;
            case 6:
                if (this.mFormatString != null) {
                    try
                    {
                        newValue = this.mFormatter.formatRawCellContents(Double.parseDouble(value), this.mFormatIndex, this.mFormatString).trim();
                    }
                    catch (NumberFormatException e)
                    {
                        newValue = this.mEmptyCellValue;
                    }
                } else {
                    newValue = value;
                }
                newValue = newValue != null ? newValue.replace("_", "").trim() : null;
                break;
            case 7:
                newValue = this.mFormatter.formatRawCellContents(Double.parseDouble(value), this.mFormatIndex, this.mFormatString);
                newValue = newValue.replace(" ", "T");
                break;
            default:
                newValue = this.mEmptyCellValue;
        }
        return newValue;
    }

    public void startElement(String uri, String localName, String name, Attributes attributes)
            throws SAXException
    {
        if ("c".equals(name))
        {
            setNextDataType(attributes);


            this.mPreviousRef = (this.mPreviousRef == null ? attributes.getValue("r") : this.mCurrentRef);

            this.mCurrentRef = attributes.getValue("r");

            String cellType = attributes.getValue("t");
            this.mNextIsString = ((cellType != null) && (cellType.equals("s")));
        }
        this.mIsTElement = "t".equals(name);

        this.mLastContents = "";
    }

    public void endElement(String uri, String localName, String name)
            throws SAXException
    {
        if (this.mNextIsString)
        {
            int idx = Integer.parseInt(this.mLastContents);
            this.mLastContents = new XSSFRichTextString(this.mSharedStringsTable.getEntryAt(idx)).toString();
            this.mNextIsString = false;
        }
        if (this.mIsTElement)
        {
            String value = this.mLastContents.trim();
            this.mRowData.add(this.mCurrentColumnIndex, value);
            this.mCurrentColumnIndex += 1;
            this.mIsTElement = false;
        }
        else if ("c".equals(name))
        {
            String value = getDataValue(this.mLastContents.trim(), "");
            if (!this.mCurrentRef.equals(this.mPreviousRef)) {
                for (int i = 0; i < countNullCell(this.mCurrentRef, this.mPreviousRef); i++)
                {
                    this.mRowData.add(this.mCurrentColumnIndex, this.mEmptyCellValue);
                    this.mCurrentColumnIndex += 1;
                }
            }
            this.mRowData.add(this.mCurrentColumnIndex, value);
            this.mCurrentColumnIndex += 1;
        }
        else if ("row".equals(name))
        {
            if (this.mCurrentRowIndex == 0) {
                this.mMaxRef = this.mCurrentRef;
            }
            if (this.mMaxRef != null) {
                for (int i = 0; i <= countNullCell(this.mMaxRef, this.mCurrentRef); i++)
                {
                    this.mRowData.add(this.mCurrentColumnIndex, this.mEmptyCellValue);
                    this.mCurrentColumnIndex += 1;
                }
            }
            if (!this.mRowData.isEmpty()) {
                this.mReadHandler.handler(this.mSheetIndex, this.mCurrentRowIndex, this.mRowData);
            }
            this.mRowData.clear();
            this.mCurrentRowIndex += 1;
            this.mCurrentColumnIndex = 0;
            this.mPreviousRef = null;
            this.mCurrentRef = null;
        }
    }

    public void characters(char[] ch, int start, int length)
            throws SAXException
    {
        this.mLastContents += new String(ch, start, length);
    }

    private int countNullCell(String ref, String ref2)
    {
        String xfd = ref.replaceAll("\\d+", "");
        String xfd_1 = ref2.replaceAll("\\d+", "");

        xfd = fillChar(xfd, 3, '@', true);
        xfd_1 = fillChar(xfd_1, 3, '@', true);

        char[] letter = xfd.toCharArray();
        char[] letter_1 = xfd_1.toCharArray();
        int res = (letter[0] - letter_1[0]) * 26 * 26 + (letter[1] - letter_1[1]) * 26 + (letter[2] - letter_1[2]);
        return res - 1;
    }

    private String fillChar(String str, int len, char let, boolean isPre)
    {
        int len_1 = str.length();
        if (len_1 < len) {
            if (isPre) {
                for (int i = 0; i < len - len_1; i++) {
                    str = let + str;
                }
            } else {
                for (int i = 0; i < len - len_1; i++) {
                    str = str + let;
                }
            }
        }
        return str;
    }
}
