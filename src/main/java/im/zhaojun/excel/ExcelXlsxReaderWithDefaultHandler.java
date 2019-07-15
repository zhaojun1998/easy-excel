package im.zhaojun.excel;

import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.ss.usermodel.BuiltinFormats;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.xssf.eventusermodel.XSSFReader;
import org.apache.poi.xssf.model.SharedStringsTable;
import org.apache.poi.xssf.model.StylesTable;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFRichTextString;
import org.xml.sax.Attributes;
import org.xml.sax.InputSource;
import org.xml.sax.SAXException;
import org.xml.sax.XMLReader;
import org.xml.sax.helpers.DefaultHandler;
import org.xml.sax.helpers.XMLReaderFactory;

import java.io.InputStream;
import java.util.ArrayList;
import java.util.List;

/**
 * @author qjwyss
 * @date 2018/12/19
 * @description 讀取EXCEL輔助類
 */
public class ExcelXlsxReaderWithDefaultHandler extends DefaultHandler {

    private ExcelReadDataDelegated excelReadDataDelegated;

    public ExcelReadDataDelegated getExcelReadDataDelegated() {
        return excelReadDataDelegated;
    }

    public void setExcelReadDataDelegated(ExcelReadDataDelegated excelReadDataDelegated) {
        this.excelReadDataDelegated = excelReadDataDelegated;
    }

    public ExcelXlsxReaderWithDefaultHandler(ExcelReadDataDelegated excelReadDataDelegated) {
        this.excelReadDataDelegated = excelReadDataDelegated;
    }

    /**
     * 單元格中的資料可能的資料型別
     */
    enum CellDataType {
        BOOL, ERROR, FORMULA, INLINESTR, SSTINDEX, NUMBER, DATE, NULL
    }

    /**
     * 共享字串表
     */
    private SharedStringsTable sst;

    /**
     * 上一次的索引值
     */
    private String lastIndex;

    /**
     * 檔案的絕對路徑
     */
    private String filePath = "";

    /**
     * 工作表索引
     */
    private int sheetIndex = 0;

    /**
     * sheet名
     */
    private String sheetName = "";

    /**
     * 總行數
     */
    private int totalRows = 0;

    /**
     * 一行內cell集合
     */
    private List<String> cellList = new ArrayList<String>();

    /**
     * 判斷整行是否為空行的標記
     */
    private boolean flag = false;

    /**
     * 當前行
     */
    private int curRow = 1;

    /**
     * 當前列
     */
    private int curCol = 0;

    /**
     * T元素標識
     */
    private boolean isTElement;

    /**
     * 異常資訊，如果為空則表示沒有異常
     */
    private String exceptionMessage;

    /**
     * 單元格資料型別，預設為字串型別
     */
    private CellDataType nextDataType = CellDataType.SSTINDEX;

    private final DataFormatter formatter = new DataFormatter();

    /**
     * 單元格日期格式的索引
     */
    private short formatIndex;

    /**
     * 日期格式字串
     */
    private String formatString;

    //定義前一個元素和當前元素的位置，用來計算其中空的單元格數量，如A6和A8等
    private String preRef = null, ref = null;

    //定義該文件一行最大的單元格數，用來補全一行最後可能缺失的單元格
    private String maxRef = null;

    /**
     * 單元格
     */
    private StylesTable stylesTable;


    /**
     * 總行號
     */
    private Integer totalRowCount;

    /**
     * 遍歷工作簿中所有的電子表格
     * 並快取在mySheetList中
     *
     * @param filename
     * @throws Exception
     */
    public int process(String filename) throws Exception {
        filePath = filename;
        OPCPackage pkg = OPCPackage.open(filename);
        XSSFReader xssfReader = new XSSFReader(pkg);
        stylesTable = xssfReader.getStylesTable();
        SharedStringsTable sst = xssfReader.getSharedStringsTable();
        XMLReader parser = XMLReaderFactory.createXMLReader("org.apache.xerces.parsers.SAXParser");
        this.sst = sst;
        parser.setContentHandler(this);
        XSSFReader.SheetIterator sheets = (XSSFReader.SheetIterator) xssfReader.getSheetsData();
        while (sheets.hasNext()) { //遍歷sheet
            curRow = 1; //標記初始行為第一行
            sheetIndex++;
            InputStream sheet = sheets.next(); //sheets.next()和sheets.getSheetName()不能換位置，否則sheetName報錯
            sheetName = sheets.getSheetName();
            InputSource sheetSource = new InputSource(sheet);
            parser.parse(sheetSource); //解析excel的每條記錄，在這個過程中startElement()、characters()、endElement()這三個函式會依次執行
            sheet.close();
        }
        return totalRows; //返回該excel檔案的總行數，不包括首列和空行
    }

    /**
     * 第一個執行
     *
     * @param uri
     * @param localName
     * @param name
     * @param attributes
     * @throws SAXException
     */
    @Override
    public void startElement(String uri, String localName, String name, Attributes attributes) throws SAXException {

        // 獲取總行號  格式： A1:B5    取最後一個值即可
        if("dimension".equals(name)) {
            String dimensionStr = attributes.getValue("ref");
            totalRowCount = Integer.parseInt(dimensionStr.substring(dimensionStr.indexOf(":") + 2)) - 1;
        }

        //c => 單元格
        if ("c".equals(name)) {
            //前一個單元格的位置
            if (preRef == null) {
                preRef = attributes.getValue("r");
            } else {
                preRef = ref;
            }

            //當前單元格的位置
            ref = attributes.getValue("r");
            //設定單元格型別
            this.setNextDataType(attributes);
        }

        //當元素為t時
        if ("t".equals(name)) {
            isTElement = true;
        } else {
            isTElement = false;
        }

        //置空
        lastIndex = "";
    }


    /**
     * 第二個執行
     * 得到單元格對應的索引值或是內容值
     * 如果單元格型別是字串、INLINESTR、數字、日期，lastIndex則是索引值
     * 如果單元格型別是布林值、錯誤、公式，lastIndex則是內容值
     *
     * @param ch
     * @param start
     * @param length
     * @throws SAXException
     */
    @Override
    public void characters(char[] ch, int start, int length) throws SAXException {
        lastIndex += new String(ch, start, length);
    }


    /**
     * 第三個執行
     *
     * @param uri
     * @param localName
     * @param name
     * @throws SAXException
     */
    @Override
    public void endElement(String uri, String localName, String name) throws SAXException {

        //t元素也包含字串
        if (isTElement) {//這個程式沒經過
            //將單元格內容加入rowlist中，在這之前先去掉字串前後的空白符
            String value = lastIndex.trim();
            cellList.add(curCol, value);
            curCol++;
            isTElement = false;
            //如果裡面某個單元格含有值，則標識該行不為空行
            if (value != null && !"".equals(value)) {
                flag = true;
            }
        } else if ("v".equals(name)) {
            //v => 單元格的值，如果單元格是字串，則v標籤的值為該字串在SST中的索引
            String value = this.getDataValue(lastIndex.trim(), "");//根據索引值獲取對應的單元格值
            //補全單元格之間的空單元格
            if (!ref.equals(preRef)) {
                int len = countNullCell(ref, preRef);
                for (int i = 0; i < len; i++) {
                    cellList.add(curCol, "");
                    curCol++;
                }
            }
            cellList.add(curCol, value);
            curCol++;
            //如果裡面某個單元格含有值，則標識該行不為空行
            if (value != null && !"".equals(value)) {
                flag = true;
            }
        } else {
            //如果標籤名稱為row，這說明已到行尾，呼叫optRows()方法
            if ("row".equals(name)) {
                //預設第一行為表頭，以該行單元格數目為最大數目
                if (curRow == 1) {
                    maxRef = ref;
                }
                //補全一行尾部可能缺失的單元格
                if (maxRef != null) {
                    int len = countNullCell(maxRef, ref);
                    for (int i = 0; i <= len; i++) {
                        cellList.add(curCol, "");
                        curCol++;
                    }
                }

                if (flag && curRow != 1) { //該行不為空行且該行不是第一行，則傳送（第一行為列名，不需要）
                    // 呼叫excel讀資料委託類進行讀取插入操作
                    excelReadDataDelegated.readExcelDate(sheetIndex, totalRowCount, curRow, cellList);
                    totalRows++;
                }

                cellList.clear();
                curRow++;
                curCol = 0;
                preRef = null;
                ref = null;
                flag = false;
            }
        }
    }

    /**
     * 處理資料型別
     *
     * @param attributes
     */
    public void setNextDataType(Attributes attributes) {
        nextDataType = CellDataType.NUMBER; //cellType為空，則表示該單元格型別為數字
        formatIndex = -1;
        formatString = null;
        String cellType = attributes.getValue("t"); //單元格型別
        String cellStyleStr = attributes.getValue("s"); //
        String columnData = attributes.getValue("r"); //獲取單元格的位置，如A1,B1

        if ("b".equals(cellType)) { //處理布林值
            nextDataType = CellDataType.BOOL;
        } else if ("e".equals(cellType)) {  //處理錯誤
            nextDataType = CellDataType.ERROR;
        } else if ("inlineStr".equals(cellType)) {
            nextDataType = CellDataType.INLINESTR;
        } else if ("s".equals(cellType)) { //處理字串
            nextDataType = CellDataType.SSTINDEX;
        } else if ("str".equals(cellType)) {
            nextDataType = CellDataType.FORMULA;
        }

        if (cellStyleStr != null) { //處理日期
            int styleIndex = Integer.parseInt(cellStyleStr);
            XSSFCellStyle style = stylesTable.getStyleAt(styleIndex);
            formatIndex = style.getDataFormat();
            formatString = style.getDataFormatString();
            if (formatString.contains("m/d/yy") || formatString.contains("yyyy/mm/dd") || formatString.contains("yyyy/m/d")) {
                nextDataType = CellDataType.DATE;
                formatString = "yyyy-MM-dd hh:mm:ss";
            }

            if (formatString == null) {
                nextDataType = CellDataType.NULL;
                formatString = BuiltinFormats.getBuiltinFormat(formatIndex);
            }
        }
    }

    /**
     * 對解析出來的資料進行型別處理
     *
     * @param value   單元格的值，
     *                value代表解析：BOOL的為0或1， ERROR的為內容值，FORMULA的為內容值，INLINESTR的為索引值需轉換為內容值，
     *                SSTINDEX的為索引值需轉換為內容值， NUMBER為內容值，DATE為內容值
     * @param thisStr 一個空字串
     * @return
     */
    @SuppressWarnings("deprecation")
    public String getDataValue(String value, String thisStr) {
        switch (nextDataType) {
            // 這幾個的順序不能隨便交換，交換了很可能會導致資料錯誤
            case BOOL: //布林值
                char first = value.charAt(0);
                thisStr = first == '0' ? "FALSE" : "TRUE";
                break;
            case ERROR: //錯誤
                thisStr = "\"ERROR:" + value.toString() + '"';
                break;
            case FORMULA: //公式
                thisStr = '"' + value.toString() + '"';
                break;
            case INLINESTR:
                XSSFRichTextString rtsi = new XSSFRichTextString(value.toString());
                thisStr = rtsi.toString();
                rtsi = null;
                break;
            case SSTINDEX: //字串
                String sstIndex = value.toString();
                try {
                    int idx = Integer.parseInt(sstIndex);
                    XSSFRichTextString rtss = new XSSFRichTextString(sst.getEntryAt(idx));//根據idx索引值獲取內容值
                    thisStr = rtss.toString();
                    rtss = null;
                } catch (NumberFormatException ex) {
                    thisStr = value.toString();
                }
                break;
            case NUMBER: //數字
                if (formatString != null) {
                    thisStr = formatter.formatRawCellContents(Double.parseDouble(value), formatIndex, formatString).trim();
                } else {
                    thisStr = value;
                }
                thisStr = thisStr.replace("_", "").trim();
                break;
            case DATE: //日期
                thisStr = formatter.formatRawCellContents(Double.parseDouble(value), formatIndex, formatString);
                // 對日期字串作特殊處理，去掉T
                thisStr = thisStr.replace("T", " ");
                break;
            default:
                thisStr = " ";
                break;
        }
        return thisStr;
    }

    public int countNullCell(String ref, String preRef) {
        //excel2007最大行數是1048576，最大列數是16384，最後一列列名是XFD
        String xfd = ref.replaceAll("\\d+", "");
        String xfd_1 = preRef.replaceAll("\\d+", "");

        xfd = fillChar(xfd, 3, '@', true);
        xfd_1 = fillChar(xfd_1, 3, '@', true);

        char[] letter = xfd.toCharArray();
        char[] letter_1 = xfd_1.toCharArray();
        int res = (letter[0] - letter_1[0]) * 26 * 26 + (letter[1] - letter_1[1]) * 26 + (letter[2] - letter_1[2]);
        return res - 1;
    }

    public String fillChar(String str, int len, char let, boolean isPre) {
        int len_1 = str.length();
        if (len_1 < len) {
            if (isPre) {
                for (int i = 0; i < (len - len_1); i++) {
                    str = let + str;
                }
            } else {
                for (int i = 0; i < (len - len_1); i++) {
                    str = str + let;
                }
            }
        }
        return str;
    }


}