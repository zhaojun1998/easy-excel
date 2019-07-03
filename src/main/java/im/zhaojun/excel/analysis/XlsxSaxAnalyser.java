package im.zhaojun.excel.analysis;

import im.zhaojun.excel.annotation.EasyExcelField;
import im.zhaojun.excel.annotation.EasyExcelMapping;
import im.zhaojun.excel.annotation.EasyExcelMappings;
import im.zhaojun.excel.annotation.FieldType;
import im.zhaojun.excel.exception.NotSupportTypeException;
import im.zhaojun.excel.handler.ExcelRowHandler;
import im.zhaojun.excel.util.ExcelParseUtil;
import org.apache.poi.hssf.util.CellReference;
import org.apache.poi.openxml4j.exceptions.OpenXML4JException;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.xssf.eventusermodel.XSSFReader;
import org.apache.poi.xssf.model.SharedStringsTable;
import org.apache.poi.xssf.model.StylesTable;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.xml.sax.Attributes;
import org.xml.sax.InputSource;
import org.xml.sax.SAXException;
import org.xml.sax.XMLReader;
import org.xml.sax.helpers.DefaultHandler;

import javax.xml.parsers.ParserConfigurationException;
import javax.xml.parsers.SAXParser;
import javax.xml.parsers.SAXParserFactory;
import java.io.IOException;
import java.io.InputStream;
import java.lang.reflect.Field;
import java.util.*;

public class XlsxSaxAnalyser extends DefaultHandler {

    private static final Logger log = LoggerFactory.getLogger(XlsxSaxAnalyser.class);

    /** Cell 单元格元素 */
    private static final String C_ELEMENT = "c";

    /** 行元素 */
    private static final String ROW_ELEMENT = "row";

    /** Cell 中的行列号 */
    private static final String R_ATTR = "r";

    /** SST (SharedStringsTable) 的索引 */
    private static final String S_ATTR_VALUE = "s";

    // 列中属性值
    private static final String T_ATTR_VALUE = "t";

    // sheet r:Id前缀
    private static final String RID_PREFIX = "rId";

    // 单元格的格式表, 对应 style.xml
    private StylesTable stylesTable;

    //共享字符串表
    private SharedStringsTable sharedStringsTable;

    // 上一次的内容
    private String curContent;

    private Integer lastRowNum;

    private Short lastCellNum;

    private FieldType curFieldType;

    private List<Object> rowCellList = new ArrayList<>();

    private ExcelRowHandler excelRowHandler;

    private Class<?> clz;

    private int startRow;

    // 缓存 Excel 列号和字段的 Map 关系
    private Map<Integer, Field> fieldMap;

    private Field[] fields;

    // 字段映射 Map<fieldName, Map<key, value>>
    private Map<String, Map<String, String>> fieldMapping;

    public XlsxSaxAnalyser(ExcelRowHandler excelRowHandler, Class<?> clz) {
        this.excelRowHandler = excelRowHandler;
        this.clz = clz;
        startRow = ExcelParseUtil.parseStartRow(clz);

        init();
    }


    private void init() {
        fields = clz.getDeclaredFields();
        fieldMap = getFieldMap();
        fieldMapping = getFieldMapping();
    }


    public void processOneSheet(InputStream inputStream, int sheetId) {
        InputStream sheetInputStream = null;
        OPCPackage pkg = null;
        try {
            pkg = OPCPackage.open(inputStream);

            // 获取解析器
            XSSFReader xssfReader = new XSSFReader(pkg);
            XMLReader xmlReader = fetchSheetParser();

            // 获取 共享字符串表 和 单元格样式表.
            this.stylesTable = xssfReader.getStylesTable();
            this.sharedStringsTable = xssfReader.getSharedStringsTable();
            // 根据 rId# 获取 sheet 页
            sheetInputStream = xssfReader.getSheet(RID_PREFIX + sheetId);
            InputSource sheetSource = new InputSource(sheetInputStream);
            xmlReader.parse(sheetSource);

        } catch (IOException | ParserConfigurationException | OpenXML4JException | SAXException e) {
            e.printStackTrace();
        } finally {
            try {
                if (sheetInputStream != null) {
                    sheetInputStream.close();
                }
            } catch (IOException e) {
                e.printStackTrace();
            }

            try {
                if (pkg != null) {
                    pkg.close();
                }
            } catch (IOException e) {
                e.printStackTrace();
            }
        }
    }


    private XMLReader fetchSheetParser()
            throws SAXException, ParserConfigurationException {

        // 防止 XEE 漏洞攻击.
        SAXParserFactory saxFactory = SAXParserFactory.newInstance();
        saxFactory.setFeature("http://apache.org/xml/features/disallow-doctype-decl", true);
        saxFactory.setFeature("http://xml.org/sax/features/external-general-entities", false);
        saxFactory.setFeature("http://xml.org/sax/features/external-parameter-entities", false);
        SAXParser saxParser = saxFactory.newSAXParser();
        XMLReader parser = saxParser.getXMLReader();
        parser.setContentHandler(this);
        return parser;
    }

    @Override
    public void startDocument() {
        excelRowHandler.before();
    }

    @Override
    public void startElement(String uri, String localName, String qName, Attributes attributes) {

        // 行开始, 记录行号
        if (ROW_ELEMENT.equals(qName)) {
            lastRowNum = Integer.parseInt(attributes.getValue(R_ATTR));
        }

        // 单元格元素
        if (C_ELEMENT.equals(qName)) {

            // 获取列数
            String cellRef = attributes.getValue(R_ATTR);
            CellReference cellReference = new CellReference(cellRef);
            short curCellNum = cellReference.getCol();
            curCellNum += 1;  // 这里得到的序号是 从 0 开始的.

            //空单元判断，添加空字符到list
            if (lastCellNum != null) {
                int gap = curCellNum - lastCellNum;
                for (int i = 0; i < gap - 1; i++) {
                    rowCellList.add(null);
                }
            } else {
                // 第一个单元格可能不是在第一列
                if (!"A1".equals(cellRef)) {
                    for (int i = 0; i < curCellNum - 1; i++) {
                        rowCellList.add(null);
                    }
                }
            }
            lastCellNum = curCellNum;

            setCellType(attributes);
        }
        curContent = "";
    }


    @Override
    public void characters(char[] ch, int start, int length) {
        curContent += new String(ch, start, length);
    }


    @Override
    public void endElement(String uri, String localName, String qName) {
        String contentStr = curContent.trim();

        // 如果是单元格元素
        if (C_ELEMENT.equals(qName)) {
            // cell 标签
            Object value = ExcelParseUtil.getDataValue(curFieldType, contentStr, sharedStringsTable);
            rowCellList.add(value);
        } else if (ROW_ELEMENT.equals(qName)) {
            lastCellNum = null;
            if (lastRowNum > startRow) {
                excelRowHandler.execute(convertToObject());
            }
            rowCellList.clear();
        }
    }

    @Override
    public void endDocument() {
        excelRowHandler.doAfterAll();
    }

    private void setCellType(Attributes attribute) {
        // 重置 numFmtIndex, numFmtString 的值
        // 单元格存储格式的索引, 对应 style.xml 中的 numFmts 元素的子元素索引
        int numFmtIndex;
        String numFmtString = "";
        this.curFieldType = FieldType.of(attribute.getValue(T_ATTR_VALUE));

        // 获取单元格的 xf 索引, 对应 style.xml 中 cellXfs 的子元素 xf 的第几个
        String xfIndexStr = attribute.getValue(S_ATTR_VALUE);
        // 判断是否为日期类型
        if (xfIndexStr != null) {
            int xfIndex = Integer.parseInt(xfIndexStr);
            XSSFCellStyle xssfCellStyle = stylesTable.getStyleAt(xfIndex);
            numFmtIndex = xssfCellStyle.getDataFormat();
            numFmtString = xssfCellStyle.getDataFormatString();

            if (numFmtString == null) {
                curFieldType = FieldType.EMPTY;
            } else if (org.apache.poi.ss.usermodel.DateUtil.isADateFormat(numFmtIndex, numFmtString)) {
                curFieldType = FieldType.DATE;
            }
        }
    }


    /**
     * 将类转化为业务类
     */
    private Object convertToObject() {
        Object obj = null;
        try {
            obj = clz.newInstance();

            for (Map.Entry<Integer, Field> fieldEntry : fieldMap.entrySet()) {
                Integer key = fieldEntry.getKey();
                Field field = fieldEntry.getValue();

                field.setAccessible(true);

                Object o = rowCellList.get(key);
                if (o != null) {
                    field.set(obj, parseValueWithFieldType(field, o));
                }
            }
        } catch (IllegalAccessException | InstantiationException e) {
            e.printStackTrace();
        }
        return obj;
    }


    private  <T> Map<Integer, Field> getFieldMap() {
        Map<Integer, Field> fieldMap = new HashMap<>();

        for (Field field : fields) {
            EasyExcelField easyExcelField = field.getAnnotation(EasyExcelField.class);
            if (easyExcelField != null) {
                fieldMap.put(easyExcelField.index(), field);
            }
        }
        return fieldMap;
    }


    private Map<String, Map<String, String>> getFieldMapping() {
        Map<String, Map<String, String>> fieldMapping = new HashMap<>();

        for (Field field : fields) {
            Map<String, String> map = new HashMap<>();

            EasyExcelMappings easyExcelMappings = field.getDeclaredAnnotation(EasyExcelMappings.class);

            if (easyExcelMappings != null) {

                EasyExcelMapping[] easyExcelMapping = easyExcelMappings.value();

                for (EasyExcelMapping excelMapping : easyExcelMapping) {
                    String key = excelMapping.key();
                    String value = excelMapping.value();
                    map.put(key, value);
                }
                fieldMapping.put(field.getName(), map);
            }
        }
        return fieldMapping;
    }


    /**
     * 将 Excel 列中的值, 转化成实体类的字段对应的数据类型
     * @param field     实体类的字段
     * @param obj       Excel 列的值
     * @return          转换后的值
     */
    private Object parseValueWithFieldType(Field field, Object obj) {
        Map<String, String> fieldMap = fieldMapping.get(field.getName());
        if (fieldMap != null) {
            obj = fieldMap.get(obj);
        }

        Class<?> type = field.getType();

        EasyExcelField easyExcelField = field.getDeclaredAnnotation(EasyExcelField.class);

        String format = easyExcelField.format();
        // 如果是日期类型, 或字符串类型, 但标注了格式化日志的字段, 则尝试转换成日期格式.
        if (Date.class.equals(type) && ExcelParseUtil.objIsString(obj)) {
            return ExcelParseUtil.parseDate(String.valueOf(obj), format);
        } else if (Date.class.equals(type)) {
            return obj;
        }

        return convertToBasicType(type, obj);
    }


    /**
     * 将数据转化为指定数据类型.
     */
    private Object convertToBasicType(Class<?> fieldType, Object obj) {
        if (Byte.class.equals(fieldType) || Byte.TYPE.equals(fieldType)) {
            return Byte.valueOf(ExcelParseUtil.convertString(obj));
        } else if (Boolean.class.equals(fieldType) || Boolean.TYPE.equals(fieldType)) {
            return Boolean.valueOf(ExcelParseUtil.convertString(obj));
        } else if (String.class.equals(fieldType)) {
            return ExcelParseUtil.convertString(obj);
        } else if (Short.class.equals(fieldType) || Short.TYPE.equals(fieldType)) {
            return Short.valueOf(ExcelParseUtil.convertString(obj));
        } else if (Integer.class.equals(fieldType) || Integer.TYPE.equals(fieldType)) {
            return Integer.valueOf(ExcelParseUtil.convertString(obj));
        } else if (Long.class.equals(fieldType) || Long.TYPE.equals(fieldType)) {
            return Long.valueOf(ExcelParseUtil.convertString(obj));
        } else if (Float.class.equals(fieldType) || Float.TYPE.equals(fieldType)) {
            return Float.valueOf(ExcelParseUtil.convertString(obj));
        } else if (Double.class.equals(fieldType) || Double.TYPE.equals(fieldType)) {
            return Double.valueOf(ExcelParseUtil.convertString(obj));
        } else {
            throw new NotSupportTypeException("Illegal data type: " + fieldType + ", value: " + obj);
        }
    }

}