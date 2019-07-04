package im.zhaojun.excel.analysis;

import im.zhaojun.excel.context.EasyExcelContext;
import im.zhaojun.excel.exception.EasyExcelException;
import im.zhaojun.excel.metadata.Sheet;
import org.apache.poi.openxml4j.exceptions.OpenXML4JException;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.xssf.eventusermodel.XSSFReader;
import org.apache.poi.xssf.model.SharedStringsTable;
import org.apache.poi.xssf.model.StylesTable;
import org.xml.sax.InputSource;
import org.xml.sax.XMLReader;
import org.xml.sax.helpers.DefaultHandler;

import javax.xml.parsers.SAXParser;
import javax.xml.parsers.SAXParserFactory;
import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.List;

public class XlsxSaxAnalyser extends DefaultHandler {

    // 单元格的格式表, 对应 style.xml
    private StylesTable stylesTable;

    //共享字符串表
    private SharedStringsTable sharedStringsTable;

    private List<SheetSource> sheetSourceList = new ArrayList<>();

    private EasyExcelContext easyExcelContext;

    public XlsxSaxAnalyser(EasyExcelContext easyExcelContext) throws IOException, OpenXML4JException {
        this.easyExcelContext = easyExcelContext;

        OPCPackage pkg = OPCPackage.open(easyExcelContext.getInputStream());

        // 获取解析器
        XSSFReader xssfReader = new XSSFReader(pkg);

        // 获取 共享字符串表 和 单元格样式表.
        this.stylesTable = xssfReader.getStylesTable();
        this.sharedStringsTable = xssfReader.getSharedStringsTable();

        XSSFReader.SheetIterator ite = (XSSFReader.SheetIterator) xssfReader.getSheetsData();
        while (ite.hasNext()) {
            InputStream inputStream = ite.next();
            String sheetName = ite.getSheetName();
            SheetSource sheetSource = new SheetSource(sheetName, inputStream);
            sheetSourceList.add(sheetSource);
        }
    }

    public void execute() {
        Sheet currentSheet = easyExcelContext.getCurrentSheet();
        InputStream sheetInputStream = sheetSourceList.get(currentSheet.getSheetNo() - 1).getInputStream();
        parseXmlSource(sheetInputStream);
    }

    private void parseXmlSource(InputStream inputStream) {
        try {
            InputSource inputSource = new InputSource(inputStream);
            // 防止 XEE 漏洞攻击.
            SAXParserFactory saxFactory = SAXParserFactory.newInstance();
            saxFactory.setFeature("http://apache.org/xml/features/disallow-doctype-decl", true);
            saxFactory.setFeature("http://xml.org/sax/features/external-general-entities", false);
            saxFactory.setFeature("http://xml.org/sax/features/external-parameter-entities", false);
            SAXParser saxParser = saxFactory.newSAXParser();
            XMLReader parser = saxParser.getXMLReader();
            parser.setContentHandler(new XlsxRowHandler(easyExcelContext, sharedStringsTable, stylesTable));
            parser.parse(inputSource);
        } catch (Exception e) {
            e.printStackTrace();
            throw new EasyExcelException(e);
        }

    }

    class SheetSource {

        private String sheetName;

        private InputStream inputStream;

        public SheetSource(String sheetName, InputStream inputStream) {
            this.sheetName = sheetName;
            this.inputStream = inputStream;
        }

        public String getSheetName() {
            return sheetName;
        }

        public void setSheetName(String sheetName) {
            this.sheetName = sheetName;
        }

        public InputStream getInputStream() {
            return inputStream;
        }

        public void setInputStream(InputStream inputStream) {
            this.inputStream = inputStream;
        }
    }

}