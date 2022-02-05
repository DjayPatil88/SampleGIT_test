package xlsbDemo;

import java.io.IOException;
import java.io.InputStream;

import org.apache.poi.ooxml.extractor.POIXMLTextExtractor;
import org.apache.poi.openxml4j.exceptions.OpenXML4JException;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.util.POILogFactory;
import org.apache.poi.util.POILogger;
import org.apache.poi.xssf.binary.XSSFBCommentsTable;
import org.apache.poi.xssf.binary.XSSFBHyperlinksTable;
import org.apache.poi.xssf.binary.XSSFBSharedStringsTable;
import org.apache.poi.xssf.binary.XSSFBSheetHandler;
import org.apache.poi.xssf.binary.XSSFBStylesTable;
import org.apache.poi.xssf.eventusermodel.XSSFBReader;
import org.apache.poi.xssf.eventusermodel.XSSFSheetXMLHandler.SheetContentsHandler;
import org.apache.poi.xssf.extractor.XSSFEventBasedExcelExtractor;
import org.apache.poi.xssf.model.SharedStrings;
import org.apache.poi.xssf.usermodel.XSSFRelation;
import org.apache.xmlbeans.XmlException;
import org.xml.sax.SAXException;

/**
 * Implementation of a text extractor or xlsb Excel
 * files that uses SAX-like binary parsing.
 *
 * @since 3.16-beta3
 */
public class TestXLSB extends XSSFEventBasedExcelExtractor
        implements org.apache.poi.ss.extractor.ExcelExtractor {

    private static final POILogger LOGGER = POILogFactory.getLogger(TestXLSB.class);

    public static final XSSFRelation[] SUPPORTED_TYPES = new XSSFRelation[]{
            XSSFRelation.XLSB_BINARY_WORKBOOK
    };

    private boolean handleHyperlinksInCells;

    public TestXLSB(String path) throws XmlException, OpenXML4JException, IOException {
        super(path);
    }

    public TestXLSB(OPCPackage container) throws XmlException, OpenXML4JException, IOException {
        super(container);
    }

    public static void main(String[] args) throws Exception {
    
    System.out.println("Changes in GIT");
        if (args.length < 1) {
            System.err.println("Use:");
            System.err.println("  XSSFBEventBasedExcelExtractor <filename.xlsb>");
            System.exit(1);
        }
        POIXMLTextExtractor extractor =
                new TestXLSB(args[0]);
        System.out.println(extractor.getText());
        extractor.close();
    }

    public void setHandleHyperlinksInCells(boolean handleHyperlinksInCells) {
        this.handleHyperlinksInCells = handleHyperlinksInCells;
    }

    /**
     * Should we return the formula itself, and not
     * the result it produces? Default is false
     * This is currently unsupported for xssfb
     */
    @Override
    public void setFormulasNotResults(boolean formulasNotResults) {
        throw new IllegalArgumentException("Not currently supported");
    }

    /**
     * Processes the given sheet
     */
    public void processSheet(
            SheetContentsHandler sheetContentsExtractor,
            XSSFBStylesTable styles,
            XSSFBCommentsTable comments,
            SharedStrings strings,
            InputStream sheetInputStream)
            throws IOException {

        DataFormatter formatter;
        if (getLocale() == null) {
            formatter = new DataFormatter();
        } else {
            formatter = new DataFormatter(getLocale());
        }

        XSSFBSheetHandler xssfbSheetHandler = new XSSFBSheetHandler(
                sheetInputStream,
                styles, comments, strings, sheetContentsExtractor, formatter, getFormulasNotResults()
        );
        xssfbSheetHandler.parse();
    }

    /**
     * Processes the file and returns the text
     */
    public String getText() {
        try {
            XSSFBSharedStringsTable strings = new XSSFBSharedStringsTable(getPackage());
            XSSFBReader xssfbReader = new XSSFBReader(getPackage());
            XSSFBStylesTable styles = xssfbReader.getXSSFBStylesTable();
            XSSFBReader.SheetIterator iter = (XSSFBReader.SheetIterator) xssfbReader.getSheetsData();

            StringBuilder text = new StringBuilder(64);
            SheetTextExtractor sheetExtractor = new SheetTextExtractor();
            XSSFBHyperlinksTable hyperlinksTable = null;
            while (iter.hasNext()) {
                InputStream stream = iter.next();
                if (getIncludeSheetNames()) {
                    text.append(iter.getSheetName());
                    text.append('\n');
                }
                if (handleHyperlinksInCells) {
                    hyperlinksTable = new XSSFBHyperlinksTable(iter.getSheetPart());
                }
                XSSFBCommentsTable comments = getIncludeCellComments() ? iter.getXSSFBSheetComments() : null;
                processSheet(sheetExtractor, styles, comments, strings, stream);
                if (getIncludeHeadersFooters()) {
                    sheetExtractor.appendHeaderText(text);
                }
                sheetExtractor.appendCellText(text);
                if (getIncludeTextBoxes()) {
                    processShapes(iter.getShapes(), text);
                }
                if (getIncludeHeadersFooters()) {
                    sheetExtractor.appendFooterText(text);
                }
                sheetExtractor.reset();
                stream.close();
            }

            return text.toString();
        } catch (IOException | OpenXML4JException | SAXException e) {
            LOGGER.log(POILogger.WARN, e);
            return null;
        }
    }

}
