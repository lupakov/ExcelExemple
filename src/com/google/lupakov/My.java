package com.google.lupakov;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.openxml4j.exceptions.OpenXML4JException;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.xssf.eventusermodel.ReadOnlySharedStringsTable;
import org.apache.poi.xssf.eventusermodel.XSSFReader;
import org.apache.poi.xssf.eventusermodel.XSSFSheetXMLHandler;
import org.apache.poi.xssf.model.StylesTable;
import org.apache.poi.xssf.usermodel.XSSFComment;
import org.xml.sax.ContentHandler;
import org.xml.sax.InputSource;
import org.xml.sax.SAXException;
import org.xml.sax.XMLReader;
import org.xml.sax.helpers.XMLReaderFactory;

import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;


public class My {
    OPCPackage pk = OPCPackage.open("/home/vladimir/Документы/pot/somedoc.xlsx");
    ArrayList<String> list = new ArrayList<>();


    public My() throws InvalidFormatException {


        XSSFReader xssfReader = null;
        try {
            xssfReader = new XSSFReader(pk);

        StylesTable styles = xssfReader.getStylesTable();
            ReadOnlySharedStringsTable strings = null;
            try {
                strings = new ReadOnlySharedStringsTable(pk);

            XSSFReader.SheetIterator iter = (XSSFReader.SheetIterator) xssfReader.getSheetsData();
                XMLReader parser = XMLReaderFactory.createXMLReader();
                
            while (iter.hasNext()) {

                    InputStream stream=iter.next();

                    ContentHandler handler = new XSSFSheetXMLHandler(styles, strings, new XSSFSheetXMLHandler.SheetContentsHandler() {
                        @Override
                        public void startRow(int i) {

                        }

                        @Override
                        public void endRow(int i) {

                        }

                        @Override
                        public void cell(String s, String s1, XSSFComment xssfComment) {

                            list.add(s1);


                        }

                        @Override
                        public void headerFooter(String s, boolean b, String s1) {

                        }
                    }, true);


                    parser.setContentHandler(handler);
                    parser.parse(new InputSource(stream));


                   

        }
            } catch (SAXException e) {
                e.printStackTrace();
            }

        } catch (IOException e) {
            e.printStackTrace();
        } catch (OpenXML4JException e) {
            e.printStackTrace();
        }


    }
}
