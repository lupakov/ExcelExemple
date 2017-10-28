package com.google.lupakov;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.openxml4j.exceptions.OpenXML4JException;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.xssf.eventusermodel.ReadOnlySharedStringsTable;
import org.apache.poi.xssf.eventusermodel.XSSFReader;
import org.apache.poi.xssf.eventusermodel.XSSFSheetXMLHandler;
import org.apache.poi.xssf.eventusermodel.XSSFSheetXMLHandler.SheetContentsHandler;
import org.apache.poi.xssf.model.StylesTable;
import org.apache.poi.xssf.usermodel.XSSFComment;
import org.xml.sax.SAXException;
import org.xml.sax.XMLReader;
 
 
 
import javax.xml.parsers.ParserConfigurationException;
import javax.xml.parsers.SAXParserFactory;
import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
 
/**
 * 
 */
public class Ny {
 
    OPCPackage pk;
    ArrayList<String> list = new ArrayList<>();
 
    public void openXlsx()
    {
 
        try {
            pk=OPCPackage.open("/home/vladimir/Документы/pot/somedoc.xlsx");
            ReadOnlySharedStringsTable strings = new ReadOnlySharedStringsTable(pk);
            XSSFReader xssfReader = new XSSFReader(pk);
            StylesTable styles = xssfReader.getStylesTable();
            XSSFReader.SheetIterator iter = (XSSFReader.SheetIterator) xssfReader.getSheetsData();
            while(iter.hasNext())
            {
                InputStream stream = iter.next();
 
                processSheet(styles,strings,stream);
                stream.close();
            }
 
            System.out.println(list);
 
 
        } catch (InvalidFormatException e) {
            e.printStackTrace();
            System.out.println("Ошибка при открытии файла");
        } catch (SAXException e) {
            e.printStackTrace();
            System.out.println("Ошибка SAX");
        } catch (IOException e) {
            e.printStackTrace();
            System.out.println("Ошибка IO при создании ReadOnlySharedStringsTable");
        } catch (OpenXML4JException e) {
            e.printStackTrace();
            System.out.println("Ошибка при создании XSSFReader");
        }
 
 
    }
 
    public void processSheet(StylesTable styles, ReadOnlySharedStringsTable strings, InputStream sheetInputstream)
    {
        org.xml.sax.InputSource sheetSource = new org.xml.sax.InputSource(sheetInputstream);
        SAXParserFactory saxFactory = SAXParserFactory.newInstance();
 
        try {
 
            javax.xml.parsers.SAXParser saxParser = saxFactory.newSAXParser();
            XMLReader sheetParser = saxParser.getXMLReader();
 
            org.xml.sax.ContentHandler handler = new XSSFSheetXMLHandler(styles, strings, new SheetContentsHandler() {
                @Override
                public void startRow(int i) {
 
                }
 
                @Override
                public void endRow(int i) {
 
 
                }
 
                @Override
                public void cell(String s, String s1, XSSFComment xssfComment) {
 
                    if (s.equals("A1")) list.add(s1);
 
                }
 
                @Override
                public void headerFooter(String s, boolean b, String s1) {
 
 
                }
            } , false);
 
            sheetParser.setContentHandler(handler);
            sheetParser.parse(sheetSource);
 
 
 
 
 
 
 
        } catch (ParserConfigurationException e) {
            e.printStackTrace();
        } catch (SAXException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        }
 
    }
 
 
}
