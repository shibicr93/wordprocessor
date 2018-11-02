package com.rfpio.wordprocessor;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.xwpf.usermodel.IBodyElement;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.xmlbeans.XmlException;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.HashMap;
import java.util.Iterator;
import java.util.List;
import java.util.Map;

/**
 * Created by shibi.c on 2/11/18.
 */
public class WordProcessor {

    private static final Map<String,Integer> styleCountMap = new HashMap<>();

    public static void main(String[] args) {

        try {
            ClassLoader classLoader = WordProcessor.class.getClassLoader();
            File file = new File(classLoader.getResource("rfpio.docx").getFile());

            FileInputStream fis = new FileInputStream(file);
            XWPFDocument xdoc = new XWPFDocument(OPCPackage.open(fis));

            Iterator bodyElementIterator = xdoc.getBodyElementsIterator();
            while (bodyElementIterator.hasNext()) {
                IBodyElement element = (IBodyElement) bodyElementIterator.next();

                if("PARAGRAPH".equalsIgnoreCase(element.getElementType().name())) {
                    updateStyleCount(element.getBody().getParagraphs());
                }

                if ("TABLE".equalsIgnoreCase(element.getElementType().name())) {
                    processTable(element.getBody().getTables());
                }
            }

            styleCountMap.forEach((k,v) -> {
                System.out.println("style : " +k + "\t" + "count : " + v);
            });

        } catch (InvalidFormatException e) {
            e.printStackTrace();
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        } catch (XmlException e) {
            e.printStackTrace();
        }

    }

    private static void processTable(List<XWPFTable> tableList) throws IOException, XmlException {
        for (XWPFTable table : tableList) {
            for (int i = 0; i < table.getRows().size(); i++) {
                for (int j = 0; j < table.getRow(i).getTableCells().size(); j++) {
                    if(ifExists(table,i,j)) {
                        processTable(table.getRow(i).getCell(j).getTables());
                    }else {
                        updateStyleCount(table.getRow(i).getCell(j).getParagraphs());
                    }
                }
            }
        }
    }

    private static void updateStyleCount(List<XWPFParagraph> paragraphs) {
        paragraphs.stream().forEach(paragraph -> {
            String style = paragraph.getStyle();
            if(style != null) {
                if(styleCountMap.get(style) == null) {
                    styleCountMap.put(style,1);
                } else {
                    int existingCount = styleCountMap.get(style);
                    int newCount = existingCount + 1;
                    styleCountMap.put(style,newCount);
                }
            }
        });
    }

    private static boolean ifExists(XWPFTable table, int i, int j) {
        return table != null && table.getRow(i) != null && table.getRow(i).getCell(j) != null && table.getRow(i).getCell(j).getTables() != null && table.getRow(i).getCell(j).getTables().size() > 0;
    }


}
