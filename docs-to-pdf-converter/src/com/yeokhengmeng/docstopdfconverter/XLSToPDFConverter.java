package com.yeokhengmeng.docstopdfconverter;

import java.io.InputStream;
import java.io.OutputStream;
import java.util.Iterator;

import com.lowagie.text.pdf.PdfCell;
import com.lowagie.text.pdf.PdfTable;
import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import com.itextpdf.text.*;
import com.itextpdf.text.pdf.*;

public class XLSToPDFConverter extends Converter {

    public XLSToPDFConverter(InputStream inStream, OutputStream outStream, boolean showMessages, boolean closeStreamsWhenComplete) {
        super(inStream, outStream, showMessages, closeStreamsWhenComplete);
    }

    @Override
    public void convert() throws Exception {
        loading();

        HSSFWorkbook wb = new HSSFWorkbook(inStream);
        HSSFSheet wsh = wb.getSheetAt(0);
        Iterator<Row> rowIterator = wsh.rowIterator();

        Document doc = new Document();
        PdfWriter.getInstance(doc, outStream);
        doc.open();
        PdfTable table = new PdfTable(columns);
        PdfCell pdfCell;
        while (rowIterator.hasNext()) {
            Row row = rowIterator.next();
            Iterator<Cell> cellIterator = row.cellIterator();
            while (cellIterator.hasNext()) {
                Cell cell = cellIterator.next();
                switch (cell.getCellType()) {
                    case Cell.CELL_TYPE_STRING:
                        pdfCell = new PdfPCell(new Phrase(cell.getStringCellValue()));
                        table.addCell(cell);
                        break;
                }
            }
        }
        finished();
        inStream.close();
        doc.close();
    }
}
