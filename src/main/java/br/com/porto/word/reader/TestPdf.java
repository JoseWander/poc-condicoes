package br.com.porto.word.reader;

import com.itextpdf.awt.geom.Rectangle;
import com.itextpdf.text.Chunk;
import com.itextpdf.text.Document;
import com.itextpdf.text.Paragraph;
import com.itextpdf.text.pdf.*;
import com.itextpdf.text.pdf.draw.*;

import java.io.*;

// WORKS CORRECTLY USING itext version 5.5.5
// FAILS WITH 5.5.6
// CAUSES AN EXCEPTION 
// "com.itextpdf.text.pdf.PdfDictionary cannot be cast to com.itextpdf.text.pdf.PdfArray"
// with makeRemoteNamedDestinationsLocal()
public class TestPdf {
    public static void main (String[] args) throws Exception {
        // Create simple document
        ByteArrayOutputStream main = new ByteArrayOutputStream();
        Document doc = new Document();
        PdfWriter pdfwrite = PdfWriter.getInstance(doc,main);
        doc.open();
        doc.add(new Paragraph("Testing Page"));
        doc.close();

        // Create TOC document
        ByteArrayOutputStream two = new ByteArrayOutputStream();
        Document doc2 = new Document();
        PdfWriter pdfwrite2 = PdfWriter.getInstance(doc2,two);      
        doc2.open();
        Chunk chn = new Chunk("<<-- Link To Testing Page -->>");
        chn.setRemoteGoto("DUMMY.PDF","page-num-1");
        doc2.add(new Paragraph(chn));
        doc2.close();

        // Merge documents
        ByteArrayOutputStream three = new ByteArrayOutputStream();
        PdfReader reader1 = new PdfReader("C:\\Porto\\POC_DOCS\\produto.pdf");
        PdfReader reader2 = new PdfReader("C:\\Porto\\POC_DOCS\\servicos.pdf");
        Document doc3 = new Document();
        PdfCopy DocCopy = new PdfCopy(doc3,three);
        doc3.open();
        DocCopy.addPage(DocCopy.getImportedPage(reader2,1));
        DocCopy.addPage(DocCopy.getImportedPage(reader1,1));
        DocCopy.addNamedDestination("page-num-1",2,new PdfDestination(PdfDestination.FIT));
        doc3.close();

        // Fix references and write to file
        PdfReader finalReader = new PdfReader(three.toByteArray());
        // Fails on this line
        finalReader.makeRemoteNamedDestinationsLocal();
        PdfStamper stamper = new PdfStamper(finalReader,new FileOutputStream("C:\\Porto\\POC_DOCS\\Testing.pdf"));
        stamper.close();    
    }
}