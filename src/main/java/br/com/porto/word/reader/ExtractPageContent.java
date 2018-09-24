package br.com.porto.word.reader;


import java.io.ByteArrayOutputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.PrintWriter;

import com.itextpdf.text.Document;
import com.itextpdf.text.DocumentException;
import com.itextpdf.text.Paragraph;
import com.itextpdf.text.pdf.PdfContentByte;
import com.itextpdf.text.pdf.PdfReader;
import com.itextpdf.text.pdf.PdfStamper;
import com.itextpdf.text.pdf.PdfWriter;
import com.itextpdf.text.pdf.parser.PdfReaderContentParser;
import com.itextpdf.text.pdf.parser.SimpleTextExtractionStrategy;
import com.itextpdf.text.pdf.parser.TextExtractionStrategy;
 
public class ExtractPageContent {
 
    /** The original PDF that will be parsed. */
    public static final String PDF = "C:\\Porto\\POC_DOCS\\servicos.pdf";
    /** The resulting text file. */
    public static final String PDF2 = "C:\\Porto\\POC_DOCS\\servicos2.pdf";
    
    public static final String PDF3 = "C:\\Porto\\POC_DOCS\\servicos3.pdf";
 
    /**
     * Parses a PDF to a plain text file.
     * @param pdf the original PDF
     * @param txt the resulting text
     * @throws IOException
     * @throws DocumentException 
     */
    public void readPDF() throws IOException, DocumentException {
    	PdfReader reader;
    	ByteArrayOutputStream baos;
    	PdfStamper stamper;
    	    	
    	reader = new PdfReader(PDF);
    	 
    	reader.selectPages(String.valueOf(0));
    	baos = new ByteArrayOutputStream();
    	stamper = new PdfStamper(reader, baos);
    	stamper.close();
    	
    	Document document = new Document();
    	PdfWriter.getInstance(document, baos);
    	document.open();
    	document.add(new Paragraph("Hello World!"));
    	FileOutputStream fos = new FileOutputStream(PDF3,true);
    	fos.write(baos.toByteArray());
    	 
    /*	reader = new PdfReader(PDF2);
    	reader.selectPages(String.valueOf(0));
    	baos = new ByteArrayOutputStream();
    	stamper = new PdfStamper(reader, baos);
    	
    	PdfContentByte underContent = stamper. 
    	
    	stamper.close();
    	
    	PdfWriter.getInstance(document, baos);
    	fos.write(baos.toByteArray()); */
    	
    	fos.close();
    	reader.close();
            
    }
 
    /**
     * Main method.
     * @param    args    no arguments needed
     * @throws IOException
     */
    public static void main(String[] args) throws IOException, DocumentException {
    	//new ExtractPageContent().readPDF(PDF);
    	new ExtractPageContent().readPDF();
        //new ExtractPageContent().parsePdf(PREFACE, RESULT);
    }
}