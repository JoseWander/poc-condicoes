package br.com.porto.word.reader;

import java.io.FileOutputStream;
import java.io.IOException;
import java.sql.SQLException;
 

import com.itextpdf.text.Document;
import com.itextpdf.text.DocumentException;
import com.itextpdf.text.pdf.PdfCopy;
import com.itextpdf.text.pdf.PdfReader;
 

 
public class Concatenate {
 
 
    /**
     * Main method.
     * @param    args    no arguments needed
     * @throws DocumentException 
     * @throws IOException
     * @throws SQLException
     */
    public static void main(String[] args)
        throws IOException, DocumentException, SQLException {
        // using previous examples to create PDFs
    
        String[] files = { "C:\\Porto\\POC_DOCS\\template.pdf", "C:\\Porto\\POC_DOCS\\produto.pdf","C:\\Porto\\POC_DOCS\\servicos.pdf" };
        // step 1
        Document document = new Document();
        // step 2
        PdfCopy copy = new PdfCopy(document, new FileOutputStream("C:\\Porto\\POC_DOCS\\saida.pdf"));
        // step 3
        document.open();
        // step 4
        PdfReader reader;
        int n;
        // loop over the documents you want to concatenate
        for (int i = 0; i < files.length; i++) {
            reader = new PdfReader(files[i]);
            // loop over the pages in that document
            n = reader.getNumberOfPages();
            for (int page = 0; page < n; ) {
                copy.addPage(copy.getImportedPage(reader, ++page));
            }
            copy.freeReader(reader);
            reader.close();
        }
        // step 5
        document.close();
        
        System.out.println("fim");
    }
}