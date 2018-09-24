package br.com.porto.word.reader;

import java.io.FileOutputStream;
import java.io.IOException;
import java.sql.SQLException;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;

 
import com.itextpdf.text.Document;
import com.itextpdf.text.DocumentException;
import com.itextpdf.text.pdf.PdfCopy;
import com.itextpdf.text.pdf.PdfReader;
import com.itextpdf.text.pdf.SimpleBookmark;
 
public class ConcatenateBookmarks {
 
    /**
     * Manipulates a PDF file src with the file dest as result
     * @param src the original PDF
     * @param dest the resulting PDF
     * @throws IOException
     * @throws DocumentException
     */
    public void manipulatePdf(String[] src, String dest)
        throws IOException, DocumentException {
    	// step 1
        Document document = new Document();
        // step 2
        PdfCopy copy
            = new PdfCopy(document, new FileOutputStream(dest));
        // step 3
        document.open();
        // step 4
        PdfReader reader;
        int page_offset = 0;
        int n;
        // Create a list for the bookmarks
        ArrayList<HashMap<String, Object>> bookmarks = new ArrayList<HashMap<String, Object>>();
        List<HashMap<String, Object>> tmp;
        for (int i  = 0; i < src.length; i++) {
            reader = new PdfReader(src[i]);
            // merge the bookmarks
            tmp = SimpleBookmark.getBookmark(reader);
            SimpleBookmark.shiftPageNumbers(tmp, page_offset, null);
            bookmarks.addAll(tmp);
            // add the pages
            n = reader.getNumberOfPages();
            page_offset += n;
            for (int page = 0; page < n; ) {
                copy.addPage(copy.getImportedPage(reader, ++page));
            }
            copy.freeReader(reader);
            reader.close();
        }
        // Add the merged bookmarks
        copy.setOutlines(bookmarks);
        // step 5
        document.close();
    }
 
    /**
     * Main method.
     * @param    args    no arguments needed
     * @throws DocumentException 
     * @throws IOException
     * @throws SQLException
     */
    public static void main(String[] args)
        throws IOException, DocumentException, SQLException {
  
        new ConcatenateBookmarks().manipulatePdf(
            new String[]{ "C:\\Porto\\POC_DOCS\\template.pdf","C:\\Porto\\POC_DOCS\\template2.pdf"},
            "C:\\Porto\\POC_DOCS\\saida22.pdf");
    }
}