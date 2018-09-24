package br.com.porto.word.reader;

import java.io.ByteArrayOutputStream;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;
import java.util.TreeMap;
 




import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;

import com.itextpdf.text.Chunk;
import com.itextpdf.text.Document;
import com.itextpdf.text.DocumentException;
import com.itextpdf.text.Element;
import com.itextpdf.text.Paragraph;
import com.itextpdf.text.Phrase;
import com.itextpdf.text.pdf.ColumnText;
import com.itextpdf.text.pdf.PdfAction;
import com.itextpdf.text.pdf.PdfAnnotation;
import com.itextpdf.text.pdf.PdfCopy;
import com.itextpdf.text.pdf.PdfCopy.PageStamp;
import com.itextpdf.text.pdf.PdfImportedPage;
import com.itextpdf.text.pdf.PdfReader;
import com.itextpdf.text.pdf.PdfStamper;
import com.itextpdf.text.pdf.draw.DottedLineSeparator;
 
public class MergeWithToc2 {
 

    public static final String SRC1 = "C:\\Porto\\POC_DOCS\\CASE1\\produto.pdf";
    public static final String SRC2 = "C:\\Porto\\POC_DOCS\\CASE1\\servico_1.pdf";
    public static final String SRC3 = "C:\\Porto\\POC_DOCS\\CASE1\\servico_2.pdf";
    public static final String SRC4 = "C:\\Porto\\POC_DOCS\\CASE1\\cobertura_1.pdf";
    public static final String SRC5 = "C:\\Porto\\POC_DOCS\\CASE1\\cobertura_2.pdf";
    public static final String SRC6 = "C:\\Porto\\POC_DOCS\\CASE1\\cobertura_3.pdf";
    public static final String TOC = "C:\\Porto\\POC_DOCS\\CASE1\\toc.pdf";
    public static final String DEST = "C:\\Porto\\POC_DOCS\\CASE1\\saida.pdf";
 
    public Map<String, PdfReader> filesToMerge;
 
    public static void main(String[] args) throws IOException, DocumentException {
    	setHeader("C:\\Porto\\POC_DOCS\\CASE1\\servico_1.docx", "2.1");
    	
        //File file = new File(DEST);
        //file.getParentFile().mkdirs();
        //MergeWithToc2 app = new MergeWithToc2();
        //app.createPdf(DEST);
        //System.out.println("fim");
    }
 
    public MergeWithToc2() throws IOException {
        filesToMerge = new LinkedHashMap<String, PdfReader>();  
        filesToMerge.put("Produto", new PdfReader(SRC1));
        filesToMerge.put("Serviço 1", new PdfReader(SRC2));
        filesToMerge.put("Serviço 2", new PdfReader(SRC3));
        filesToMerge.put("Cobertura 1", new PdfReader(SRC4));
        filesToMerge.put("Cobertura 2", new PdfReader(SRC5));
        filesToMerge.put("Cobertura 3", new PdfReader(SRC6));
        
    }
 
    public void createPdf(String filename) throws IOException, DocumentException {
        ByteArrayOutputStream baos = new ByteArrayOutputStream();
        Map<Integer, String> toc = new TreeMap<Integer, String>();
        Document document = new Document();
        PdfCopy copy = new PdfCopy(document, baos);
        PageStamp stamp;
        document.open();
        int n;
        int pageNo = 0;
        PdfImportedPage page;
        Chunk chunk;
        for (Map.Entry<String, PdfReader> entry : filesToMerge.entrySet()) {
            n = entry.getValue().getNumberOfPages();
            toc.put(pageNo + 1, entry.getKey());
            for (int i = 0; i < n; ) {
                pageNo++;
                page = copy.getImportedPage(entry.getValue(), ++i);
                stamp = copy.createPageStamp(page);
                chunk = new Chunk(String.format("Página %d", pageNo));
                if (i == 1)
                    chunk.setLocalDestination("p" + pageNo);
                ColumnText.showTextAligned(stamp.getUnderContent(),
                        Element.ALIGN_RIGHT, new Phrase(chunk),
                        559, 810, 0);
                stamp.alterContents();
                copy.addPage(page);
            }
        }
        PdfReader reader = new PdfReader(TOC);
        page = copy.getImportedPage(reader, 1);
        stamp = copy.createPageStamp(page);
        Paragraph p;
        PdfAction action;
        PdfAnnotation link;
        float y = 770;
        ColumnText ct = new ColumnText(stamp.getOverContent());
        ct.setSimpleColumn(36, 36, 559, y);
        for (Map.Entry<Integer, String> entry : toc.entrySet()) {
            p = new Paragraph(entry.getValue());
            p.add(new Chunk(new DottedLineSeparator()));
            p.add(String.valueOf(entry.getKey()));
            ct.addElement(p);
            ct.go();
            action = PdfAction.gotoLocalPage("p" + entry.getKey(), false);
            link = new PdfAnnotation(copy, 36, ct.getYLine(), 559, y, action);
            stamp.addAnnotation(link);
            y = ct.getYLine();
        }
        ct.go();
        stamp.alterContents();
        copy.addPage(page);
        document.close();
        for (PdfReader r : filesToMerge.values()) {
            r.close();
        }
        reader.close();
 
        reader = new PdfReader(baos.toByteArray());
        n = reader.getNumberOfPages();
        reader.selectPages(String.format("%d, 1-%d", n, n-1));
        PdfStamper stamper = new PdfStamper(reader, new FileOutputStream(filename));
        stamper.close();
    }
    
    static void setHeader(String file, String text){
    		 
   		 FileInputStream entrada;   
   		 FileOutputStream saida;   
   		 
   	        try {    
   	        	entrada = new FileInputStream(file);      	
   	        	saida = new FileOutputStream(new File(file + "saida.docx"));
   	        	
   	            XWPFDocument xEntrada=new XWPFDocument(OPCPackage.open(entrada));
   	            XWPFDocument xSaida =new XWPFDocument(); 
   	            
   	            xSaida=xEntrada;

   	            XWPFParagraph p=xSaida.getParagraphArray(0);
   	            String textOld=p.getText();
   	            
   	            System.out.println(p.getRuns().size() + ">" + textOld);
	            
   	            //removendo old runs
   	         	List<XWPFRun> runs = p.getRuns();
   	         	for(int i=runs.size()-1; i>0; i--) {
   	         		p.removeRun(i);
   	         	}
   	         	
   	         	XWPFRun r = p.getRuns().get(0);
   	         	r.setText(text + " " + textOld,0);
    	              	            
   	            xSaida.write(saida);
   	            saida.flush();        
   	            saida.close();

   	            entrada.close();
   	            
   	            System.out.println("fim");
   	        } catch( InvalidFormatException e){
   	        	
   	        }
   	        catch(FileNotFoundException e){
   	            e.printStackTrace();
   	        }
   	        catch(IOException e){
   	            e.printStackTrace();
   	        }
   	}
}