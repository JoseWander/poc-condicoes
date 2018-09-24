package br.com.porto.word.reader;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.math.BigInteger;
import java.util.List;

import org.apache.poi.hpsf.Section;
import org.apache.poi.hwpf.HWPFDocument;
import org.apache.poi.hwpf.extractor.WordExtractor;
import org.apache.poi.hwpf.usermodel.Range;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.wp.usermodel.CharacterRun;
import org.apache.poi.wp.usermodel.Paragraph;
import org.apache.poi.xwpf.usermodel.TOC;
import org.apache.poi.xwpf.usermodel.XWPFAbstractNum;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFNumbering;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.xmlbeans.XmlCursor;
import org.docx4j.math.CTOnOff;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.docx4j.XmlUtils;
import org.docx4j.convert.out.pdf.viaXSLFO.PdfSettings;
import org.docx4j.dml.chart.CTStyle;
import org.docx4j.fonts.IdentityPlusMapper;
import org.docx4j.fonts.Mapper;
import org.docx4j.fonts.PhysicalFont;
import org.docx4j.fonts.PhysicalFonts;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.docx4j.openpackaging.parts.WordprocessingML.MainDocumentPart;
import org.docx4j.wml.CTDecimalNumber;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTAbstractNum;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTNumbering;

import com.itextpdf.text.Document;
import com.itextpdf.text.pdf.PdfWriter;

public class TestsPoi {
	
	public static void main(String[] args) throws InvalidFormatException {
		 FileInputStream servicosLista, servicosHeader;   
		 FileOutputStream saida;   
		 
	        try {            
	   	
	        	saida = new FileOutputStream(new File("C:\\Porto\\POC_DOCS\\CASE2\\saida1.docx"));
	        	

	            XWPFDocument xSaida =new XWPFDocument(); 
	            
	            String [] fruits = {"Apple", "Banana", "mango", "guava", "pear", "mellon" };
	            // for each item a paragraph is created and the Style and NumId is set
	            for (int i = 0; i < fruits.length; i++) {
	              
	                XWPFParagraph p = xSaida.createParagraph();
	                p.setStyle("ListParagraph");
	                
	                
	                
	                // good to see the XML structure
	                //System.out.println(p.getCTP());
	                
	            	  XWPFRun r = p.createRun();
	            	  String item=fruits[i];
	            	  r.setText(item);
	            	  p.setNumID(BigInteger.valueOf(1));
	            	  if(item.equals("mango")){
	            		  
	            		  //tab
	            		  p.getCTP().getPPr().getNumPr().setNil();
	            		  //p.getCTP().getPPr().getNumPr().addNewIlvl().setVal(BigInteger.valueOf(1));
	            	  }
	            	  else {
	            		 
	            		 // p.getCTP().getPPr().getNumPr().addNewIlvl().setVal(BigInteger.valueOf(1));
	            	  }
	            	      	                
	            	System.out.println();
	            	  
	            }  

		   
	            
	        xSaida.write(saida);
	        saida.flush();        
	        saida.close();

            
	           
	        }
	        catch(FileNotFoundException e){
	            e.printStackTrace();
	        }
	        catch(IOException e){
	            e.printStackTrace();
	        }
	}
	
	
	protected XWPFDocument doc;     
	private BigInteger addListStyle(String style)
	{
	    try
	    {
	        XWPFNumbering numbering = doc.getNumbering();
	        // generate numbering style from XML
	        CTAbstractNum abstractNum = CTAbstractNum.Factory.parse(style);
	        XWPFAbstractNum abs = new XWPFAbstractNum(abstractNum, numbering);

	        // find available id in document
	        BigInteger id = BigInteger.valueOf(0);
	        boolean found = false;
	        while (!found)
	        {
	            Object o = numbering.getAbstractNum(id);
	            found = (o == null);
	            if (!found) id = id.add(BigInteger.ONE);
	        }
	        // assign id
	        abs.getAbstractNum().setAbstractNumId(id);
	        // add to numbering, should get back same id
	        id = numbering.addAbstractNum(abs);
	        // add to num list, result is numid
	        return doc.getNumbering().addNum(id);           
	    }
	    catch (Exception e)
	    {
	        e.printStackTrace();
	        return null;
	    }
	}
}
