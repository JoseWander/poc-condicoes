package br.com.porto.word.reader;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.math.BigInteger;

import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.apache.poi.xwpf.usermodel.XWPFStyle;
import org.apache.poi.xwpf.usermodel.XWPFStyles;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTDecimalNumber;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTOnOff;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTPPr;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTString;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTStyle;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STStyleType;



public class Test2 {

	public static void main(String[] args)  throws IOException {
		
		FileInputStream servicosLista, servicosHeader;   
		 FileOutputStream saida;   
		 
	        try {            
	   	
	        	saida = new FileOutputStream(new File("C:\\Porto\\POC_DOCS\\CASE2\\saida.docx"));
	        	

	            XWPFDocument xSaida =new XWPFDocument(); 
	            
	    		
	    	    String strStyleId = "TitleWander";
	    	
	    	    addCustomHeadingStyle(xSaida, strStyleId, 1);
	    	
	    	    XWPFParagraph paragraph = xSaida.createParagraph();
	    	    XWPFRun run = paragraph.createRun();
	    	    run.setText("ola");
	    	    paragraph.setNumID(BigInteger.valueOf(1));
	    	
	    	    paragraph.setStyle(strStyleId);

	    	    
	    	    strStyleId = "TitleWander2";
		    	
	    	    addCustomHeadingStyle(xSaida, strStyleId, 2);
	    	
	    	    paragraph = xSaida.createParagraph();
	    	    run = paragraph.createRun();
	    	    run.setText("ola1");
	    	    paragraph.setNumID(BigInteger.valueOf(20));
	    	
	    	    paragraph.setStyle(strStyleId);
	    	    
	            
	        xSaida.write(saida);
	        saida.flush();        
	        saida.close();

		    
	          
           

	            System.out.println("fim");
	            
	            //Conversor.converte("C:\\Porto\\POC_DOCS\\saida");
	            
	           
	        }
	        catch(FileNotFoundException e){
	            e.printStackTrace();
	        }
	        catch(IOException e){
	            e.printStackTrace();
	        }
		
	}


	
	private static void addCustomHeadingStyle(XWPFDocument docxDocument, String strStyleId, int headingLevel) {
	
	    CTStyle ctStyle = CTStyle.Factory.newInstance();
	    ctStyle.setStyleId(strStyleId);
	    
	    ctStyle.
	
	    CTString styleName = CTString.Factory.newInstance();
	    styleName.setVal(strStyleId);
	    ctStyle.setName(styleName);
	
	    CTDecimalNumber indentNumber = CTDecimalNumber.Factory.newInstance();
	    indentNumber.setVal(BigInteger.valueOf(headingLevel));
	
	    // lower number > style is more prominent in the formats bar
	    ctStyle.setUiPriority(indentNumber);
	
	    CTOnOff onoffnull = CTOnOff.Factory.newInstance();
	    ctStyle.setUnhideWhenUsed(onoffnull);
	
	    // style shows up in the formats bar
	    ctStyle.setQFormat(onoffnull);
	
	    // style defines a heading of the given level
	    CTPPr ppr = CTPPr.Factory.newInstance();
	    ppr.setOutlineLvl(indentNumber);
	    ctStyle.setPPr(ppr);
	
	    XWPFStyle style = new XWPFStyle(ctStyle);
	
	    // is a null op if already defined
	    XWPFStyles styles = docxDocument.createStyles();
	
	    style.setType(STStyleType.PARAGRAPH);
	    styles.addStyle(style);
	
	}

}