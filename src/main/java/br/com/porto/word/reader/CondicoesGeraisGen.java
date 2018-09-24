package br.com.porto.word.reader;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
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
import org.apache.poi.xwpf.usermodel.XWPFDocument;
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

import com.itextpdf.text.Document;
import com.itextpdf.text.pdf.PdfWriter;

public class CondicoesGeraisGen {
	
	static String repositorio ="C:\\Porto\\POC_DOCS\\CASE2\\";
	
	public static void main(String[] args) throws InvalidFormatException {
		coberturas();
		//servicos();
		//coberturasServicos() ;
		//servicosProduto() ;
		//produtoSaidaFinal() ;
		 
		//Conversor.converte(repositorio + "saidaFinal");

	}
	
	static void coberturas() throws InvalidFormatException 	{
	 
		 FileInputStream coberturasLista, coberturasHeader;   
		 FileOutputStream coberturasSaida;   
		 
	        try {    
	        	
	        	//COBETURAS
	        	coberturasLista = new FileInputStream(repositorio + "coberturasLista.docx"); 
	        	coberturasHeader = new FileInputStream(repositorio + "coberturasHeader.docx");        	
	        	coberturasSaida = new FileOutputStream(new File(repositorio + "coberturasSaida.docx"));
	        	
	            XWPFDocument xCoberturasLista=new XWPFDocument(OPCPackage.open(coberturasLista));
	            XWPFDocument xCoberturasHeader=new XWPFDocument(OPCPackage.open(coberturasHeader));
	            XWPFDocument xCoberturasSaida =new XWPFDocument(); 
	            
	            xCoberturasSaida=xCoberturasHeader;

	            List<XWPFParagraph> paraCoberturas = xCoberturasLista.getParagraphs();
	            boolean inicio=true;
	    
	            for(XWPFParagraph para : paraCoberturas){
	            	if(inicio){
	            		XWPFRun r = para.createRun();
	            		r.setText("Cobertura 1");
	            		xCoberturasSaida.setParagraph(para, xCoberturasSaida.getParagraphs().size()-1);
	            		
	            		inicio=false;
	            	}
	            	else{
	            		xCoberturasSaida.createParagraph();
	            		xCoberturasSaida.setParagraph(para, xCoberturasSaida.getParagraphs().size()-1);
	            	}
			    }
	            
	            xCoberturasSaida.write(coberturasSaida);
	            coberturasSaida.flush();        
	            coberturasSaida.close();

	            coberturasLista.close();
	            coberturasHeader.close();
	            
	            System.out.println("fim");
	        }
	        catch(FileNotFoundException e){
	            e.printStackTrace();
	        }
	        catch(IOException e){
	            e.printStackTrace();
	        }
	}
	
	static void servicos() throws InvalidFormatException 	{
		 FileInputStream servicosLista, servicosHeader;   
		 FileOutputStream servicosSaida;   
		 
	        try {    
	        	servicosLista = new FileInputStream(repositorio + "servicosLista.docx"); 
	        	servicosHeader = new FileInputStream(repositorio + "servicosHeader.docx");        	
	        	servicosSaida = new FileOutputStream(new File(repositorio + "servicosSaida.docx"));
	        	
	            XWPFDocument xServicosLista=new XWPFDocument(OPCPackage.open(servicosLista));
	            XWPFDocument xServicosHeader=new XWPFDocument(OPCPackage.open(servicosHeader));
	            XWPFDocument xServicosSaida =new XWPFDocument(); 
	            
	            xServicosSaida=xServicosHeader;

	            List<XWPFParagraph> paraServicos = xServicosLista.getParagraphs();
	            	    
	            for(XWPFParagraph para : paraServicos){
	            	xServicosSaida.createParagraph();
		            xServicosSaida.setParagraph(para, xServicosSaida.getParagraphs().size()-1);
			    }
	            
	            xServicosSaida.write(servicosSaida);
		        servicosSaida.flush();        
		        servicosSaida.close();

		        servicosLista.close();
		        servicosHeader.close();

	            System.out.println("fim");
	        }
	        catch(FileNotFoundException e){
	            e.printStackTrace();
	        }
	        catch(IOException e){
	            e.printStackTrace();
	        }
	}
	
	
	static void coberturasServicos() throws InvalidFormatException 	{
		 
		 FileInputStream coberturas, servicos;   
		 FileOutputStream coberturasServicosSaida;   
		 
	        try {    
	        	
	        	//COBETURAS
	        	servicos = new FileInputStream(repositorio + "servicosSaida.docx"); 
	        	coberturas = new FileInputStream(repositorio + "coberturasSaida.docx");        	
	        	coberturasServicosSaida = new FileOutputStream(new File(repositorio + "coberturasServicosSaida.docx"));
	        	
	            XWPFDocument xServicos=new XWPFDocument(OPCPackage.open(servicos));
	            XWPFDocument xCoberturas=new XWPFDocument(OPCPackage.open(coberturas));
	            XWPFDocument xCoberturasServicosSaida =new XWPFDocument(); 
	            
	            xCoberturasServicosSaida=xServicos;

	            List<XWPFParagraph> paraCoberturas = xCoberturas.getParagraphs();
	    
	            for(XWPFParagraph para : paraCoberturas){
	            	xCoberturasServicosSaida.createParagraph();
	            	xCoberturasServicosSaida.setParagraph(para, xCoberturasServicosSaida.getParagraphs().size()-1);
			    }
	            
	            xCoberturasServicosSaida.write(coberturasServicosSaida);
	            coberturasServicosSaida.flush();        
	            coberturasServicosSaida.close();

	            servicos.close();
	            coberturas.close();
	            
	            System.out.println("fim");
	        }
	        catch(FileNotFoundException e){
	            e.printStackTrace();
	        }
	        catch(IOException e){
	            e.printStackTrace();
	        }
	}
	
	
	static void servicosProduto() throws InvalidFormatException 	{
		 
		 FileInputStream coberturasServicosSaida, produto;   
		 FileOutputStream servicosProdutoSaida;   
		 
	        try {    
	        	
	        	//COBETURAS
	        	coberturasServicosSaida = new FileInputStream(repositorio + "coberturasServicosSaida.docx"); 
	        	produto = new FileInputStream(repositorio + "produto.docx");        	
	        	servicosProdutoSaida = new FileOutputStream(new File(repositorio + "servicosProdutoSaida.docx"));
	        	
	            XWPFDocument xProduto=new XWPFDocument(OPCPackage.open(produto));
	            XWPFDocument xCoberturasServicosSaida=new XWPFDocument(OPCPackage.open(coberturasServicosSaida));
	            XWPFDocument xServicosProdutoSaida =new XWPFDocument(); 
	            
	            xServicosProdutoSaida=xProduto;

	            List<XWPFParagraph> paraCoberturas = xCoberturasServicosSaida.getParagraphs();
	    
	            for(XWPFParagraph para : paraCoberturas){
	            	xServicosProdutoSaida.createParagraph();
	            	xServicosProdutoSaida.setParagraph(para, xServicosProdutoSaida.getParagraphs().size()-1);
			    }
	            
	            xServicosProdutoSaida.write(servicosProdutoSaida);
	            servicosProdutoSaida.flush();        
	            servicosProdutoSaida.close();

	            coberturasServicosSaida.close();
	            produto.close();
	            
	            System.out.println("fim");
	        }
	        catch(FileNotFoundException e){
	            e.printStackTrace();
	        }
	        catch(IOException e){
	            e.printStackTrace();
	        }
	}
	
	
	static void produtoSaidaFinal() throws InvalidFormatException 	{
		 
		 FileInputStream servicosProdutoSaida, indice;   
		 FileOutputStream saidaFinal;   
		 
	        try {    
	        	
	        	//COBETURAS
	        	servicosProdutoSaida = new FileInputStream(repositorio + "servicosProdutoSaida.docx"); 
	        	indice = new FileInputStream(repositorio + "indice.docx");        	
	        	saidaFinal = new FileOutputStream(new File(repositorio + "saidaFinal.docx"));
	        	
	            XWPFDocument xServicosProdutoSaida=new XWPFDocument(OPCPackage.open(servicosProdutoSaida));
	            XWPFDocument xIndice=new XWPFDocument(OPCPackage.open(indice));
	            XWPFDocument xSaidaFinal =new XWPFDocument(); 
	            
	            xSaidaFinal=xIndice;

	            List<XWPFParagraph> paraCoberturas = xServicosProdutoSaida.getParagraphs();
	    
	            for(XWPFParagraph para : paraCoberturas){
	            	xSaidaFinal.createParagraph();
	            	xSaidaFinal.setParagraph(para, xSaidaFinal.getParagraphs().size()-1);
			    }
	            
	            xSaidaFinal.write(saidaFinal);
	            saidaFinal.flush();        
	            saidaFinal.close();

	            servicosProdutoSaida.close();
	            indice.close();
	            
	            System.out.println("fim");
	        }
	        catch(FileNotFoundException e){
	            e.printStackTrace();
	        }
	        catch(IOException e){
	            e.printStackTrace();
	        }
	}
}
