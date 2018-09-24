package br.com.porto.word.reader;

import java.io.FileNotFoundException;
import java.io.OutputStream;
import java.util.HashMap;
import java.util.Iterator;
import java.util.Map;

import javax.xml.bind.JAXBException;

import org.docx4j.XmlUtils;
import org.docx4j.convert.out.pdf.viaXSLFO.PdfSettings;
import org.docx4j.fonts.IdentityPlusMapper;
import org.docx4j.fonts.Mapper;
import org.docx4j.fonts.PhysicalFont;
import org.docx4j.fonts.PhysicalFonts;
import org.docx4j.openpackaging.exceptions.Docx4JException;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.docx4j.openpackaging.parts.WordprocessingML.MainDocumentPart;


public class Conversor {


	public static void main(String[] args) {
			
		converte("C:\\Porto\\POC_DOCS\\CASE2\\template");
		/*
		converte("C:\\Porto\\POC_DOCS\\CASE1\\cobertura_2");
		converte("C:\\Porto\\POC_DOCS\\CASE1\\cobertura_3");
		converte("C:\\Porto\\POC_DOCS\\CASE1\\servico_1");
		converte("C:\\Porto\\POC_DOCS\\CASE1\\servico_2");
		converte("C:\\Porto\\POC_DOCS\\CASE1\\produto");
		*/
	}
	
	static void converte(String inputfilepath){
		
		WordprocessingMLPackage wordMLPackage;
		
		try {
			wordMLPackage = WordprocessingMLPackage.load(new java.io.File(inputfilepath + ".docx"));

			org.docx4j.convert.out.pdf.PdfConversion c
		 = new org.docx4j.convert.out.pdf.viaXSLFO.Conversion(wordMLPackage);

			((org.docx4j.convert.out.pdf.viaXSLFO.Conversion)c).setSaveFO(new java.io.File(inputfilepath+".fo"));
			
			
				((org.docx4j.convert.out.pdf.viaXSLFO.Conversion)c).setSaveFO(
						new java.io.File(inputfilepath + ".fo"));
				OutputStream os;
				
					os = new java.io.FileOutputStream(inputfilepath + ".pdf");
						
				c.output(os, new PdfSettings() );
				System.out.println("Saved " + inputfilepath + ".pdf");
			 
			
		} catch (Docx4JException e1) {
			// TODO Auto-generated catch block
			e1.printStackTrace();
		} catch (FileNotFoundException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}	
	}
	

}