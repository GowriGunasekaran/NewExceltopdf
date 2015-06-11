package demo.exceltopdf;

import java.io.*;
import java.io.FileInputStream;

import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.ss.usermodel.*;
import java.util.Iterator;

import com.itextpdf.text.*;
import com.itextpdf.text.pdf.*;
 
public class PdfDemo {

	 public static void main(String[] args) throws Exception
	 {
	try{	 
		  
		FileInputStream input_document = new FileInputStream(new File("D:\\read\\Book1.xls"));
		XSSFWorkbook workbook = new XSSFWorkbook(input_document);
		XSSFSheet my_worksheet = workbook.getSheetAt(0);
		
	    Iterator<Row> rowIterator = my_worksheet.iterator();
	    
	    Document iText_xls_2_pdf = new Document();
	    PdfWriter.getInstance(iText_xls_2_pdf, new FileOutputStream("D:\\read\\at.pdf"));
	    iText_xls_2_pdf.open();
	    
	    PdfPTable my_table = new PdfPTable(4);
	    
	    
	    PdfPCell table_cell;
	    
	    while(rowIterator.hasNext()){
	    	Row row= rowIterator.next();
	    	Iterator<Cell> cellIterator = row.cellIterator();
	    	while(cellIterator.hasNext()){
	    		Cell cell = cellIterator.next();
	    		switch (cell.getCellType()) {
	    		case Cell.CELL_TYPE_STRING:
	    			table_cell=new PdfPCell(new Phrase(cell.getStringCellValue()));
	    			my_table.addCell(table_cell);
	    			break;
	    		}
	    	}
	    }
	    iText_xls_2_pdf.add(my_table);
	    iText_xls_2_pdf.close();
	    
	    input_document.close();
	    System.out.println("file created pdf");
	}
	catch(FileNotFoundException e)
	{
		e.printStackTrace();
		System.out.println("file not created pdf");
	}
	catch(IOException es)
	{
		es.printStackTrace();
	}
	}
	}
	 
		
	

 
