package concat;

import java.io.FileInputStream;
import java.io.IOException;
import java.util.Scanner;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

public class Search {

	public static void main(String[] args) throws EncryptedDocumentException, InvalidFormatException, IOException {
		// TODO Auto-generated method stub
		@SuppressWarnings("resource")
		Scanner s=new Scanner(System.in);
		System.out.println("Enter sheet number");
           int sno=s.nextInt();
           
          
           
			FileInputStream fin=new FileInputStream("abc.xls");
		
		
           Workbook w=WorkbookFactory.create(fin);
           try {
             if(w.getSheetAt(sno-1)==null) {
            	 System.out.println("NO Sheet With given sheet no");
            	 System.exit(0);
            	 
             }
           }
           catch(Exception e){
        	   System.out.println("NO Sheet With given sheet no ");
        	   System.exit(0);
           }
             System.out.println("Enter row number");
             int rno=s.nextInt();
             Sheet sh=w.getSheetAt(sno-1);
             try {
            
             if(sh.getRow(rno-1)==null) {
            	 System.out.println("No data in regarding row no");
            	 System.exit(0);
             }
             }
             catch(Exception e){
          	   System.out.println("out of range exception");
          	   System.exit(0);
             }
             Row r=sh.getRow(rno-1);
             System.out.println("Enter cell number");
             int cno=s.nextInt();
             try {
             if(r.getCell(cno-1)==null) {
            	 System.out.println("No data in reagarding cell no");
            	 System.exit(0);
             }
             }
             catch(Exception e){
            	   System.out.println("out of range exception");
            	   System.exit(0);
               }
             Cell c=r.getCell(cno-1);
             
            
            
             String sd=c.toString();
             System.out.println("_______THE DATA FOUND __________");
             System.out.println("Value is ="+sd);
           
           
           
           
	}

}
