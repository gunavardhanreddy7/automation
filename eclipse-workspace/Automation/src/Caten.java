


import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;


public class Caten {
  
public static void main(String [] args) throws FileNotFoundException {
	 

	  
       try   { 
    	   FileInputStream is = new FileInputStream(new File("xam4.xlsx"));
    	   Workbook workbook = WorkbookFactory.create(is);
    	   Sheet s= workbook.getSheetAt(0);
    	 
           
              for(int i=1;i<4;i++) {
            	Row r2=s.getRow(i-1); 
            	Cell cat,cat1,cat2;
            	if(r2.getCell(0)==null) {
            		 cat=r2.createCell(0);
            	}
            	else {
            		 cat=r2.getCell(0);
            	}
            	if(r2.getCell(2)==null) {
           		 cat1=r2.createCell(2);
           	}
           	else {
           		 cat2=r2.getCell(2);
           	}
            	if(r2.getCell(4)==null) {
           		 cat=r2.createCell(4);
           	}
           	else {
           		 cat=r2.getCell(4);
           	}
            	

            	Cell sr=r2.getCell(0);
            	Cell sr1=r2.getCell(2);
            	 String a=sr.toString();
            	 String a2=sr1.toString();
                 
            	cat.setCellValue(a+a2);
       
       /*     String sam="I"+i;
             String sum="J"+i;
             Row r21=s.getRow(i-1);
             Cell c=r21.getCell(11);
			c.setCellFormula(CONCATENATE(sam,sum));
         */   
              
             
       }
           is.close();
           workbook.write(new FileOutputStream("xam4.xlsx"));  
           workbook.close();
           
       }catch(Exception e) {  
           System.out.println(e.getMessage());  
       }  
	   
   }



	}


