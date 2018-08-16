package concat;


import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.OutputStream;

import org.apache.poi.hssf.util.CellReference;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

public class Join {
  
public static void main(String [] args) throws FileNotFoundException {
	 
	  
       try  (OutputStream fileOut = new FileOutputStream("xam2.xls")) { 
    	 
    	 
  /*
        
            wb.createSheet("Second Sheet");
             wb.createSheet("Third Sheet");
    */         
    	  
    	   Workbook wb1 = WorkbookFactory.create(new File("xam2.xls"));
    	  
    	   Sheet s =wb1.getSheetAt(0);
    	   
             for(int i=1;i<=5;i++) {
            	 Row row     = s.getRow(i-1);  
                 Cell cell   = row.getCell(4);
            	 
            	  CellReference c1 = new CellReference(i-1, 0) ;
            	  CellReference c2 = new CellReference(i-1, 2) ;
            	  
                  
				String thisR = c1.getCellRefParts()[1]; 
                  String thisR1 = c2.getCellRefParts()[1]; 
                  
                  cell.setCellFormula("SUM(" + thisR + "" + thisR1 + ")");
            	
                  
            
             }
             
      //     wb1.write(fileOut);  
        //   fileOut.close();
           
           
       }catch(Exception e) {  
           System.out.println(e.getMessage());  
       }  
	   
   }
}


