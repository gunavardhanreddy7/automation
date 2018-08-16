import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.sl.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

public class w2 {
	
	   public static void main(String [] args) throws EncryptedDocumentException, InvalidFormatException, IOException {
	    Workbook workbook = WorkbookFactory.create(new File("xam4.xls"));

	    
	   Sheet sheet = (Sheet) workbook.getSheetAt(0);

	  
	    Row row = ((org.apache.poi.ss.usermodel.Sheet) sheet).getRow(1);
	        Cell cell = row.getCell(2);

	  
	    if (cell == null)
	        cell = row.createCell(2);

	  
	    cell.setCellType(CellType.STRING);
	    cell.setCellValue("Updated Value");

	    
	    FileOutputStream	fileOut = new FileOutputStream("xam4.xls");
		
	    try {
			workbook.write(fileOut);
		} catch (IOException e) {
			
			e.printStackTrace();
		}
	    try {
			fileOut.close();
		} catch (IOException e) {
			
			e.printStackTrace();
		}

	    
	    workbook.close();
	}
}
