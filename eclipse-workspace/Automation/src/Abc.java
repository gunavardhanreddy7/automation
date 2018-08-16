
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.OutputStream;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

public class Abc {

	public static void main(String [] args) {
		@SuppressWarnings("resource")
		Workbook wb = new HSSFWorkbook();  

		try  (OutputStream fileOut = new FileOutputStream("xam1.xls")) { 
			FileInputStream is = new FileInputStream(new File("xam.xls"));
			Workbook workbook = WorkbookFactory.create(is);
			Sheet sheet = workbook.getSheetAt(0);
			Row r1=sheet.createRow(7);
			Cell c1=r1.createCell(8);


			Sheet s=  wb.createSheet("First Sheet");  
			wb.createSheet("Second Sheet");
			wb.createSheet("Third Sheet");

			for(int i=1;i<5;i++) {
				Row row     = s.createRow(i-1);  
				Cell cell   = row.createCell(7);  
				Cell cell1   = row.createCell(8);  
				String s1="F"+i;
				String s2="+G"+i;
				String s3="*G"+i;
				cell.setCellFormula(s1+s2);
				cell1.setCellFormula(s1+s3);
			}

			for(int i=5;i<10;i++) {
				Row r2=s.createRow(i-1); 
				Cell cat,cat1,cat2;
				if(r2.getCell(11)==null) {
					cat=r2.createCell(11);
				}
				else {
					cat=r2.getCell(11);
				}
				if(r2.getCell(9)==null) {
					cat1=r2.createCell(9);
				}
				else {
					cat2=r2.getCell(9);
				}
				if(r2.getCell(10)==null) {
					cat=r2.createCell(10);
				}
				else {
					cat=r2.getCell(10);
				}


				Cell sr=r2.getCell(9);
				Cell sr1=r2.getCell(10);
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


			wb.write(fileOut);  

		}catch(Exception e) {  
			System.out.println(e.getMessage());  
		}  

	}
}
