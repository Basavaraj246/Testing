import java.io.FileInputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Iterator;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.util.NumberToTextConverter;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class dataDriven {

		 

		 
				// TODO Auto-generated method stub
		  
			public static ArrayList getData(String testCaseName) throws IOException 
			{	
				ArrayList<String> a =new ArrayList<String>();
            FileInputStream fis=new FileInputStream("C://Users//patil//Documents//demoData.xlsx");
             XSSFWorkbook workbook=new XSSFWorkbook(fis);
				
				int sheets=workbook.getNumberOfSheets();
				
				for(int i=0;i<sheets;i++)
				{
					if(workbook.getSheetName(i).equalsIgnoreCase("testData"))
					{
						XSSFSheet sheet=workbook.getSheetAt(i);//till here we got the first sheet
						//identify testcase row by iterating all 
		               Iterator<Row> rows=sheet.rowIterator();
		               
		               Row firstrow=rows.next();//firstrow has all the first row info
		               
		               Iterator<Cell> ce=firstrow.cellIterator();//now ce has ability to move to respective cell in row
		               int coulmn=0;
		               int k=0;
		               while(ce.hasNext())
		               {
		            	   Cell value=ce.next();
		            	   
		            	   if(value.getStringCellValue().equalsIgnoreCase("TestCases"))
		            	   {
		            		   coulmn=k;
		            	   }
		            	   k++;
		               }
		               System.out.println(coulmn);
		               while(rows.hasNext())
		               {
		            	   Row r=rows.next();
		            	   if(r.getCell(coulmn).getStringCellValue().equalsIgnoreCase(testCaseName))
		            	   {
		            		   Iterator<Cell> cv=r.cellIterator();
		            		   while(cv.hasNext())
		            		   {
		            			  //System.out.println(cv.next().getStringCellValue()); 
		            			   Cell c=cv.next();
		            			   if(c.getCellType()==CellType.STRING)
		            			   {
		            			  //a.add(cv.next().getStringCellValue());//
		            			  a.add(c.getStringCellValue());
		            			   }
		            			   else
		            			   {
		            				   //a.add(NumberToTextConverter.toText(c.getNumericCellValue()));
		            				   //a.add(c.getNumericCellValue());
		            				  a.add(NumberToTextConverter.toText(c.getNumericCellValue()));           
		            				  //NumberToTextConverter.toText(c.getNumericCellValue());
		            			   }
		            		   }
		            	    }
		               }
					}
					
					
					
				}
				return a;
				
				
			}

		

	

           public static void main(String[] args) throws IOException {
	       // TODO Auto-generated method stub
        	   
}
}

