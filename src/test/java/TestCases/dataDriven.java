package TestCases;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
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

	
	public ArrayList<String> getData(String TestCaseName) throws IOException {

		// TODO Auto-generated method stub
		//fileinputstream
		
		ArrayList<String> a=new ArrayList<String>();
		
		FileInputStream fis=new FileInputStream("C:\\Users\\ROYAL COMPUTER\\Documents\\datademo.xlsx");
        XSSFWorkbook workbook=new XSSFWorkbook(fis);
      int sheets=  workbook.getNumberOfSheets();
      for(int i=0;i<sheets;i++)
      {   if(workbook.getSheetName(i).equalsIgnoreCase("Sheet1"))
      {
    	  
      
    	  XSSFSheet sheet=workbook.getSheetAt(i);  
    	  Iterator<Row>  rows=sheet.iterator();//Identify TestCASES COLUMN by scanning entire 1st row.
    	  
    	   Row firstrow= rows.next();
    	  Iterator<Cell> ce=firstrow.cellIterator();//Read each and every cell of first row
    	  int k=0;
    	  int column=0; 
    	while(ce.hasNext())    //This step won't check but just move the cell value
    	{
    		Cell value=ce.next();
    		if(value.getStringCellValue().equalsIgnoreCase("Testcases"))
    		{
    		column=k;
    		}
    		k++;
    	}
    	
    	System.out.println(column);
    	while(rows.hasNext())
    	{
    	Row r= rows.next();
    	if(r.getCell(column).getStringCellValue().equalsIgnoreCase(TestCaseName))
    	{
    		//desired row
    		//pull all your data of the row into the test
    		Iterator<Cell> cv=r.cellIterator();
    		while(cv.hasNext()) 
    		{
    			Cell c=cv.next();
    			if(c.getCellTypeEnum()==CellType.STRING)
    			{
    		a.add(c.getStringCellValue());
    		}
    			else
    			{
    				a.add(NumberToTextConverter.toText(c.getNumericCellValue()));
    				
    			}
    			
    	}
      }
	}
	}}
      return a;
	}
	public static void main(String[] args) throws IOException {}
	{
		
	}
}
