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

public class dataDriven2 {
	
	public ArrayList<String> dataEnter(String TCname)  throws IOException
	{     
		
		//Defining the arraylist
		ArrayList <String> a=new ArrayList<String>();
		FileInputStream fis=new FileInputStream("C:\\Users\\ROYAL COMPUTER\\Desktop\\demo2.xlsx");
		XSSFWorkbook wrkbook=new XSSFWorkbook(fis);
		int number=wrkbook.getNumberOfSheets();
		for(int i=0;i<number;i++)
		{
			if(wrkbook.getSheetName(i).equalsIgnoreCase("credentials"))
			{
				XSSFSheet sheet=wrkbook.getSheetAt(i);
				Iterator<Row> rows=sheet.iterator();
				Row firstrow=rows.next();
				Iterator<Cell> cv=firstrow.iterator();
				int k=0;
				int column=0;
				while(cv.hasNext())
				{   Cell value=cv.next();
					if(value.getStringCellValue().equalsIgnoreCase("username"))
					{
						column=k;
					}
					k++;
				}
				System.out.println("Column number is "+column);
				while(rows.hasNext())
				{
					Row r=rows.next();
					if(r.getCell(column).getStringCellValue().equalsIgnoreCase(TCname))
					{
						Iterator<Cell> ce=r.cellIterator();
						while(ce.hasNext())
							
						{
							Cell c=ce.next();
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
