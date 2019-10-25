package scripts;

import java.io.FileInputStream;
import java.io.FileOutputStream;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.ss.usermodel.Row.MissingCellPolicy;
import org.testng.Reporter;

public class FWUtils 
{
	static Workbook wb;
	static FileInputStream fis;	
	static String path="./src/main/resources/data/Data.xlsx";
	static String sheet="Sheet1";
	static Sheet sh;
	static Row r;
	static Cell c;
	public static String readData(int row, int cell)
	{
		String value="";
		try
		{
			fis=new FileInputStream(path);
			wb=WorkbookFactory.create(fis);
			sh=wb.getSheet(sheet);
			r=sh.getRow(row);
			if(sh.getRow(row)==null)
			{
				r=sh.createRow(row);
			}
			c=r.getCell(cell,MissingCellPolicy.CREATE_NULL_AS_BLANK);
			value=c.toString();
		}
		catch(Exception e)
		{
			Reporter.log(e.getMessage(),true);
		}
		return value;
	}
	
	public static void writeData(int row, int cell, String data)
	{
		try
		{
			fis = new FileInputStream(path);
			wb=WorkbookFactory.create(fis);
			r=wb.getSheet(sheet).getRow(row);
			if(r==null)
			{
				r=wb.getSheet(sheet).createRow(row);
			}
			c=r.createCell(cell);
			c.setCellValue(data);
			wb.write(new FileOutputStream(path));
		}
		catch(Exception e)
		{
			Reporter.log(e.getMessage(),true);
		}
	}
}
