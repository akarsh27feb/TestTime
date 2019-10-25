package scripts;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.time.LocalDate;
import java.time.LocalTime;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.testng.annotations.Test;

public class TimeStamp_01 
{		
	Workbook wb;
	Sheet sh;
	Row r;
	Cell c;
	FileInputStream fis;
	String path="./src/main/resources/data/Data.xlsx";
	String sheet="Sheet1";
	
	public void writeData(int row, int cell, String data)
	{
		try
		{
			wb=WorkbookFactory.create(new FileInputStream(path));
			sh=wb.getSheet(sheet);
			r=sh.getRow(row);
			if(r==null)
				r=sh.createRow(row);
			c=r.createCell(cell);
			c.setCellValue(data);
			wb.write(new FileOutputStream(path));
		}
		catch(Exception e)
		{
			System.out.println(e.getMessage());
		}
	}
	
	@Test
	public void writeTest()
	{		
		int Day=LocalDate.now().getDayOfMonth();
		int Month=LocalDate.now().getMonthValue();
		int Year=LocalDate.now().getYear();
		int hour=LocalTime.now().getHour();
		int min=LocalTime.now().getMinute();
		String date=Day+"-"+Month+"-"+Year;
		String time=hour+":"+min;
		writeData(0, 0, "Date");
		writeData(0, 1, "Time");
		int writeRow=sh.getLastRowNum()+1;
		writeData(writeRow, 0, date);
		writeData(writeRow, 1, time);
	}
}
