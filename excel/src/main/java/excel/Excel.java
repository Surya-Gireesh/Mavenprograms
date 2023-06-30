package excel;
import java.io.FileInputStream;
import java.io.IOException;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Excel {
	XSSFSheet sheet;
	Cell cell;
	Row row;
	public Excel() throws IOException
	{
		FileInputStream file=new FileInputStream("C:\\Users\\AJITH\\Documents\\EXAMP.xlsx");
				XSSFWorkbook workbook= new XSSFWorkbook(file);
		sheet=workbook.getSheet("Sheet2");
		
		}
	public String getData(int i,int j)
	{
		row=sheet.getRow(i);
		cell=row.getCell(j);
		CellType type=cell.getCellType();
		switch(type)
		{
		case NUMERIC:
			double data=cell.getNumericCellValue();
			return String.valueOf(data);
		case STRING:
			return cell.getStringCellValue();
		}
		return "invalid";
	
	}
	

public static void main(String args[]) throws IOException
{ 
	Excel e=new Excel();
	for(int i=0;i<2;i++)
	{
		for(int j=0;j<2;j++)
		{
			System.out.println(e.getData(i, j)+ " ");
		}
		System.out.print("");
		}
	}



}

