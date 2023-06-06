package utils;

import java.io.FileInputStream;
import java.io.IOException;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.WorkbookFactory;

public class Utility {

	
	public static String getDataFromExcel(String sheetName, int row, int cell) throws EncryptedDocumentException, IOException {
		String data = "no Data" ;
		FileInputStream file = new FileInputStream("D:\\AUOMATION ALL\\SS\\Screenshots\\TC123.xlsx");
		
		Cell getCell = WorkbookFactory.create(file).getSheet(sheetName).getRow(row).getCell(cell);
		Cell getCell = WorkbookFactory.create(file).getSheet(sheetName).getRow(row).getCell(cell);
		try {
			data = getCell.getStringCellValue();
		}
		catch(IllegalStateException e)
		{
			double value = getCell.getNumericCellValue();
			data = String.valueOf(value);
		}
		catch(Exception e)
		{
			e.printStackTrace();
		}
		
		System.out.println(data);
		return data;
	}
}
