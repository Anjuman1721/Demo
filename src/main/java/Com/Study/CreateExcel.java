package Com.Study;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

public class CreateExcel 
{
	static Workbook wb;
	/*String path="../Study/Excel/mm.xls";
	String sheetName="Demo";*/

	public  static void create(String path) throws EncryptedDocumentException, InvalidFormatException, IOException
	{

		wb = new HSSFWorkbook();
		FileOutputStream fileOut = new FileOutputStream(path);
		wb.write(fileOut);
		fileOut.close();
	}
	
	public static void createSheet1(String path,String SheetName) throws IOException
	{
		Sheet sheet = wb.createSheet(SheetName);
		FileOutputStream fileOut = new FileOutputStream(path);
		wb.write(fileOut);
		fileOut.close();
	}
	
	public static void createrows(String SheetName, String path) throws IOException
	{
		Sheet sheet = wb.createSheet(SheetName);
		CreationHelper createHelper = wb.getCreationHelper();
		Row row1= sheet.createRow((short)0);
		row1.createCell(0).setCellValue(createHelper.createRichTextString("ID"));
		row1.createCell(1).setCellValue(createHelper.createRichTextString("Result"));
		row1.createCell(2).setCellValue(createHelper.createRichTextString("Actual result"));
		FileOutputStream fileOut = new FileOutputStream(path);
		wb.write(fileOut);
		fileOut.close();
	}
	
}
