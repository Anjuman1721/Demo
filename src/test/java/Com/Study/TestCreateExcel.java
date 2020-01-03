package Com.Study;

import java.io.IOException;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.testng.annotations.BeforeMethod;
import org.testng.annotations.BeforeTest;
import org.testng.annotations.Test;

public class TestCreateExcel 
{
	//CreateExcel createexcel;
   @Test
   public void Create() throws EncryptedDocumentException, InvalidFormatException, IOException
   {
	   String path="../Study/Excel/mm.xls";
	   String sheetName="Demo";
	 CreateExcel.create(path);
	// CreateExcel.createSheet1(path, sheetName);
	 CreateExcel.createrows(sheetName, path);
	   
   }
}
