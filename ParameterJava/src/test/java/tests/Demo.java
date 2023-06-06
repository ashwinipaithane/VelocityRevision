package tests;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.ss.usermodel.WorkbookFactory;

import utils.Utility;

public class Demo {
	public static void main(String[] args) throws EncryptedDocumentException, IOException {
		System.out.println("Start");
		
		String v = Utility.getDataFromExcel("TestData", 0, 0);
		System.out.println(v);
		
		
		v = Utility.getDataFromExcel("TestData", 1, 0);
		System.out.println(v);
		
		v = Utility.getDataFromExcel("TestData", 2, 1);
		System.out.println(v);

		v = Utility.getDataFromExcel("TestData", 3, 1);
		System.out.println(v);

		v = Utility.getDataFromExcel("TestData", 3, 0);
		System.out.println(v);

		
	}
}
