package template;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;
import java.util.Scanner;

import logger.SomeClass;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hwpf.HWPFDocument;
import org.apache.poi.hwpf.extractor.WordExtractor;
import org.apache.poi.hwpf.usermodel.Range;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Row;

import service.impl.ServiceManager;
import data.excel.ContactPerson;

public class GenCSVToDoc {
	private ServiceManager serviceManager;

	/**
	 * @return
	 */
	public CSVToDocService getCSVToDocService() {
		return serviceManager.getCSVToDocService();
	}

	/**
	 * @throws EncryptedDocumentException
	 * @throws InvalidFormatException
	 * @throws IOException
	 */
	public GenCSVToDoc() throws EncryptedDocumentException, InvalidFormatException, IOException {
		generation();
	}
	
	/**
	 * @throws IOException
	 * @throws EncryptedDocumentException
	 * @throws InvalidFormatException
	 */
	@SuppressWarnings("resource")
	public  void generation() throws IOException, EncryptedDocumentException, InvalidFormatException {
		
		serviceManager = ServiceManager.getInstance();
		
		List<ContactPerson> contacts = new ArrayList<ContactPerson>();
		SomeClass log = new SomeClass();
		log.pathToXls();
		
		String pathToExcel = new Scanner(System.in).nextLine();
		
		InputStream csv = new FileInputStream(pathToExcel); 
		
		HSSFWorkbook workBook = new HSSFWorkbook(csv);
		
		HSSFSheet sheet = workBook.getSheetAt(0);
		
		Iterator<Row> rows = sheet.rowIterator(); 
		
		if (rows.hasNext()) {
			rows.next();
		}

		log.pathToDoc();
		String pathToWorld = new Scanner(System.in).nextLine();
		
		while (rows.hasNext()) {
			HSSFRow row = (HSSFRow) rows.next();

			getCSVToDocService().getCellHSSF(row); 

			ContactPerson person = new ContactPerson();
			
			contacts.add(getCSVToDocService().getStringCellValue(person)); 
		
			FileInputStream inputStream = new FileInputStream(new File(pathToWorld));
			
			HWPFDocument doc = new HWPFDocument(inputStream);
			
			checkContains(doc,log);
			
			Range range = doc.getRange();
			
			getCSVToDocService().getRange(range); 
			

			getCSVToDocService().createDoc(row, doc); 
			log.createDoc();
			
			inputStream.close();

		}
	}
	/**
	 * @param doc
	 * @param log
	 */
	public  void checkContains(HWPFDocument doc, SomeClass log){
		String text="";
		@SuppressWarnings("resource")
		WordExtractor extractor = new WordExtractor(doc);
        String[] fileData = extractor.getParagraphText();
        for (int i = 0; i < fileData.length; i++) {
            if (fileData[i] != null) {
            	text += fileData[i];
            }
        }
        if (!(text.contains("{ToPosition}") && text.contains("{Organization}")
				&& text.contains("{ToName}") && text.contains("{FromName}")
				&& text.contains("{Address}") && text.contains("{StartDate}")
				&& text.contains("{Department}") && text.contains("{Position}")
				&& text.contains("{SignedName}")
				&& text.contains("{SignedDate}"))){	
        	
        	log.CheckContains();
			throw new IllegalArgumentException("Error, document does not required content");
		}
	}
}
}
