package launcher;

import java.io.IOException;
import java.util.Scanner;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;

import logger.SomeClass;
import template.GenXlsToDoc;
import template.GenXlsxToDocx;


/*** Class designed to run project
** @author gruppa
*/

public class Launcher {
	
	
	/**
	 * @param args
	 * @throws EncryptedDocumentException
	 * @throws InvalidFormatException
	 * @throws IOException
	 */
	@SuppressWarnings("resource")
	public static void main(String[] args) throws EncryptedDocumentException, InvalidFormatException, IOException {
		
		SomeClass log = new SomeClass();
		log.laucnherLog();
		String format = new Scanner(System.in).nextLine();
			if (format.equals("xls")) {
				new GenXlsToDoc();
			} else if (format.equals("xlsx")) {
				new GenXlsxToDocx();
				
			} else if (format.equals("csv")) {
				new GenCSVToDoc();
				
			}else {
				log.launcherError();
				throw new IllegalArgumentException("Error, please input 'xls' or 'xlsx' or 'csv' ");
			}
	}

}
