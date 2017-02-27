package logger;

import java.util.logging.Logger;

import launcher.Launcher;
import template.GenXlsToDoc;
import template.GenXlsxToDocx;


/**
 *Class displays messages
 * @author Bogdan
 */
public class SomeClass {
	private static Logger log1 = Logger.getLogger(Launcher.class.getName());
	private static Logger log2 = Logger.getLogger(GenXlsToDoc.class.getName());
	private static Logger log3 = Logger.getLogger(GenXlsxToDocx.class.getName());
		

	public void laucnherLog() {
		System.out.println("Please, input format 'xls' or 'xlsx'");
		log1.info("Good!");
	}
	public void launcherError(){
		log1.severe("Error");
	}

	public void pathToXls() {
		System.out.println("Input path to excel file");
		log2.info("Good!");
	}
	
	public void pathToDoc(){
		System.out.println("Input path to world file");
		log2.info("Good!");
	}
	
	public void CheckContains(){
		log2.severe("Good");
	}
	
	public void createDoc(){
		System.out.println("Generation is successful. file created on disk C");
		log2.info("Good!");
	}
	
	public void pathToXlsx() {
		System.out.println("Input path to excel file");
		log3.info("Good");
	}
	
	public void pathToDocx(){
		System.out.println("Input path to world file");
		log3.info("Good");
	}
	
	public void createDocx(){
		System.out.println("Generation is successful. file created on disk C");
		log3.info("Good");
	}
	
	
	

}
