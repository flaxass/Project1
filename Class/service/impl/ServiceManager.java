package service.impl;

import service.XlsToDocService;
import service.XlsxToDocxService;


/**
 * Class for manager service
 * @author gruppa
 *
 */
public class ServiceManager {

	private XlsToDocService xlsToDocService;
	private XlsxToDocxService xlsxToDocxService;
	private CSVToDocService csvToDocService;
	
	/**
	 * @return
	 */
	public XlsToDocService getXlsToDocService(){
		return xlsToDocService;
	}
	/**
	 * @return
	 */
	public XlsxToDocxService getXlsxToDocxService(){
		return xlsxToDocxService;
	}
	public CSVToDocService getCSVToDocService(){
		return csvToDocService;
	}
	/**
	 * Created service
	 */
	private ServiceManager(){
		super();
		xlsToDocService = (XlsToDocService) ServiceFactory.createService(new XlsToDocServiceImpl());
		xlsxToDocxService = (XlsxToDocxService) ServiceFactory.createService(new XlsxToDocxServiceImpl());
		csvToDocService = (CSVToDocService) ServiceFactory.createService(new CSVToDocServiceImpl());
	}
	/**
	 * @return
	 */
	public static ServiceManager getInstance(){
		return new ServiceManager();
	}
}
