package com.moudles;

import java.io.File;

import org.artofsolving.jodconverter.OfficeDocumentConverter;
import org.artofsolving.jodconverter.office.DefaultOfficeManagerConfiguration;
import org.artofsolving.jodconverter.office.OfficeManager;

public class DocToPdf {
	
	private static String libreOfficePath = "D:/ProgramFiles/LibreOffice5233";
	private static String swftoolsPath = "D:/ProgramFiles/SWFTools";
	private static String filePath = "D:/work/wordtopdf/";

	public static void main(String[] args) {
		File docFile = new File(filePath + "新建文本文档.txt");
		File pdfFile = new File(filePath + "新建文本文档"+".pdf");
		try {
			convertDocToPdf(docFile, pdfFile);
		} catch (Exception e) {
			e.printStackTrace();
		}
	}
	
	public static void convertDocToPdf(File docFile, File pdfFile) throws Exception {
		DefaultOfficeManagerConfiguration configuration = new DefaultOfficeManagerConfiguration();
		OfficeManager officeManager = null;
		
		configuration.setOfficeHome(new File(libreOfficePath));
		configuration.setPortNumber(8100);
		// 设置任务执行超时为50分钟，这里就是导致转换超时的配置
		//configuration.setTaskExecutionTimeout(1000 * 60 * 50L);
		// 设置任务队列超时为24小时
		//configuration.setTaskQueueTimeout(1000 * 60 * 60 * 24L);
		officeManager = configuration.buildOfficeManager();
        officeManager.start();
		
		OfficeDocumentConverter converter = new OfficeDocumentConverter(officeManager);
		//DocumentConverter converter = new MyOpenOfficeDocumentConverter(connection);
		converter.convert(docFile, pdfFile);
		// close the connection
		System.out.println("****pdf转换成功，PDF输出：" + pdfFile.getPath() + "****");
		officeManager.stop();
	}

}
