package com.util;

import java.io.BufferedInputStream;
import java.io.BufferedOutputStream;
import java.io.BufferedReader;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.InputStreamReader;
import java.util.Date;

import org.artofsolving.jodconverter.OfficeDocumentConverter;
import org.artofsolving.jodconverter.office.DefaultOfficeManagerConfiguration;
import org.artofsolving.jodconverter.office.OfficeManager;

/**
 * 使用LibreOffice转换txt、doc、docx、xls、xlsx、ppt、pptx、pdf
 * @author wanggang
 */
public class DocConverter {
	
	private String environment;// 环境1：windows,2:linux(涉及pdf2swf路径问题)
	private String fileString;
	private String outputPath = "";// 输入路径，如果不设置就输出在默认位置
	private String fileName;
	private File pdfFile;
	private File swfFile;
	private File docFile;
	private File odtFile;
	private String txtName;
	private String extName;
	private String libreOfficePath = "D:/ProgramFiles/LibreOffice5233";
	//private String libreOfficePath = "/opt/libreoffice5.2";
	private String swftoolsPath = "D:/ProgramFiles/SWFTools";
	//private String swftoolsPath = "/usr/swftools/bin/";

	public DocConverter(String fileString) throws Exception {
		ini(fileString);
	}

	public DocConverter() {

	}

	/*
	 * 重新设置 file @param fileString
	 */
	public void setFile(String fileString) throws Exception {
		ini(fileString);

	}

	/*
	 * 初始化 @param fileString
	 */
	/*
	 * private void ini(String fileString) { this.fileString = fileString;
	 * fileName = fileString.substring(0, fileString.lastIndexOf(".")); docFile
	 * = new File(fileString); pdfFile = new File(fileName + ".pdf"); swfFile =
	 * new File(fileName + ".swf"); //odtFile = new File(fileName + ".odt"); }
	 */
	private void ini(String fileString) throws Exception {

		try {
			/*
			 * this.fileString = fileString; fileName = fileString.substring(0,
			 * fileString.lastIndexOf("/")); docFile = new File(fileString);
			 * String s = fileString.substring(fileString.lastIndexOf("/") + 1,
			 * fileString.lastIndexOf(".")); fileName = fileName + "/" + s;
			 */
			initEnvironment();
			this.fileString = fileString;
			fileName = fileString.substring(0, fileString.lastIndexOf("."));
			docFile = new File(fileString);
			pdfFile = new File(fileName + ".pdf");
			swfFile = new File(fileName + ".swf");
			odtFile = new File(fileName + ".odt");
			// 用于处理TXT文档转化为PDF格式乱码,获取上传文件的名称（不需要后面的格式）
			String txtName = fileString.substring(fileString.lastIndexOf("."));
			System.out.println("txtName是：" + txtName);
			// 判断上传的文件是否是TXT文件
			if (txtName.equals(".txt") || txtName.equals(".TXT")) {
				// 定义相应的ODT格式文件名称
				odtFile = new File(fileName + ".odt");
				// 将上传的文档重新copy一份，并且修改为ODT格式，然后有ODT格式转化为PDF格式
				this.copyFile(docFile, odtFile);
				// 用于处理PDF文档
				pdfFile = new File(fileName + ".pdf");
			} else if (txtName.equals(".pdf") || txtName.equals(".PDF")) {
				// pdfFile = new File(fileName + ".pdf");
				// this.copyFile(docFile, pdfFile);
				pdf2swf();
				swfFile = new File(fileName + ".swf");
				return;
			} else {
				pdfFile = new File(fileName + ".pdf");
			}
			swfFile = new File(fileName + ".swf");
		} catch (Exception e) {
			e.printStackTrace();
		}
	}

	private void odt2pdf() throws Exception {
//		// 这里是OpenOffice的安装目录, 在我的项目中,为了便于拓展接口,没有直接写成这个样子,但是这样是绝对没问题的
//		String OpenOffice_HOME = "D:\\OpenOffice 4";
//		// 如果从文件中读取的URL地址最后一个字符不是 '\'，则添加'\'
//		if (OpenOffice_HOME.charAt(OpenOffice_HOME.length() - 1) != '\\') {
//			OpenOffice_HOME += "\\";
//		}
//		String command = OpenOffice_HOME
//				+ "program\\soffice.exe -headless -accept=\"socket,host=127.0.0.1,port=8100;urp;\"";
//		Process pro = Runtime.getRuntime().exec(command);
		if (odtFile.exists()) {
			if (!pdfFile.exists()) {
				//OpenOfficeConnection connection = new SocketOpenOfficeConnection(8100);
				DefaultOfficeManagerConfiguration configuration = new DefaultOfficeManagerConfiguration();
				OfficeManager officeManager = null;
				try {
					configuration.setOfficeHome(new File(libreOfficePath));
					configuration.setPortNumber(8100);
					// 设置任务执行超时为50分钟，这里就是导致转换超时的配置
					//configuration.setTaskExecutionTimeout(1000 * 60 * 50L);
					// 设置任务队列超时为24小时
					//configuration.setTaskQueueTimeout(1000 * 60 * 60 * 24L);
					officeManager = configuration.buildOfficeManager();
					officeManager.start();
					
					OfficeDocumentConverter converter = new OfficeDocumentConverter(officeManager);
					// DocumentConverter converter = new
					// MyOpenOfficeDocumentConverter(connection);
					converter.convert(odtFile, pdfFile);
					// close the connection
					System.out.println("****pdf转换成功，PDF输出：" + pdfFile.getPath() + "****");
				} catch (com.artofsolving.jodconverter.openoffice.connection.OpenOfficeException e) {
					e.printStackTrace();
					System.out.println("****swf转换器异常，读取转换文件失败****");
					throw e;
				} catch (Exception e) {
					e.printStackTrace();
					throw e;
				} finally {
					if(null!=officeManager) officeManager.stop();
				}
			} else {
				System.out.println("****已经转换为pdf，不需要再进行转化****");
			}
		} else {
			System.out.println("****swf转换器异常，需要转换的文档不存在，无法转换****");
		}
	}

	private void copyFile(File sourceFile, File targetFile) throws Exception {
		// 新建文件输入流并对它进行缓冲 
		FileInputStream input = new FileInputStream(sourceFile);
		BufferedInputStream inBuff = new BufferedInputStream(input);
		//  新建文件输出流并对它进行缓冲
		FileOutputStream output = new FileOutputStream(targetFile);
		BufferedOutputStream outBuff = new BufferedOutputStream(output);
		//  缓冲数组 
		byte[] b = new byte[1024 * 5];
		int len;
		while ((len = inBuff.read(b)) != -1) {
			outBuff.write(b, 0, len);
		}
		//  刷新此缓冲的输出流
		outBuff.flush();
		//  关闭流
		inBuff.close();
		outBuff.close();
		output.close();
		input.close();
	}

	/*
	 * 转为PDF @param file
	 */
	private void doc2pdf() throws Exception {
//		// 这里是OpenOffice的安装目录, 在我的项目中,为了便于拓展接口,没有直接写成这个样子,但是这样是绝对没问题的
//		String OpenOffice_HOME = "D:\\OpenOffice 4";
//		// 如果从文件中读取的URL地址最后一个字符不是 '\'，则添加'\'
//		if (OpenOffice_HOME.charAt(OpenOffice_HOME.length() - 1) != '\\') {
//			OpenOffice_HOME += "\\";
//		}
//		String command = OpenOffice_HOME
//				+ "program\\soffice.exe -headless -accept=\"socket,host=127.0.0.1,port=8100;urp;\"";
//		Process pro = Runtime.getRuntime().exec(command);
		System.out.println("libreOffice开始转换..............................");
		Long startTime = new Date().getTime();
		
		if (docFile.exists()) {
			if (!pdfFile.exists()) {
				DefaultOfficeManagerConfiguration configuration = new DefaultOfficeManagerConfiguration();
				OfficeManager officeManager = null;
				try {
					configuration.setOfficeHome(new File(libreOfficePath));
					configuration.setPortNumber(8100);
					// 设置任务执行超时为50分钟，这里就是导致转换超时的配置
					//configuration.setTaskExecutionTimeout(1000 * 60 * 50L);
					// 设置任务队列超时为24小时
					//configuration.setTaskQueueTimeout(1000 * 60 * 60 * 24L);
					officeManager = configuration.buildOfficeManager();
					officeManager.start();
					System.out.println("...start 开始喽 .....");
					
					OfficeDocumentConverter converter = new OfficeDocumentConverter(officeManager);
					// DocumentConverter converter = new
					// MyOpenOfficeDocumentConverter(connection);
					//converter.convert(inputFile, outputFile);
					converter.convert(docFile, pdfFile);
					// close the connection
					System.out.println("转换结束，累死啦。。。。。");
					long endTime = new Date().getTime();
					long time = endTime - startTime;
					System.out.println("libreOffice转换所用时间为："+time);
					System.out.println("****pdf转换成功，PDF输出：" + pdfFile + "****");
				} catch (com.artofsolving.jodconverter.openoffice.connection.OpenOfficeException e) {
					e.printStackTrace();
					System.out.println("****swf转换器异常，读取转换文件失败****");
					throw e;
				} catch (Exception e) {
					e.printStackTrace();
					throw e;
				} finally {
					if(null!=officeManager) officeManager.stop();
				}
			} else {
				System.out.println("****已经转换为pdf，不需要再进行转化****");
			}
		} else {
			System.out.println("****swf转换器异常，需要转换的文档不存在，无法转换****");
		}
	}

	/*
	 * 转换成swf
	 */

	private void pdf2swf() throws Exception {
//		// 这里是OpenOffice的安装目录, 在我的项目中,为了便于拓展接口,没有直接写成这个样子,但是这样是绝对没问题的
//		String OpenOffice_HOME = "D:\\ProgramFiles\\OpenOffice 4";
//		// 如果从文件中读取的URL地址最后一个字符不是 '\'，则添加'\'
//		if (OpenOffice_HOME.charAt(OpenOffice_HOME.length() - 1) != '\\') {
//			OpenOffice_HOME += "\\";
//		}
//		String command = OpenOffice_HOME
//				+ "program\\soffice.exe -headless -accept=\"socket,host=127.0.0.1,port=8100;urp;\"";
//		Process pro = Runtime.getRuntime().exec(command);
		Runtime r = Runtime.getRuntime();
		if (!swfFile.exists()) {
			if (pdfFile.exists()) {
				if ("1".equals(environment))// windows环境处理
				{
					try {
						// 这里根据SWFTools安装路径需要进行相应更改

						Process p = r.exec(swftoolsPath + "/pdf2swf.exe "
								+ pdfFile.getPath() + " -o "
								+ swfFile.getPath() + " -T 9");

						System.out.print(loadStream(p.getInputStream()));
						System.err.print(loadStream(p.getErrorStream()));
						System.out.print(loadStream(p.getInputStream()));
						System.err.println("****swf转换成功，文件输出："
								+ swfFile.getPath() + "****");
						if (pdfFile.exists()) {
							pdfFile.delete();
						}
					} catch (Exception e) {
						e.printStackTrace();
						throw e;
					}
				} else if ("2".equals(environment))// linux环境处理
				{
					try {
						Process p = r.exec(swftoolsPath+"/pdf2swf " + pdfFile.getPath() + " -o " + swfFile.getPath() + " -T 9 -Q 500");
						System.out.print(loadStream(p.getInputStream()));
						System.err.print(loadStream(p.getErrorStream()));
						System.err.println("****swf转换成功，文件输出："
								+ swfFile.getPath() + "****");
						if (pdfFile.exists()) {
							pdfFile.delete();
						}
					} catch (Exception e) {
						e.printStackTrace();
						throw e;
					}
				}
			} else {
				System.out.println("****pdf不存在，无法转换****");
			}
		} else {
			System.out.println("****swf已存在不需要转换****");
		}
	}

	static String loadStream(InputStream in) throws IOException {
		int ptr = 0;
		// 把InputStream字节流 替换为BufferedReader字符流 2013-07-17修改
		BufferedReader reader = new BufferedReader(new InputStreamReader(in));
		StringBuilder buffer = new StringBuilder();
		while ((ptr = reader.read()) != -1) {
			buffer.append((char) ptr);
		}
		return buffer.toString();
	}

	/*
	 * 转换主方法
	 */
	public boolean conver() {

		if (swfFile.exists()) {
			System.out.println("****swf转换器开始工作，该文件已经转换为swf****");
			return true;
		}

		if ("1".equals(environment)) {
			System.out.println("****swf转换器开始工作，当前设置运行环境windows****");
		} else {
			System.out.println("****swf转换器开始工作，当前设置运行环境linux****");
		}

		try {
			//如果是txt
			if(!(odtFile==null)) {
				odt2pdf();
			} else {
				doc2pdf();
			}
			pdf2swf();
		} catch (Exception e) {
			e.printStackTrace();
			return false;
		}

		if (swfFile.exists()) {
			return true;
		} else {
			return false;
		}
	}

	/*
	 * 返回文件路径 @param s
	 */
	public String getswfPath() {
		if (swfFile.exists()) {
			String tempString = swfFile.getPath();
			tempString = tempString.replaceAll("\\\\", "/");
			return tempString;
		} else {
			return "";
		}
	}

	/*
	 * 设置输出路径
	 */
	public void setOutputPath(String outputPath) {
		this.outputPath = outputPath;
		if (!outputPath.equals("")) {
			String realName = fileName.substring(fileName.lastIndexOf("/"),
					fileName.lastIndexOf("."));
			if (outputPath.charAt(outputPath.length()) == '/') {
				swfFile = new File(outputPath + realName + ".swf");
			} else {
				swfFile = new File(outputPath + realName + ".swf");
			}
		}
	}
	
	/**
	 * 判断运行环境 windows or linux
	 * 环境1：windows,2:linux(涉及pdf2swf路径问题)
	 */
	public void initEnvironment() {
		
		String osName = System.getProperties().getProperty("os.name");
		
		if(osName.toLowerCase().contains("windows")) {
			environment = "1";
		} else {
			environment = "2";
		}
	}
	
	public static void main(String s[]) {
		DocConverter d;

		String fileName = "D:/upload/3.pdf";

		try {
			// if(fileName.contains(".pdf")) {
			// d = new DocConverter();
			// } else {
			d = new DocConverter(fileName);
			// }
			d.conver();
		} catch (Exception e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
	}
}