package com.util;

import com.artofsolving.jodconverter.openoffice.connection.OpenOfficeConnection;
import com.artofsolving.jodconverter.openoffice.converter.OpenOfficeDocumentConverter;
import com.sun.star.awt.Size;
import com.sun.star.beans.PropertyValue;
import com.sun.star.lang.XComponent;
import com.sun.star.uno.UnoRuntime;
import com.sun.star.view.PaperFormat;
import com.sun.star.view.XPrintable;

/**
 * 自定义转换pdf方法，修改默认A4纸
 * @author wanggang
 * @version V1.0
 * @date 2016年11月18日 下午2:24:34
 */
public class MyOpenOfficeDocumentConverter extends OpenOfficeDocumentConverter {

	public MyOpenOfficeDocumentConverter(OpenOfficeConnection connection) {
		super(connection);
	}
	
	public final static Size A5;
	public final static Size A4;
	public final static Size A3;
	public final static Size B4, B5, B6;
	public final static Size KaoqinReport;
	
	static {
		A5 = new Size(14800, 21000);
		A4 = new Size(21, 29);
		A3 = new Size(29700, 42000);
		B4 = new Size(25000, 35300);
		B5 = new Size(17600, 25000);
		B6 = new Size(12500, 17600);
		KaoqinReport = new Size(254000, 2794000);
	}
	
	/*
	 * XComponent:xCalcComponent
	 * @seecom.artofsolving.jodconverter.openoffice.converter.
	 * AbstractOpenOfficeDocumentConverter
	 * #refreshDocument(com.sun.star.lang.XComponent)
	 */
	protected void refreshDocument(XComponent document) {
		
		super.refreshDocument(document);
		// The default paper format and orientation is A4 and portrait. To
		// change paper orientation
		// re set page size
		XPrintable xPrintable = (XPrintable) UnoRuntime.queryInterface(XPrintable.class, document);
		
		PropertyValue[] printerDesc = new PropertyValue[2];
		// Paper Orientation
//		printerDesc[0] = new PropertyValue();
//		printerDesc[0].Name ="PaperOrientation";
//		printerDesc[0].Value = PaperOrientation.PORTRAIT;
		// Paper Format
		printerDesc[0] = new PropertyValue();
		printerDesc[0].Name ="PaperFormat";
		printerDesc[0].Value = PaperFormat.USER;
		// Paper Size
		printerDesc[1] = new PropertyValue();
		printerDesc[1].Name ="PaperSize";
		printerDesc[1].Value = KaoqinReport;
		
		try {
			xPrintable.setPrinter(printerDesc);
		} catch (Exception e) {
			e.printStackTrace();
		}
	}

}
