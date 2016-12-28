package com.encode.util;

import info.monitorenter.cpdetector.CharsetPrinter;

import java.io.BufferedReader;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStreamReader;
import java.io.OutputStreamWriter;
import java.io.UnsupportedEncodingException;
import java.io.Writer;

/**
 * 转码工具类
 * @author wanggang
 * @version V1.0
 * @date 2016年11月19日 下午2:40:51
 */
public class ConvertEncoding {
	
	public static void convert(String fileName) throws Exception {
		String filecode = guessEncoding(fileName);
		if(!"utf-8".equals(filecode.toLowerCase())) {
			
			convert(fileName, filecode, fileName, "UTF-8");
		}
	}
	
	/**
     * <dt><span class="strong">方法描述:</span></dt><dd>转换文件编码格式</dd>
     * <dt><span class="strong">作者:</span></dt><dd>wanggang</dd>
     * <dt><span class="strong">时间:</span></dt><dd>2016年11月19日 下午2:25:21</dd>
     * @param oldFile 要转换的文件
     * @param oldCharset 要转换的文件编码
     * @param newFlie 转换后的文件
     * @param newCharset 转换后的文件编码
     * @since v1.0
     */
    private static void convert(String oldFile, String oldCharset,
            String newFlie, String newCharset) {
        BufferedReader bin;
        FileOutputStream fos;
        StringBuffer content = new StringBuffer();
        try {
            bin = new BufferedReader(new InputStreamReader(new FileInputStream(
                    oldFile), oldCharset));
            String line = null;
            while ((line = bin.readLine()) != null) {
                // System.out.println("content:" + content);
                content.append(line);
                content.append(System.getProperty("line.separator"));
            }
            bin.close();
//            File dir = new File(newFlie.substring(0, newFlie.lastIndexOf("\\")));
//            if (!dir.exists()) {
//                dir.mkdirs();
//            }
            fos = new FileOutputStream(newFlie);
            Writer out = new OutputStreamWriter(fos, newCharset);
            out.write(content.toString());
            out.close();
            fos.close();
        } catch (UnsupportedEncodingException e) {
            e.printStackTrace();
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        }
    }
	
	/**
     * <dt><span class="strong">方法描述:</span></dt><dd>获得TXT文件的编码格式</dd>
     * <dt><span class="strong">作者:</span></dt><dd>wanggang</dd>
     * <dt><span class="strong">时间:</span></dt><dd>2016年11月19日 下午2:27:48</dd>
     * @param filename 文件名
     * @return
     * @since v1.0
     */
    private static String guessEncoding(String filename) {
        try {
            CharsetPrinter charsetPrinter = new CharsetPrinter();
            String encode = charsetPrinter.guessEncoding(new File(filename));
            if("windows-1252".equals(encode)) {
            	encode = "utf-16";
            }
            if("US-ASCII".equals(encode)) {
            	encode = "gb2312";
            }
            return encode;
        } catch (Exception e) {
            throw new RuntimeException(e);
        }
    }
	
}
