package com.aspose.gridweb.test.util;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

public class FileUtil {
	static File file = new File("E:\\fileoutput\\body.html");

	public static void putFile(String text) {
		FileOutputStream output = null;
		try {
			output = new FileOutputStream(file);
		} catch (FileNotFoundException e1) {
			e1.printStackTrace();
		}
		byte[] buff = text.getBytes();
		try {
			output.write(buff, 0, buff.length);
			output.close();
		} catch (IOException e) {
			e.printStackTrace();
		}
	}
}
