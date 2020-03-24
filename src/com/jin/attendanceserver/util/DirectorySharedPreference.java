package com.jin.attendanceserver.util;

import java.io.File;
import java.io.FileReader;
import java.io.FileWriter;
import java.io.IOException;

public class DirectorySharedPreference {

	final private File file = new File("C:\\attendance_path");

	public void setDirectory(String directory) {
		try {
			FileWriter fw = new FileWriter(file);
			fw.write(directory);
			fw.close();
		} catch (IOException e) {
			// TODO Auto-generated catch block
		}
	}

	public String getDirectory() {
		String directory = "";
		try {
			FileReader fr = new FileReader(file);
			while (true) {
				int c = fr.read();
				if (c == -1) {
					break;
				}
				directory += String.valueOf((char) c);
			}
			fr.close();
		} catch (IOException e) {
			// TODO Auto-generated catch block
			directory = "C:\\Application";
		}

		return directory;
	}
}
