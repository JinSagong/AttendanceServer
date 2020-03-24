package com.jin.attendanceserver.util;

import java.io.File;
import java.io.FileReader;
import java.io.FileWriter;
import java.io.IOException;

public class RunningSharedPreference {

	final private File file = new File("C:\\attendance_running");

	public void setRunning(boolean running) {
		try {
			FileWriter fw = new FileWriter(file);
			if (running) {
				fw.write("running");
			} else {
				fw.write("stop");
			}
			fw.close();
		} catch (IOException e) {
			// TODO Auto-generated catch block
		}
	}

	public boolean getRunning() {
		String status = "";
		boolean running = false;
		try {
			FileReader fr = new FileReader(file);
			while (true) {
				int c = fr.read();
				if (c == -1) {
					break;
				}
				status += String.valueOf((char) c);
			}
			fr.close();
			if (status.equals("running")) {
				running = true;
			}
		} catch (IOException e) {
			// TODO Auto-generated catch block
		}

		return running;
	}
}
