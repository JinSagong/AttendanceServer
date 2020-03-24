package com.jin.attendanceserver.util;

public class RunTimeMeasurement {

	public String getRunTime(long timegab) {
		long hh, mm, ss;
		String HH, MM, SS;

		hh = timegab / 3600000;
		HH = String.valueOf(hh);
		mm = (timegab % 3600000) / 60000;
		MM = String.valueOf(mm);
		ss = (timegab % 60000) / 1000;
		SS = String.valueOf(ss);

		if (hh < 10) {
			HH = "0" + HH;
		}
		if (mm < 10) {
			MM = "0" + MM;
		}
		if (ss < 10) {
			SS = "0" + SS;
		}

		return HH + ":" + MM + ":" + SS;
	}
}