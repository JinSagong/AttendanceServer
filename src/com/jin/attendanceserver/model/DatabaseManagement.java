package com.jin.attendanceserver.model;

import java.awt.TextArea;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.FileReader;
import java.io.IOException;
import java.text.SimpleDateFormat;
import java.util.Calendar;
import java.util.ConcurrentModificationException;
import java.util.Date;

import org.apache.commons.math3.util.Pair;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.jin.attendanceserver.formatting.Formatting;
import com.jin.attendanceserver.formatting.FpFormatting;
import com.jin.attendanceserver.formatting.FpMonthlyFormatting;
import com.jin.attendanceserver.formatting.FpYearlyFormatting;
import com.jin.attendanceserver.formatting.HomeFormatting;

public class DatabaseManagement {

	// Set file name
	final private String fn_attendance = "2020_MAIN";
	final private String fn_attendance_general = "2020_GENERAL";
	final private String fn_attendance_group = "2020_GROUP";
	final private String fn_attendance_bc = "2020_BC";
	final private String fn_attendance_home = "2020_MAIN_home";
	final private String fn_attendance_general_home = "2020_GENERAL_home";
	final private String fn_attendance_group_home = "2020_GROUP_home";
	final private String fn_attendance_bc_home = "2020_BC_home";
	final private String fn_logs = "logs";
	final private String fn_id = "id";
	final private String fn_status = "status";
	final private String fn_fpattendance = "fp_attendance";
	final private String fn_fpbeliever = "fp_believer";
	final private String fn_fpwordmovement = "fp_wordmovement";

	TextArea Logs;
	SimpleDateFormat formatter;
	String path;
	String fp_attendance, fp_attendance_general, fp_attendance_group, fp_attendance_bc, fp_attendance_home,
			fp_attendance_general_home, fp_attendance_group_home, fp_attendance_bc_home, fp_logs, fp_id, fp_status,
			fp_fpattendance, fp_fpbeliever, fp_fpwordmovement;
	File f_attendance, f_attendance_general, f_attendance_group, f_attendance_bc, f_attendance_home,
			f_attendance_general_home, f_attendance_group_home, f_attendance_bc_home, f_logs, f_id, f_status,
			f_fpattendance, f_fpbeliever, f_fpwordmovement;
	XSSFWorkbook wb_attendance, wb_attendance_general, wb_attendance_group, wb_attendance_bc, wb_attendance_home,
			wb_attendance_general_home, wb_attendance_group_home, wb_attendance_bc_home, wb_logs, wb_id, wb_status,
			wb_fpattendance, wb_fpbeliever, wb_fpwordmovement;
	XSSFSheet sh_attendance, sh_attendance_general, sh_attendance_group, sh_attendance_bc, sh_attendance_bc_read,
			sh_attendance_bc_group, sh_attendance_home, sh_attendance_general_home, sh_attendance_group_home,
			sh_attendance_bc_home, sh_attendance_bc_read_home, sh_logs, sh_id, sh_status_general, sh_status_group,
			sh_status_bc, sh_status_fp, sh_status_general_home, sh_status_group_home, sh_status_bc_home,
			sh_fpattendance_attendance;
	int num_attendance, num_attendance_general, num_attendance_group, num_attendance_bc, num_attendance_bc_read,
			num_attendance_bc_group, num_logs, num_id, num_status_general, num_status_group, num_status_bc,
			num_status_fp, num_status_general_home, num_status_group_home, num_status_bc_home,
			num_fpattendance_attendance;

	int week_of_year;

	public DatabaseManagement(String directory_path, TextArea Logs, SimpleDateFormat formatter) {
		// Set file path
		fp_attendance = directory_path + "\\data\\" + fn_attendance + ".xlsx";
		fp_attendance_general = directory_path + "\\data\\" + fn_attendance_general + ".xlsx";
		fp_attendance_group = directory_path + "\\data\\" + fn_attendance_group + ".xlsx";
		fp_attendance_bc = directory_path + "\\data\\" + fn_attendance_bc + ".xlsx";
		fp_attendance_home = directory_path + "\\data_home\\" + fn_attendance_home + ".xlsx";
		fp_attendance_general_home = directory_path + "\\data_home\\" + fn_attendance_general_home + ".xlsx";
		fp_attendance_group_home = directory_path + "\\data_home\\" + fn_attendance_group_home + ".xlsx";
		fp_attendance_bc_home = directory_path + "\\data_home\\" + fn_attendance_bc_home + ".xlsx";
		fp_logs = directory_path + "\\src\\" + fn_logs + ".xlsx";
		fp_id = directory_path + "\\src\\" + fn_id + ".xlsx";
		fp_status = directory_path + "\\src\\" + fn_status + ".xlsx";
		fp_fpattendance = directory_path + "\\src\\" + fn_fpattendance + ".xlsx";
		fp_fpbeliever = directory_path + "\\src\\" + fn_fpbeliever + ".xlsx";
		fp_fpwordmovement = directory_path + "\\src\\" + fn_fpwordmovement + ".xlsx";

		// Set functions
		path = directory_path;
		this.Logs = Logs;
		this.formatter = formatter;
		week_of_year = Calendar.getInstance().get(Calendar.WEEK_OF_YEAR);
		if (week_of_year == 1) {
			if (Calendar.getInstance().get(Calendar.MONTH) == 11) {
				week_of_year = 53;
			}
		}
	}

	public void init() {
		try {
			f_attendance = new File(fp_attendance);
			f_attendance_general = new File(fp_attendance_general);
			f_attendance_group = new File(fp_attendance_group);
			f_attendance_bc = new File(fp_attendance_bc);
			f_attendance_home = new File(fp_attendance_home);
			f_attendance_general_home = new File(fp_attendance_general_home);
			f_attendance_group_home = new File(fp_attendance_group_home);
			f_attendance_bc_home = new File(fp_attendance_bc_home);
			f_logs = new File(fp_logs);
			f_id = new File(fp_id);
			f_status = new File(fp_status);
			f_fpattendance = new File(fp_fpattendance);
			f_fpbeliever = new File(fp_fpbeliever);
			f_fpwordmovement = new File(fp_fpwordmovement);

			wb_attendance = new XSSFWorkbook(new FileInputStream(f_attendance));
			wb_attendance_general = new XSSFWorkbook(new FileInputStream(f_attendance_general));
			wb_attendance_group = new XSSFWorkbook(new FileInputStream(f_attendance_group));
			wb_attendance_bc = new XSSFWorkbook(new FileInputStream(f_attendance_bc));
			wb_attendance_home = new XSSFWorkbook(new FileInputStream(f_attendance_home));
			wb_attendance_general_home = new XSSFWorkbook(new FileInputStream(f_attendance_general_home));
			wb_attendance_group_home = new XSSFWorkbook(new FileInputStream(f_attendance_group_home));
			wb_attendance_bc_home = new XSSFWorkbook(new FileInputStream(f_attendance_bc_home));
			wb_logs = new XSSFWorkbook(new FileInputStream(f_logs));
			wb_id = new XSSFWorkbook(new FileInputStream(f_id));
			wb_status = new XSSFWorkbook(new FileInputStream(f_status));
			wb_fpattendance = new XSSFWorkbook(new FileInputStream(f_fpattendance));
			wb_fpbeliever = new XSSFWorkbook(new FileInputStream(f_fpbeliever));
			wb_fpwordmovement = new XSSFWorkbook(new FileInputStream(f_fpwordmovement));

			sh_attendance = wb_attendance.getSheetAt(0);
			sh_attendance_general = wb_attendance_general.getSheetAt(0);
			sh_attendance_group = wb_attendance_group.getSheetAt(0);
			sh_attendance_bc = wb_attendance_bc.getSheetAt(0);
			sh_attendance_bc_read = wb_attendance_bc.getSheetAt(1);
			sh_attendance_bc_group = wb_attendance_bc.getSheetAt(2);
			sh_attendance_home = wb_attendance_home.getSheetAt(0);
			sh_attendance_general_home = wb_attendance_general_home.getSheetAt(0);
			sh_attendance_group_home = wb_attendance_group_home.getSheetAt(0);
			sh_attendance_bc_home = wb_attendance_bc_home.getSheetAt(0);
			sh_attendance_bc_read_home = wb_attendance_bc_home.getSheetAt(1);
			sh_logs = wb_logs.getSheetAt(0);
			sh_id = wb_id.getSheetAt(0);
			sh_status_general = wb_status.getSheetAt(0);
			sh_status_group = wb_status.getSheetAt(1);
			sh_status_bc = wb_status.getSheetAt(2);
			sh_status_fp = wb_status.getSheetAt(3);
			sh_status_general_home = wb_status.getSheetAt(4);
			sh_status_group_home = wb_status.getSheetAt(5);
			sh_status_bc_home = wb_status.getSheetAt(6);
			sh_fpattendance_attendance = wb_fpattendance.getSheetAt(0);

			num_attendance = sh_attendance.getPhysicalNumberOfRows();
			num_attendance_general = sh_attendance_general.getPhysicalNumberOfRows();
			num_attendance_group = sh_attendance_group.getPhysicalNumberOfRows();
			num_attendance_bc = sh_attendance_bc.getPhysicalNumberOfRows();
			num_attendance_bc_read = sh_attendance_bc_read.getPhysicalNumberOfRows();
			num_attendance_bc_group = sh_attendance_bc_group.getPhysicalNumberOfRows();
			num_logs = sh_logs.getPhysicalNumberOfRows();
			num_id = sh_id.getPhysicalNumberOfRows();
			num_status_general = sh_status_general.getPhysicalNumberOfRows();
			num_status_group = sh_status_group.getPhysicalNumberOfRows();
			num_status_bc = sh_status_bc.getPhysicalNumberOfRows();
			num_status_fp = sh_status_fp.getPhysicalNumberOfRows();
			num_status_general_home = sh_status_general_home.getPhysicalNumberOfRows();
			num_status_group_home = sh_status_group_home.getPhysicalNumberOfRows();
			num_status_bc_home = sh_status_bc_home.getPhysicalNumberOfRows();
			num_fpattendance_attendance = sh_fpattendance_attendance.getPhysicalNumberOfRows();
		} catch (IOException e) {
			// TODO Auto-generated catch block
		}
	}

	public void save() {
		try {
			wb_attendance.write(new FileOutputStream(f_attendance));
			wb_attendance_general.write(new FileOutputStream(f_attendance_general));
			wb_attendance_group.write(new FileOutputStream(f_attendance_group));
			wb_attendance_bc.write(new FileOutputStream(f_attendance_bc));
			wb_attendance_home.write(new FileOutputStream(f_attendance_home));
			wb_attendance_general_home.write(new FileOutputStream(f_attendance_general_home));
			wb_attendance_group_home.write(new FileOutputStream(f_attendance_group_home));
			wb_attendance_bc_home.write(new FileOutputStream(f_attendance_bc_home));
			wb_logs.write(new FileOutputStream(f_logs));
			wb_status.write(new FileOutputStream(f_status));
			wb_fpattendance.write(new FileOutputStream(f_fpattendance));
			wb_fpbeliever.write(new FileOutputStream(f_fpbeliever));
			wb_fpwordmovement.write(new FileOutputStream(f_fpwordmovement));

		} catch (IOException e) {
			// TODO Auto-generated catch block
		} catch (ConcurrentModificationException e) {
		}
	}

	public void exit() {
		try {
			wb_attendance.close();
			wb_attendance_general.close();
			wb_attendance_group.close();
			wb_attendance_bc.close();
			wb_attendance_home.close();
			wb_attendance_general_home.close();
			wb_attendance_group_home.close();
			wb_attendance_bc_home.close();
			wb_logs.close();
			wb_id.close();
			wb_status.close();
			wb_fpattendance.close();
			wb_fpbeliever.close();
			wb_fpwordmovement.close();

		} catch (IOException e) {
			// TODO Auto-generated catch block
		}
	}

	public Pair<Boolean, Boolean> getOnOff() {
		Boolean offline = false;
		Boolean online = false;
		File offlineFile = new File(path + "\\src\\offline");
		File onlineFile = new File(path + "\\src\\online");

		String offlineString = "";
		try {
			FileReader fr = new FileReader(offlineFile);
			while (true) {
				int c = fr.read();
				if (c == -1) {
					break;
				}
				offlineString += String.valueOf((char) c);
			}
			fr.close();
			if (offlineString.equals("TRUE"))
				offline = true;
		} catch (IOException e) {
			// TODO Auto-generated catch block
		}

		String onlineString = "";
		try {
			FileReader fr = new FileReader(onlineFile);
			while (true) {
				int c = fr.read();
				if (c == -1) {
					break;
				}
				onlineString += String.valueOf((char) c);
			}
			fr.close();
			if (onlineString.equals("TRUE"))
				online = true;
		} catch (IOException e) {
			// TODO Auto-generated catch block
		}

		return new Pair<Boolean, Boolean>(offline, online);
	}

	public int getNumLogs() {
		return num_logs - 1;
	}

	public String getObject(String type) {
		String obj = "NULL";

		String tempType = type;
		String[] criteria = type.split("#");
		if (criteria.length == 2) {
			tempType = criteria[0];
		}
		StringBuffer buf = new StringBuffer(type);
		String sub = buf.substring(0, 2);
		if (sub.equals("ge")) {
			String[][] list = readStatus("general");
			for (int i = 0; i < list[0].length; i++) {
				if (list[0][i].equals(tempType)) {
					obj = list[1][i];
					break;
				}
			}
		} else if (sub.equals("gr")) {
			String[][] list = readStatus("group");
			for (int i = 0; i < list[0].length; i++) {
				if (list[0][i].equals(tempType)) {
					obj = list[1][i];
					break;
				}
			}
		} else if (sub.equals("bc")) {
			String[][] list = readStatus("bc");
			for (int i = 0; i < list[0].length; i++) {
				if (list[0][i].equals(tempType)) {
					obj = list[1][i];
					break;
				}
			}
		} else if (sub.equals("fp")) {
			String[][] list = readStatus("fp");
			for (int i = 0; i < list[0].length; i++) {
				if (list[0][i].equals(tempType)) {
					obj = list[1][i];
					break;
				}
			}
		}

		if (criteria.length == 2) {
			switch (criteria[1]) {
			case "1":
				obj = "[주일1부예배] " + obj;
				break;
			case "2":
				obj = "[주일2부예배] " + obj;
				break;
			case "3":
				obj = "[주일오후예배] " + obj;
				break;
			case "4":
				obj = "[수요예배] " + obj;
				break;
			case "5":
				obj = "[금요기도회] " + obj;
				break;
			}
		}

		return obj;
	}

	public void writeLogs(String key, String type, String obj) {
		String time = formatter.format(new Date());
		String s_key = "NULL";
		String s_obj = "NULL";
		String LOG;

		switch (type) {
		case "TurnOn":
			LOG = "[" + time + "] " + type;
			break;
		case "TurnOff":
			LOG = "[" + time + "] " + type + " <" + obj + ">";
			break;
		case "Check":
			s_key = readId(null, key)[3];
			s_obj = getObject(obj);
			LOG = "[" + time + "] " + s_key + " " + type + " " + s_obj;
			break;
		case "FpCheck":
			s_key = readId(null, key)[3];
			s_obj = getObject(obj);
			LOG = "[" + time + "] " + s_key + " " + type + " " + s_obj;
			break;
		case "FpAdd":
			s_key = readId(null, key)[3];
			s_obj = getObject(obj);
			LOG = "[" + time + "] " + s_key + " " + type + " " + s_obj;
			break;
		case "Add":
			s_key = readId(null, key)[3];
			s_obj = getObject(obj);
			LOG = "[" + time + "] " + s_key + " " + type + " " + s_obj;
			break;
		default:
			s_key = readId(null, key)[3];
			LOG = "[" + time + "] " + s_key + " " + type;
			break;
		}

		sh_logs.createRow(num_logs).createCell(0).setCellValue(time);
		sh_logs.getRow(num_logs).createCell(1).setCellValue(s_key);
		sh_logs.getRow(num_logs).createCell(2).setCellValue(type);
		sh_logs.getRow(num_logs).createCell(3).setCellValue(s_obj);
		if (type.equals("TurnOff")) {
			sh_logs.getRow(num_logs).createCell(4).setCellValue(obj);
			Logs.append(LOG + "\n\n");
		} else {
			sh_logs.getRow(num_logs).createCell(4).setCellValue("NULL");
			Logs.append(LOG + "\n");
		}

		num_logs++;
	}

	public void writeAttendance(String type, String[][] attendance) {
		int week = week_of_year + 1;

		StringBuffer buf = new StringBuffer(type);
		String sub = buf.substring(0, 2);

		for (int i = 0; i < attendance[0].length; i++) {
			boolean bc_group = false;

			switch (sub) {
			case "ge":
				for (int j = 1; j < num_attendance_general - 1; j++) {
					if (sh_attendance_general.getRow(j).getCell(2).toString().equals(attendance[0][i])) {
						if (attendance[1][i].equals("TRUE")) {
							sh_attendance_general.getRow(j).createCell(week).setCellValue("○");
							sh_attendance_general.getRow(j).createCell(55).setCellValue("");
						} else {
							try {
								sh_attendance_general.getRow(j)
										.removeCell(sh_attendance_general.getRow(j).getCell(week));
							} catch (NullPointerException e) {
							}
							sh_attendance_general.getRow(j).createCell(55).setCellValue(attendance[2][i]);
						}
						break;
					}
				}
				break;

			case "gr":
				for (int j = 1; j < num_attendance_bc_group; j++) {
					if (sh_attendance_bc_group.getRow(j).getCell(3).toString().equals(attendance[0][i])) {
						bc_group = true;
						break;
					}
				}
				for (int j = 2; j < num_attendance_group - 2; j++) {
					try {
						if (sh_attendance_group.getRow(j).getCell(2).toString().equals(attendance[0][i])) {
						}
					} catch (NullPointerException e) {
						continue;
					}
					if (sh_attendance_group.getRow(j).getCell(2).toString().equals(attendance[0][i])) {
						if (attendance[1][i].equals("TRUE")) {
							sh_attendance_group.getRow(j).createCell(week).setCellValue("○");
							sh_attendance_group.getRow(j).createCell(55).setCellValue("");
						} else {
							try {
								sh_attendance_group.getRow(j).removeCell(sh_attendance_group.getRow(j).getCell(week));
							} catch (NullPointerException e) {
							}
							sh_attendance_group.getRow(j).createCell(55).setCellValue(attendance[2][i]);
						}
						break;
					}
				}
				if (bc_group) {
					for (int j = 2; j < num_attendance_bc_read - 3; j++) {
						try {
							if (sh_attendance_bc_read.getRow(j).getCell(2).toString().equals(attendance[0][i])) {
							}
						} catch (NullPointerException e) {
							continue;
						}
						if (sh_attendance_bc_read.getRow(j).getCell(2).toString().equals(attendance[0][i])) {
							if (attendance[1][i].equals("TRUE")) {
								sh_attendance_bc_read.getRow(j).createCell(week).setCellValue("○");
								sh_attendance_bc_read.getRow(j).createCell(55).setCellValue("");
							} else {
								try {
									sh_attendance_bc_read.getRow(j)
											.removeCell(sh_attendance_bc_read.getRow(j).getCell(week));
								} catch (NullPointerException e) {
								}
								sh_attendance_bc_read.getRow(j).createCell(55).setCellValue(attendance[2][i]);
							}
							break;
						}
					}
				}
				break;

			case "bc":
				for (int j = 1; j < num_attendance_bc_group; j++) {
					if (sh_attendance_bc_group.getRow(j).getCell(3).toString().equals(attendance[0][i])) {
						bc_group = true;
						break;
					}
				}
				for (int j = 2; j < num_attendance_bc_read - 3; j++) {
					try {
						if (sh_attendance_bc_read.getRow(j).getCell(2).toString().equals(attendance[0][i])) {
						}
					} catch (NullPointerException e) {
						continue;
					}
					if (sh_attendance_bc_read.getRow(j).getCell(2).toString().equals(attendance[0][i])) {
						if (attendance[1][i].equals("TRUE")) {
							sh_attendance_bc_read.getRow(j).createCell(week).setCellValue("○");
							sh_attendance_bc_read.getRow(j).createCell(55).setCellValue("");
						} else {
							try {
								sh_attendance_bc_read.getRow(j)
										.removeCell(sh_attendance_bc_read.getRow(j).getCell(week));
							} catch (NullPointerException e) {
							}
							sh_attendance_bc_read.getRow(j).createCell(55).setCellValue(attendance[2][i]);
						}
						break;
					}
				}
				if (!bc_group) {
					for (int j = 2; j < num_attendance_bc - 3; j++) {
						try {
							if (sh_attendance_bc.getRow(j).getCell(2).toString().equals(attendance[0][i])) {
							}
						} catch (NullPointerException e) {
							continue;
						}
						if (sh_attendance_bc.getRow(j).getCell(2).toString().equals(attendance[0][i])) {
							if (attendance[1][i].equals("TRUE")) {
								sh_attendance_bc.getRow(j).createCell(week).setCellValue("○");
								sh_attendance_bc.getRow(j).createCell(55).setCellValue("");
							} else {
								try {
									sh_attendance_bc.getRow(j).removeCell(sh_attendance_bc.getRow(j).getCell(week));
								} catch (NullPointerException e) {
								}
								sh_attendance_bc.getRow(j).createCell(55).setCellValue(attendance[2][i]);
							}
							break;
						}
					}
				} else {
					for (int j = 2; j < num_attendance_group - 2; j++) {
						try {
							if (sh_attendance_group.getRow(j).getCell(2).toString().equals(attendance[0][i])) {
							}
						} catch (NullPointerException e) {
							continue;
						}
						if (sh_attendance_group.getRow(j).getCell(2).toString().equals(attendance[0][i])) {
							if (attendance[1][i].equals("TRUE")) {
								sh_attendance_group.getRow(j).createCell(week).setCellValue("○");
								sh_attendance_group.getRow(j).createCell(55).setCellValue("");
							} else {
								try {
									sh_attendance_group.getRow(j)
											.removeCell(sh_attendance_group.getRow(j).getCell(week));
								} catch (NullPointerException e) {
								}
								sh_attendance_group.getRow(j).createCell(55).setCellValue(attendance[2][i]);
							}
							break;
						}
					}
				}
				break;
			}

			for (int j = 1; j < num_attendance - 4; j++) {
				try {
					if (sh_attendance.getRow(j).getCell(2).toString().equals(attendance[0][i])) {
					}
				} catch (NullPointerException e) {
					continue;
				}
				if (sh_attendance.getRow(j).getCell(2).toString().equals(attendance[0][i])) {
					if (attendance[1][i].equals("TRUE")) {
						sh_attendance.getRow(j).createCell(week).setCellValue("○");
						sh_attendance.getRow(j).createCell(55).setCellValue("");
					} else {
						try {
							sh_attendance.getRow(j).removeCell(sh_attendance.getRow(j).getCell(week));
						} catch (NullPointerException e) {
						}
						sh_attendance.getRow(j).createCell(55).setCellValue(attendance[2][i]);
					}
					break;
				}
			}
		}

	}

	public void writeHomeAttendance(String type, String[][] attendance) {
		String[] criteria = type.split("#");
		StringBuffer buf = new StringBuffer(criteria[0]);
		String sub = buf.substring(0, 2);
		int idx = Integer.parseInt(criteria[1]) * 2 + 1;
		String dvdd1 = "";
		int dvdd2 = 0;

		for (int i = 0; i < attendance[0].length; i++) {
			boolean bc_group = false;

			switch (sub) {
			case "ge":
				dvdd1 = "general";
				dvdd2 = Integer.parseInt(buf.substring(7));
				for (int j = 1; j < num_attendance_general - 1; j++) {
					if (sh_attendance_general_home.getRow(j).getCell(2).toString().equals(attendance[0][i])) {
						if (attendance[1][i].equals("TRUE")) {
							sh_attendance_general_home.getRow(j).createCell(idx).setCellValue("○");
							sh_attendance_general_home.getRow(j).createCell(idx + 1).setCellValue("");
						} else {
							try {
								sh_attendance_general_home.getRow(j)
										.removeCell(sh_attendance_general_home.getRow(j).getCell(idx));
							} catch (NullPointerException e) {
							}
							sh_attendance_general_home.getRow(j).createCell(idx + 1).setCellValue(attendance[2][i]);
						}
						break;
					}
				}
				break;

			case "gr":
				dvdd1 = "group";
				dvdd2 = Integer.parseInt(buf.substring(5));
				for (int j = 1; j < num_attendance_bc_group; j++) {
					if (sh_attendance_bc_group.getRow(j).getCell(3).toString().equals(attendance[0][i])) {
						bc_group = true;
						break;
					}
				}
				for (int j = 2; j < num_attendance_group - 2; j++) {
					try {
						if (sh_attendance_group_home.getRow(j).getCell(2).toString().equals(attendance[0][i])) {
						}
					} catch (NullPointerException e) {
						continue;
					}
					if (sh_attendance_group_home.getRow(j).getCell(2).toString().equals(attendance[0][i])) {
						if (attendance[1][i].equals("TRUE")) {
							sh_attendance_group_home.getRow(j).createCell(idx).setCellValue("○");
							sh_attendance_group_home.getRow(j).createCell(idx + 1).setCellValue("");
						} else {
							try {
								sh_attendance_group_home.getRow(j)
										.removeCell(sh_attendance_group_home.getRow(j).getCell(idx));
							} catch (NullPointerException e) {
							}
							sh_attendance_group_home.getRow(j).createCell(idx + 1).setCellValue(attendance[2][i]);
						}
						break;
					}
				}
				if (bc_group) {
					for (int j = 2; j < num_attendance_bc_read - 3; j++) {
						try {
							if (sh_attendance_bc_read_home.getRow(j).getCell(2).toString().equals(attendance[0][i])) {
							}
						} catch (NullPointerException e) {
							continue;
						}
						if (sh_attendance_bc_read_home.getRow(j).getCell(2).toString().equals(attendance[0][i])) {
							if (attendance[1][i].equals("TRUE")) {
								sh_attendance_bc_read_home.getRow(j).createCell(idx).setCellValue("○");
								sh_attendance_bc_read_home.getRow(j).createCell(idx + 1).setCellValue("");
							} else {
								try {
									sh_attendance_bc_read_home.getRow(j)
											.removeCell(sh_attendance_bc_read_home.getRow(j).getCell(idx));
								} catch (NullPointerException e) {
								}
								sh_attendance_bc_read_home.getRow(j).createCell(idx + 1).setCellValue(attendance[2][i]);
							}
							break;
						}
					}
				}
				break;

			case "bc":
				dvdd1 = "bc";
				dvdd2 = Integer.parseInt(buf.substring(2));
				for (int j = 1; j < num_attendance_bc_group; j++) {
					if (sh_attendance_bc_group.getRow(j).getCell(3).toString().equals(attendance[0][i])) {
						bc_group = true;
						break;
					}
				}
				for (int j = 2; j < num_attendance_bc_read - 3; j++) {
					try {
						if (sh_attendance_bc_read_home.getRow(j).getCell(2).toString().equals(attendance[0][i])) {
						}
					} catch (NullPointerException e) {
						continue;
					}
					if (sh_attendance_bc_read_home.getRow(j).getCell(2).toString().equals(attendance[0][i])) {
						if (attendance[1][i].equals("TRUE")) {
							sh_attendance_bc_read_home.getRow(j).createCell(idx).setCellValue("○");
							sh_attendance_bc_read_home.getRow(j).createCell(idx + 1).setCellValue("");
						} else {
							try {
								sh_attendance_bc_read_home.getRow(j)
										.removeCell(sh_attendance_bc_read_home.getRow(j).getCell(idx));
							} catch (NullPointerException e) {
							}
							sh_attendance_bc_read_home.getRow(j).createCell(idx + 1).setCellValue(attendance[2][i]);
						}
						break;
					}
				}
				if (!bc_group) {
					for (int j = 2; j < num_attendance_bc - 3; j++) {
						try {
							if (sh_attendance_bc_home.getRow(j).getCell(2).toString().equals(attendance[0][i])) {
							}
						} catch (NullPointerException e) {
							continue;
						}
						if (sh_attendance_bc_home.getRow(j).getCell(2).toString().equals(attendance[0][i])) {
							if (attendance[1][i].equals("TRUE")) {
								sh_attendance_bc_home.getRow(j).createCell(idx).setCellValue("○");
								sh_attendance_bc_home.getRow(j).createCell(idx + 1).setCellValue("");
							} else {
								try {
									sh_attendance_bc_home.getRow(j)
											.removeCell(sh_attendance_bc_home.getRow(j).getCell(idx));
								} catch (NullPointerException e) {
								}
								sh_attendance_bc_home.getRow(j).createCell(idx + 1).setCellValue(attendance[2][i]);
							}
							break;
						}
					}
				} else {
					for (int j = 2; j < num_attendance_group - 2; j++) {
						try {
							if (sh_attendance_group_home.getRow(j).getCell(2).toString().equals(attendance[0][i])) {
							}
						} catch (NullPointerException e) {
							continue;
						}
						if (sh_attendance_group_home.getRow(j).getCell(2).toString().equals(attendance[0][i])) {
							if (attendance[1][i].equals("TRUE")) {
								sh_attendance_group_home.getRow(j).createCell(idx).setCellValue("○");
								sh_attendance_group_home.getRow(j).createCell(idx + 1).setCellValue("");
							} else {
								try {
									sh_attendance_group_home.getRow(j)
											.removeCell(sh_attendance_group_home.getRow(j).getCell(idx));
								} catch (NullPointerException e) {
								}
								sh_attendance_group_home.getRow(j).createCell(idx + 1).setCellValue(attendance[2][i]);
							}
							break;
						}
					}
				}
				break;
			}

			for (int j = 1; j < num_attendance - 4; j++) {
				try {
					if (sh_attendance_home.getRow(j).getCell(2).toString().equals(attendance[0][i])) {
					}
				} catch (NullPointerException e) {
					continue;
				}
				if (sh_attendance_home.getRow(j).getCell(2).toString().equals(attendance[0][i])) {
					if (attendance[1][i].equals("TRUE")) {
						sh_attendance_home.getRow(j).createCell(idx).setCellValue("○");
						sh_attendance_home.getRow(j).createCell(idx + 1).setCellValue("");
					} else {
						try {
							sh_attendance_home.getRow(j).removeCell(sh_attendance_home.getRow(j).getCell(idx));
						} catch (NullPointerException e) {
						}
						sh_attendance_home.getRow(j).createCell(idx + 1).setCellValue(attendance[2][i]);

						if (criteria[1].equals("1") || criteria[1].equals("2") || criteria[1].equals("3")) {
							Boolean flag = false;
							try {
								if (sh_attendance_home.getRow(j).getCell(3).toString().equals("○")) {
									flag = true;
								}
							} catch (NullPointerException e) {
							}
							try {
								if (sh_attendance_home.getRow(j).getCell(5).toString().equals("○")) {
									flag = true;
								}
							} catch (NullPointerException e) {
							}
							try {
								if (sh_attendance_home.getRow(j).getCell(7).toString().equals("○")) {
									flag = true;
								}
							} catch (NullPointerException e) {
							}
							if (flag || (readStatus(dvdd1)[2][dvdd2 - 1].equals("TRUE")
									&& readAttendance(criteria[0])[1][i].equals("○"))) {
								attendance[1][i] = "TRUE";
							}
						}
					}
					break;
				}
			}
		}

		if (criteria[1].equals("1") || criteria[1].equals("2") || criteria[1].equals("3")) {
			writeAttendance(type, attendance);
		}
	}

	public void writeStatus(String type) {
		int week = week_of_year + 1;

		StringBuffer buf = new StringBuffer(type);
		String identifier = buf.substring(buf.length() - 2, buf.length() - 1);
		String sub = buf.substring(0, 2);
		if (!identifier.equals("#")) {
			if (sub.equals("ge")) {
				String[][] list = readStatus("general");
				for (int i = 0; i < list[0].length; i++) {
					if (list[0][i].equals(type)) {
						sh_status_general.getRow(i + 1).getCell(week).setCellValue(true);
						break;
					}
				}
			} else if (sub.equals("gr")) {
				String[][] list = readStatus("group");
				for (int i = 0; i < list[0].length; i++) {
					if (list[0][i].equals(type)) {
						sh_status_group.getRow(i + 1).getCell(week).setCellValue(true);
						break;
					}
				}
			} else if (sub.equals("bc")) {
				String[][] list = readStatus("bc");
				for (int i = 0; i < list[0].length; i++) {
					if (list[0][i].equals(type)) {
						sh_status_bc.getRow(i + 1).getCell(week).setCellValue(true);
						break;
					}
				}
			} else if (sub.equals("fp")) {
				String[][] list = readStatus("fp");
				for (int i = 0; i < list[0].length; i++) {
					if (list[0][i].equals(type)) {
						sh_status_fp.getRow(i + 1).getCell(week).setCellValue(true);
						break;
					}
				}
			}
		} else {
			String[] criteria = type.split("#");
			if (sub.equals("ge")) {
				String[][] list = readStatus(criteria[1] + "#general");
				for (int i = 0; i < list[0].length; i++) {
					if (list[0][i].equals(criteria[0])) {
						sh_status_general_home.getRow(i + 1).getCell(Integer.parseInt(criteria[1]) + 2)
								.setCellValue(true);
						break;
					}
				}
			} else if (sub.equals("gr")) {
				String[][] list = readStatus(criteria[1] + "#group");
				for (int i = 0; i < list[0].length; i++) {
					if (list[0][i].equals(criteria[0])) {
						sh_status_group_home.getRow(i + 1).getCell(Integer.parseInt(criteria[1]) + 2)
								.setCellValue(true);
						break;
					}
				}
			} else if (sub.equals("bc")) {
				String[][] list = readStatus(criteria[1] + "#bc");
				for (int i = 0; i < list[0].length; i++) {
					if (list[0][i].equals(criteria[0])) {
						sh_status_bc_home.getRow(i + 1).getCell(Integer.parseInt(criteria[1]) + 2).setCellValue(true);
						break;
					}
				}
			}
		}
	}

	public String[] readId(String id, String key) {
		String[] info_id = new String[4];

		if (id != null) {
			for (int i = 1; i < num_id; i++) {
				if (sh_id.getRow(i).getCell(1).toString().equals(id)) {
					info_id[0] = sh_id.getRow(i).getCell(0).toString();
					info_id[1] = sh_id.getRow(i).getCell(1).toString();
					info_id[2] = sh_id.getRow(i).getCell(2).toString();
					info_id[3] = sh_id.getRow(i).getCell(3).toString();
					break;
				}
			}
		} else {
			for (int i = 1; i < num_id; i++) {
				if (sh_id.getRow(i).getCell(2).toString().equals(key)) {
					info_id[0] = sh_id.getRow(i).getCell(0).toString();
					info_id[1] = sh_id.getRow(i).getCell(1).toString();
					info_id[2] = sh_id.getRow(i).getCell(2).toString();
					info_id[3] = sh_id.getRow(i).getCell(3).toString();
					break;
				}
			}
		}

		return info_id;
	}

	public String[][] readStatus(String type) {
		String[][] info_status;
		int week = week_of_year + 1;
		int count = 0;

		switch (type) {
		case "main":
			info_status = new String[3][19];
			for (int i = 1; i < num_status_general; i++) {
				if (sh_status_general.getRow(i).getCell(week).getBooleanCellValue()) {
					count++;
				}
			}
			if (count == num_status_general - 1) {
				info_status[2][0] = "TRUE";
			} else if (count == 0) {
				info_status[2][0] = "FALSE";
			} else {
				info_status[2][0] = "ONGOING";
			}
			count = 0;
			for (int i = 1; i < num_status_group; i++) {
				if (sh_status_group.getRow(i).getCell(week).getBooleanCellValue()) {
					count++;
				}
			}
			if (count == num_status_group - 1) {
				info_status[2][1] = "TRUE";
			} else if (count == 0) {
				info_status[2][1] = "FALSE";
			} else {
				info_status[2][1] = "ONGOING";
			}
			count = 0;
			for (int i = 1; i < num_status_bc; i++) {
				if (sh_status_bc.getRow(i).getCell(week).getBooleanCellValue()) {
					count++;
				}
			}
			if (count == num_status_bc - 1) {
				info_status[2][2] = "TRUE";
			} else if (count == 0) {
				info_status[2][2] = "FALSE";
			} else {
				info_status[2][2] = "ONGOING";
			}
			count = 0;
			for (int i = 1; i < num_status_fp; i++) {
				if (sh_status_fp.getRow(i).getCell(week).getBooleanCellValue()) {
					count++;
				}
			}
			if (count == num_status_fp - 1) {
				info_status[2][3] = "TRUE";
			} else if (count == 0) {
				info_status[2][3] = "FALSE";
			} else {
				info_status[2][3] = "ONGOING";
			}
			for (int j = 3; j < 8; j++) {
				count = 0;
				for (int i = 1; i < num_status_general_home; i++) {
					if (sh_status_general_home.getRow(i).getCell(j).getBooleanCellValue()) {
						count++;
					}
				}
				if (count == num_status_general_home - 1) {
					info_status[2][3 * j - 5] = "TRUE";
				} else if (count == 0) {
					info_status[2][3 * j - 5] = "FALSE";
				} else {
					info_status[2][3 * j - 5] = "ONGOING";
				}
				count = 0;
				for (int i = 1; i < num_status_group_home; i++) {
					if (sh_status_group_home.getRow(i).getCell(j).getBooleanCellValue()) {
						count++;
					}
				}
				if (count == num_status_group_home - 1) {
					info_status[2][3 * j - 4] = "TRUE";
				} else if (count == 0) {
					info_status[2][3 * j - 4] = "FALSE";
				} else {
					info_status[2][3 * j - 4] = "ONGOING";
				}
				count = 0;
				for (int i = 1; i < num_status_bc_home; i++) {
					if (sh_status_bc_home.getRow(i).getCell(j).getBooleanCellValue()) {
						count++;
					}
				}
				if (count == num_status_bc_home - 1) {
					info_status[2][3 * j - 3] = "TRUE";
				} else if (count == 0) {
					info_status[2][3 * j - 3] = "FALSE";
				} else {
					info_status[2][3 * j - 3] = "ONGOING";
				}
			}
			break;
		case "general":
			info_status = new String[3][num_status_general - 1];
			for (int i = 1; i < num_status_general; i++) {
				info_status[0][i - 1] = sh_status_general.getRow(i).getCell(1).toString();
				info_status[1][i - 1] = sh_status_general.getRow(i).getCell(2).toString();
				if (sh_status_general.getRow(i).getCell(week).getBooleanCellValue()) {
					info_status[2][i - 1] = "TRUE";
				} else {
					info_status[2][i - 1] = "FALSE";
				}
			}
			break;
		case "group":
			info_status = new String[3][num_status_group - 1];
			for (int i = 1; i < num_status_group; i++) {
				info_status[0][i - 1] = sh_status_group.getRow(i).getCell(1).toString();
				info_status[1][i - 1] = sh_status_group.getRow(i).getCell(2).toString();
				if (sh_status_group.getRow(i).getCell(week).getBooleanCellValue()) {
					info_status[2][i - 1] = "TRUE";
				} else {
					info_status[2][i - 1] = "FALSE";
				}
			}
			break;
		case "bc":
			info_status = new String[3][num_status_bc - 1];
			for (int i = 1; i < num_status_bc; i++) {
				info_status[0][i - 1] = sh_status_bc.getRow(i).getCell(1).toString();
				info_status[1][i - 1] = sh_status_bc.getRow(i).getCell(2).toString();
				if (sh_status_bc.getRow(i).getCell(week).getBooleanCellValue()) {
					info_status[2][i - 1] = "TRUE";
				} else {
					info_status[2][i - 1] = "FALSE";
				}
			}
			break;
		case "fp":
			info_status = new String[3][num_status_fp - 1];
			for (int i = 1; i < num_status_fp; i++) {
				info_status[0][i - 1] = sh_status_fp.getRow(i).getCell(1).toString();
				info_status[1][i - 1] = sh_status_fp.getRow(i).getCell(2).toString();
				if (sh_status_fp.getRow(i).getCell(week).getBooleanCellValue()) {
					info_status[2][i - 1] = "TRUE";
				} else {
					info_status[2][i - 1] = "FALSE";
				}
			}
			break;
		default:
			String[] criteria = type.split("#");
			int idx = Integer.parseInt(criteria[0]) + 2;
			XSSFSheet sh;
			int n;
			if (criteria[1].equals("general")) {
				sh = sh_status_general_home;
				n = num_status_general_home;
			} else if (criteria[1].equals("group")) {
				sh = sh_status_group_home;
				n = num_status_group_home;
			} else {
				sh = sh_status_bc_home;
				n = num_status_bc_home;
			}
			info_status = new String[3][n - 1];
			for (int i = 1; i < n; i++) {
				info_status[0][i - 1] = sh.getRow(i).getCell(1).toString();
				info_status[1][i - 1] = sh.getRow(i).getCell(2).toString();
				if (sh.getRow(i).getCell(idx).getBooleanCellValue()) {
					info_status[2][i - 1] = "TRUE";
				} else {
					info_status[2][i - 1] = "FALSE";
				}
			}
			break;
		}

		return info_status;
	}

	public String[] readLogs() {
		String[] info_logs;
		int nlog = 1;
		if (num_logs >= 300) {
			info_logs = new String[299];
			nlog = num_logs - 299;
		} else {
			info_logs = new String[num_logs - 1];
		}

		int count = 0;
		for (int i = num_logs - 1; i >= nlog; i--) {
			switch (sh_logs.getRow(i).getCell(2).toString()) {
			case "TurnOn":
				info_logs[count] = "[" + sh_logs.getRow(i).getCell(0).toString() + "] "
						+ sh_logs.getRow(i).getCell(2).toString();
				break;
			case "TurnOff":
				info_logs[count] = "[" + sh_logs.getRow(i).getCell(0).toString() + "] "
						+ sh_logs.getRow(i).getCell(2).toString() + " <" + sh_logs.getRow(i).getCell(4) + ">";
				break;
			case "Check":
				info_logs[count] = "[" + sh_logs.getRow(i).getCell(0).toString() + "] "
						+ sh_logs.getRow(i).getCell(1).toString() + " " + sh_logs.getRow(i).getCell(2).toString() + " "
						+ sh_logs.getRow(i).getCell(3).toString();
				break;
			case "FpCheck":
				info_logs[count] = "[" + sh_logs.getRow(i).getCell(0).toString() + "] "
						+ sh_logs.getRow(i).getCell(1).toString() + " " + sh_logs.getRow(i).getCell(2).toString() + " "
						+ sh_logs.getRow(i).getCell(3).toString();
				break;
			case "FpAdd":
				info_logs[count] = "[" + sh_logs.getRow(i).getCell(0).toString() + "] "
						+ sh_logs.getRow(i).getCell(1).toString() + " " + sh_logs.getRow(i).getCell(2).toString() + " "
						+ sh_logs.getRow(i).getCell(3).toString();
				break;
			case "Add":
				info_logs[count] = "[" + sh_logs.getRow(i).getCell(0).toString() + "] "
						+ sh_logs.getRow(i).getCell(1).toString() + " " + sh_logs.getRow(i).getCell(2).toString() + " "
						+ sh_logs.getRow(i).getCell(3).toString();
				break;
			default:
				info_logs[count] = "[" + sh_logs.getRow(i).getCell(0).toString() + "] "
						+ sh_logs.getRow(i).getCell(1).toString() + " " + sh_logs.getRow(i).getCell(2).toString();
				break;
			}
			count++;
		}

		return info_logs;
	}

	public String[][] readAttendance(String type) {
		String[][] info_attendance;
		int week = week_of_year + 1;
		int count = 0;
		boolean isIndex = false;
		int list_index = 0;
		int list_count = 0;
		int category_count = 0;
		int sum_count = 0;
		String list_type = "";

		StringBuffer buf = new StringBuffer(type);
		String sub = buf.substring(0, 2);
		String identifier = buf.substring(buf.length() - 2, buf.length() - 1);

		switch (sub) {
		case "ge":
			for (int i = 1; i < num_attendance_general - 1; i++) {
				try {
					if (!sh_attendance_general.getRow(i).getCell(0).toString().equals("")) {
						count++;
						if (!isIndex) {
							list_index = i;
						}
						if (!identifier.equals("#"))
							list_type = "general" + count;
						else
							list_type = "general" + count + buf.substring(buf.length() - 2, buf.length());
					}
				} catch (NullPointerException e) {
				}
				if (type.equals(list_type)) {
					isIndex = true;
					list_count++;
				} else if (isIndex) {
					break;
				}
			}

			info_attendance = new String[3][list_count];
			if (!identifier.equals("#")) {
				for (int i = 0; i < list_count; i++) {
					info_attendance[0][i] = sh_attendance_general.getRow(list_index + i).getCell(2).toString();
					try {
						info_attendance[1][i] = sh_attendance_general.getRow(list_index + i).getCell(week).toString();
					} catch (NullPointerException e) {
						info_attendance[1][i] = "NULL";
					}
					try {
						info_attendance[2][i] = sh_attendance_general.getRow(list_index + i).getCell(55).toString();
					} catch (NullPointerException e) {
						info_attendance[2][i] = "NULL";
					}
				}
			} else {
				int idx = Integer.parseInt(buf.substring(buf.length() - 1, buf.length())) * 2 + 1;
				for (int i = 0; i < list_count; i++) {
					info_attendance[0][i] = sh_attendance_general_home.getRow(list_index + i).getCell(2).toString();
					try {
						info_attendance[1][i] = sh_attendance_general_home.getRow(list_index + i).getCell(idx)
								.toString();
					} catch (NullPointerException e) {
						info_attendance[1][i] = "NULL";
					}
					try {
						info_attendance[2][i] = sh_attendance_general_home.getRow(list_index + i).getCell(idx + 1)
								.toString();
					} catch (NullPointerException e) {
						info_attendance[2][i] = "NULL";
					}
				}
			}
			break;

		case "gr":
			StringBuffer buf_group;
			for (int i = 1; i < num_attendance_group - 2; i++) {
				try {
					buf_group = new StringBuffer(sh_attendance_group.getRow(i).getCell(0).toString());
					if (buf_group.charAt(1) == ' ') {
						count++;
						if (!isIndex) {
							list_index = i + 1;
						}
						if (!identifier.equals("#"))
							list_type = "group" + count;
						else
							list_type = "group" + count + buf.substring(buf.length() - 2, buf.length());
					}
				} catch (NullPointerException | StringIndexOutOfBoundsException e) {
					buf_group = new StringBuffer("@");
				}
				if (type.equals(list_type)) {
					if (buf_group.charAt(0) != '@') {
						category_count++;
					}
					if (!isIndex) {
						list_count--;
						category_count--;
						isIndex = true;
					}
					list_count++;
				} else if (isIndex) {
					list_count--;
					category_count--;
					break;
				}
			}

			info_attendance = new String[3][list_count + category_count];
			category_count = 0;
			if (!identifier.equals("#")) {
				for (int i = 0; i < list_count; i++) {
					try {
						if (!sh_attendance_group.getRow(list_index + i).getCell(0).toString().equals("")) {
							info_attendance[0][i + category_count] = "CATEGORY";
							info_attendance[1][i + category_count] = sh_attendance_group.getRow(list_index + i)
									.getCell(0).toString();
							info_attendance[2][i + category_count] = "NULL";
							category_count++;
						}
					} catch (NullPointerException e) {
					}
					info_attendance[0][i + category_count] = sh_attendance_group.getRow(list_index + i).getCell(2)
							.toString();
					try {
						info_attendance[1][i + category_count] = sh_attendance_group.getRow(list_index + i)
								.getCell(week).toString();
					} catch (NullPointerException e) {
						info_attendance[1][i + category_count] = "NULL";
					}
					try {
						info_attendance[2][i + category_count] = sh_attendance_group.getRow(list_index + i).getCell(55)
								.toString();
					} catch (NullPointerException e) {
						info_attendance[2][i + category_count] = "NULL";
					}
				}
			} else {
				int idx = Integer.parseInt(buf.substring(buf.length() - 1, buf.length())) * 2 + 1;
				for (int i = 0; i < list_count; i++) {
					try {
						if (!sh_attendance_group_home.getRow(list_index + i).getCell(0).toString().equals("")) {
							info_attendance[0][i + category_count] = "CATEGORY";
							info_attendance[1][i + category_count] = sh_attendance_group_home.getRow(list_index + i)
									.getCell(0).toString();
							info_attendance[2][i + category_count] = "NULL";
							category_count++;
						}
					} catch (NullPointerException e) {
					}
					info_attendance[0][i + category_count] = sh_attendance_group_home.getRow(list_index + i).getCell(2)
							.toString();
					try {
						info_attendance[1][i + category_count] = sh_attendance_group_home.getRow(list_index + i)
								.getCell(idx).toString();
					} catch (NullPointerException e) {
						info_attendance[1][i + category_count] = "NULL";
					}
					try {
						info_attendance[2][i + category_count] = sh_attendance_group_home.getRow(list_index + i)
								.getCell(idx + 1).toString();
					} catch (NullPointerException e) {
						info_attendance[2][i + category_count] = "NULL";
					}
				}
			}
			break;

		case "bc":
			for (int i = 1; i < num_attendance_bc_read - 2; i++) {
				try {
					if (!sh_attendance_bc_read.getRow(i).getCell(0).toString().equals("")) {
						count++;
						if (!isIndex) {
							list_index = i + 1;
						}
						if (!identifier.equals("#"))
							list_type = "bc" + count;
						else
							list_type = "bc" + count + buf.substring(buf.length() - 2, buf.length());
					}
				} catch (NullPointerException e) {
				}
				if (type.equals(list_type)) {
					try {
						if (sh_attendance_bc_read.getRow(i).getCell(1).toString().equals("소계")) {
							sum_count++;
						} else if (!sh_attendance_bc_read.getRow(i).getCell(1).toString().equals("")) {
							category_count++;
						}
					} catch (NullPointerException e) {
					}
					if (!isIndex) {
						list_count--;
						isIndex = true;
					}
					list_count++;
				} else if (isIndex) {
					list_count--;
					category_count--;
					break;
				}
			}

			info_attendance = new String[3][list_count + category_count - sum_count];
			category_count = 0;
			sum_count = 0;
			if (!identifier.equals("#")) {
				for (int i = 0; i < list_count; i++) {
					try {
						if (sh_attendance_bc_read.getRow(list_index + i).getCell(1).toString().equals("소계")) {
							sum_count++;
							continue;
						} else if (!sh_attendance_bc_read.getRow(list_index + i).getCell(1).toString().equals("")) {
							info_attendance[0][i + category_count - sum_count] = "CATEGORY";
							info_attendance[1][i + category_count - sum_count] = sh_attendance_bc_read
									.getRow(list_index + i).getCell(1).toString();
							info_attendance[2][i + category_count - sum_count] = "NULL";
							category_count++;
						}
					} catch (NullPointerException e) {
					}
					info_attendance[0][i + category_count - sum_count] = sh_attendance_bc_read.getRow(list_index + i)
							.getCell(2).toString();
					try {
						info_attendance[1][i + category_count - sum_count] = sh_attendance_bc_read
								.getRow(list_index + i).getCell(week).toString();
					} catch (NullPointerException e) {
						info_attendance[1][i + category_count - sum_count] = "NULL";
					}
					try {
						info_attendance[2][i + category_count - sum_count] = sh_attendance_bc_read
								.getRow(list_index + i).getCell(55).toString();
					} catch (NullPointerException e) {
						info_attendance[2][i + category_count - sum_count] = "NULL";
					}
				}
			} else {
				int idx = Integer.parseInt(buf.substring(buf.length() - 1, buf.length())) * 2 + 1;
				for (int i = 0; i < list_count; i++) {
					try {
						if (sh_attendance_bc_read_home.getRow(list_index + i).getCell(1).toString().equals("소계")) {
							sum_count++;
							continue;
						} else if (!sh_attendance_bc_read_home.getRow(list_index + i).getCell(1).toString()
								.equals("")) {
							info_attendance[0][i + category_count - sum_count] = "CATEGORY";
							info_attendance[1][i + category_count - sum_count] = sh_attendance_bc_read_home
									.getRow(list_index + i).getCell(1).toString();
							info_attendance[2][i + category_count - sum_count] = "NULL";
							category_count++;
						}
					} catch (NullPointerException e) {
					}
					info_attendance[0][i + category_count - sum_count] = sh_attendance_bc_read_home
							.getRow(list_index + i).getCell(2).toString();
					try {
						info_attendance[1][i + category_count - sum_count] = sh_attendance_bc_read_home
								.getRow(list_index + i).getCell(idx).toString();
					} catch (NullPointerException e) {
						info_attendance[1][i + category_count - sum_count] = "NULL";
					}
					try {
						info_attendance[2][i + category_count - sum_count] = sh_attendance_bc_read_home
								.getRow(list_index + i).getCell(idx + 1).toString();
					} catch (NullPointerException e) {
						info_attendance[2][i + category_count - sum_count] = "NULL";
					}
				}

			}
			break;

		default:
			info_attendance = new String[0][0];
			break;
		}

		return info_attendance;
	}

	public String[][] readFpAttendance(String type) {
		String[][] info_fpattendance = new String[7][num_fpattendance_attendance - 1];
		int week = week_of_year + 3;

		for (int i = 0; i < num_fpattendance_attendance - 1; i++) {
			info_fpattendance[0][i] = sh_fpattendance_attendance.getRow(i + 1).getCell(0).toString();
			info_fpattendance[1][i] = sh_fpattendance_attendance.getRow(i + 1).getCell(1).toString();
			info_fpattendance[2][i] = sh_fpattendance_attendance.getRow(i + 1).getCell(2).toString();
			try {
				info_fpattendance[3][i] = sh_fpattendance_attendance.getRow(i + 1).getCell(3).toString();
			} catch (NullPointerException e) {
				info_fpattendance[3][i] = "NULL";
			}
			try {
				info_fpattendance[4][i] = sh_fpattendance_attendance.getRow(i + 1).getCell(4).toString();
			} catch (NullPointerException e) {
				info_fpattendance[4][i] = "NULL";
			}
			info_fpattendance[5][i] = sh_fpattendance_attendance.getRow(i + 1).getCell(week).toString();
			String tt = getObject(sh_fpattendance_attendance.getRow(i + 1).getCell(3).toString());
			int idx = 0;
			for (int j = 1; j <= 38; j = j + 2) {
				if (info_fpattendance[3][i].equals("fp" + j)) {
					idx = (j + 1) / 2;
					break;
				}
			}
			if (idx <= 7 || idx == 16) {
				info_fpattendance[6][i] = tt.substring(0, 2);
			} else if (idx == 13 || idx == 17 || idx == 19) {
				info_fpattendance[6][i] = tt.substring(3, 6);
			} else {
				info_fpattendance[6][i] = tt.substring(3, 5);
			}
		}

		return info_fpattendance;
	}

	public String[] readFpEtcAttendance(String type) {
		int idx = 0;
		for (int i = 1; i <= 38; i = i + 2) {
			if (type.equals("fp" + i)) {
				idx = (i + 1) / 2;
				break;
			}
		}

		XSSFSheet sh_etc = wb_fpattendance.getSheetAt(idx);
		int num_etc = sh_etc.getPhysicalNumberOfRows() - 1;
		String[] info_fpattendance_etc = new String[num_etc];
		for (int i = 0; i < num_etc; i++) {
			info_fpattendance_etc[i] = sh_etc.getRow(i + 1).getCell(0).toString();
		}

		return info_fpattendance_etc;
	}

	public String[][] readSearchResult(String input) {
		int input_length = input.length();
		int count = 0;
		for (int i = 0; i < num_fpattendance_attendance - 1; i++) {
			String t = sh_fpattendance_attendance.getRow(i + 1).getCell(0).toString();
			if (t.length() >= input_length && t.substring(0, input_length).equals(input)) {
				count++;
			}
		}

		String[][] info_searchresult;
		if (count > 5) {
			info_searchresult = new String[4][5];
		} else {
			info_searchresult = new String[4][count];
		}
		count = 0;

		for (int i = 0; i < num_fpattendance_attendance - 1; i++) {
			String t = sh_fpattendance_attendance.getRow(i + 1).getCell(0).toString();
			if (t.length() >= input_length && t.substring(0, input_length).equals(input)) {
				info_searchresult[0][count] = sh_fpattendance_attendance.getRow(i + 1).getCell(0).toString();
				info_searchresult[1][count] = sh_fpattendance_attendance.getRow(i + 1).getCell(1).toString();
				String tt = getObject(sh_fpattendance_attendance.getRow(i + 1).getCell(2).toString());
				int idx = 0;
				for (int j = 1; j <= 38; j = j + 2) {
					if (sh_fpattendance_attendance.getRow(i + 1).getCell(2).toString().equals("fp" + j)) {
						idx = (j + 1) / 2;
						break;
					}
				}
				if (idx == 0) {
					info_searchresult[2][count] = "조 없음";
				} else if (idx <= 7 || idx == 16) {
					info_searchresult[2][count] = tt.substring(0, 2);
				} else if (idx == 13 || idx == 17 || idx == 19) {
					info_searchresult[2][count] = tt.substring(3, 6);
				} else {
					info_searchresult[2][count] = tt.substring(3, 5);
				}
				info_searchresult[3][count] = sh_fpattendance_attendance.getRow(i + 1).getCell(2).toString();
				count++;
				if (count == 5) {
					break;
				}
			}
		}

		return info_searchresult;
	}

	public String[] readFpSearchAttendance(String type) {
		int count = 0;
		int week = week_of_year + 3;
		for (int i = 0; i < num_fpattendance_attendance - 1; i++) {
			String t1 = sh_fpattendance_attendance.getRow(i + 1).getCell(2).toString();
			String t2 = "NULL";
			try {
				t2 = sh_fpattendance_attendance.getRow(i + 1).getCell(3).toString();
			} catch (NullPointerException e) {
			}
			String t3 = sh_fpattendance_attendance.getRow(i + 1).getCell(week).toString();
			if (!t1.equals(t2) && t2.equals(type) && t3.equals("2.0")) {
				count++;
			}
		}

		String[] info_fpattendance_search = new String[count];
		count = 0;

		for (int i = 0; i < num_fpattendance_attendance - 1; i++) {
			String t1 = sh_fpattendance_attendance.getRow(i + 1).getCell(2).toString();
			String t2 = "NULL";
			try {
				t2 = sh_fpattendance_attendance.getRow(i + 1).getCell(3).toString();
			} catch (NullPointerException e) {
			}
			String t3 = sh_fpattendance_attendance.getRow(i + 1).getCell(week).toString();
			if (!t1.equals(t2) && t2.equals(type) && t3.equals("2.0")) {
				info_fpattendance_search[count] = sh_fpattendance_attendance.getRow(i + 1).getCell(0).toString();
				count++;
			}
		}

		return info_fpattendance_search;
	}

	public void writeFpAttendance(String type, String[] names, String[] contents, int[] checks, String[] etc,
			String[] search) {
		int week = week_of_year + 3;

		for (int i = 0; i < names.length; i++) {
			for (int j = 1; j < num_fpattendance_attendance; j++) {
				if (names[i].equals(sh_fpattendance_attendance.getRow(j).getCell(0).toString())) {
					String t1 = "NULL";
					try {
						t1 = sh_fpattendance_attendance.getRow(j).getCell(3).toString();
					} catch (NullPointerException e) {
					}
					String t2 = sh_fpattendance_attendance.getRow(j).getCell(week).toString();
					if (!((!t1.equals(type)) && t2.equals("2.0"))) {
						sh_fpattendance_attendance.getRow(j).createCell(3).setCellValue(type);
						if (checks[i] == 2) {
							sh_fpattendance_attendance.getRow(j).createCell(4).setCellValue("");
						} else {
							sh_fpattendance_attendance.getRow(j).createCell(4).setCellValue(contents[i]);
						}
						sh_fpattendance_attendance.getRow(j).getCell(week).setCellValue(checks[i]);
					}
					break;
				}
			}
		}

		for (int i = 0; i < search.length; i++) {
			for (int j = 1; j < num_fpattendance_attendance; j++) {
				if (search[i].equals(sh_fpattendance_attendance.getRow(j).getCell(0).toString())) {
					sh_fpattendance_attendance.getRow(j).createCell(3).setCellValue(type);
					sh_fpattendance_attendance.getRow(j).createCell(4).setCellValue("");
					sh_fpattendance_attendance.getRow(j).getCell(week).setCellValue(2);
					break;
				}
			}
		}
		for (int i = 1; i < num_fpattendance_attendance; i++) {
			String t1 = sh_fpattendance_attendance.getRow(i).getCell(2).toString();
			String t2 = "NULL";
			try {
				t2 = sh_fpattendance_attendance.getRow(i).getCell(3).toString();
			} catch (NullPointerException e) {
			}
			String t3 = sh_fpattendance_attendance.getRow(i).getCell(week).toString();
			if (!t1.equals(t2) && t2.equals(type) && t3.equals("2.0")) {
				boolean keep = false;
				for (int j = 0; j < search.length; j++) {
					if (search[j].equals(sh_fpattendance_attendance.getRow(i).getCell(0).toString())) {
						keep = true;
						break;
					}
				}
				if (!keep) {
					sh_fpattendance_attendance.getRow(i).getCell(week).setCellValue(0);
				}
			}
		}

		int idx;
		if (type.length() == 3) {
			idx = Integer.valueOf(type.substring(2, 3));
		} else {
			idx = Integer.valueOf(type.substring(2, 4));
		}
		idx = (idx + 1) / 2;
		XSSFSheet sheet = wb_fpattendance.getSheetAt(idx);
		int num_sheet = sheet.getPhysicalNumberOfRows();
		for (int i = 1; i < num_sheet; i++) {
			sheet.removeRow(sheet.getRow(i));
		}
		for (int i = 0; i < etc.length; i++) {
			sheet.createRow(i + 1).createCell(0).setCellValue(etc[i]);
		}
	}

	public String[][] readFpBeliever(String type) {
		int idx = 0;
		for (int i = 2; i <= 38; i = i + 2) {
			if (type.equals("fp" + i)) {
				idx = i / 2 - 1;
				break;
			}
		}

		XSSFSheet sh_believer = wb_fpbeliever.getSheetAt(idx);
		int num_believer = sh_believer.getPhysicalNumberOfRows() - 1;
		int week = week_of_year + 3;
		String[][] info_fpbeliever = new String[6][num_believer];

		for (int i = 0; i < num_believer; i++) {
			boolean isConducted = true;
			for (int j = 0; j < num_fpattendance_attendance - 1; j++) {
				if (sh_believer.getRow(i + 1).getCell(0).toString()
						.equals(sh_fpattendance_attendance.getRow(j + 1).getCell(0).toString())) {
					if (sh_fpattendance_attendance.getRow(j + 1).getCell(3).toString().equals("fp" + (idx * 2 + 1))
							&& sh_fpattendance_attendance.getRow(j + 1).getCell(week).toString().equals("2.0")) {
						info_fpbeliever[0][i] = sh_believer.getRow(i + 1).getCell(0).toString();
						isConducted = false;
					} else {
						info_fpbeliever[0][i] = "---";
					}
					break;
				}
			}
			if (isConducted) {
				for (int j = 0; j < wb_fpattendance.getSheetAt(idx + 1).getPhysicalNumberOfRows() - 1; j++) {
					if (sh_believer.getRow(i + 1).getCell(0).toString()
							.equals(wb_fpattendance.getSheetAt(idx + 1).getRow(j + 1).getCell(0).toString())) {
						info_fpbeliever[0][i] = sh_believer.getRow(i + 1).getCell(0).toString();
						isConducted = false;
						break;
					}
				}
			}
			if (isConducted) {
				info_fpbeliever[0][i] = "---";
			}

			info_fpbeliever[1][i] = sh_believer.getRow(i + 1).getCell(1).toString();
			try {
				info_fpbeliever[2][i] = sh_believer.getRow(i + 1).getCell(2).toString();
			} catch (NullPointerException e) {
				info_fpbeliever[2][i] = "";
			}
			try {
				info_fpbeliever[3][i] = sh_believer.getRow(i + 1).getCell(3).toString();
			} catch (NullPointerException e) {
				info_fpbeliever[3][i] = "na";
			}
			try {
				info_fpbeliever[4][i] = sh_believer.getRow(i + 1).getCell(4).toString();
			} catch (NullPointerException e) {
				info_fpbeliever[4][i] = "";
			}
			info_fpbeliever[5][i] = sh_believer.getRow(i + 1).getCell(5).toString();
		}

		return info_fpbeliever;
	}

	public String[][] readFpWordMovement(String type) {
		int idx = 0;
		for (int i = 2; i <= 38; i = i + 2) {
			if (type.equals("fp" + i)) {
				idx = i / 2 - 1;
				break;
			}
		}

		XSSFSheet sh_wordmovement = wb_fpwordmovement.getSheetAt(idx);
		int num_wordmovement = sh_wordmovement.getPhysicalNumberOfRows() - 1;
		int week = week_of_year + 3;
		String[][] info_fpwordmovement = new String[4][num_wordmovement];

		for (int i = 0; i < num_wordmovement; i++) {
			boolean isConducted = true;
			for (int j = 0; j < num_fpattendance_attendance - 1; j++) {
				if (sh_wordmovement.getRow(i + 1).getCell(0).toString()
						.equals(sh_fpattendance_attendance.getRow(j + 1).getCell(0).toString())) {
					if (sh_fpattendance_attendance.getRow(j + 1).getCell(3).toString().equals("fp" + (idx * 2 + 1))
							&& sh_fpattendance_attendance.getRow(j + 1).getCell(week).toString().equals("2.0")) {
						info_fpwordmovement[0][i] = sh_wordmovement.getRow(i + 1).getCell(0).toString();
						isConducted = false;
					} else {
						info_fpwordmovement[0][i] = "---";
					}
					break;
				}
			}
			if (isConducted) {
				for (int j = 0; j < wb_fpattendance.getSheetAt(idx + 1).getPhysicalNumberOfRows() - 1; j++) {
					if (sh_wordmovement.getRow(i + 1).getCell(0).toString()
							.equals(wb_fpattendance.getSheetAt(idx + 1).getRow(j + 1).getCell(0).toString())) {
						info_fpwordmovement[0][i] = sh_wordmovement.getRow(i + 1).getCell(0).toString();
						isConducted = false;
						break;
					}
				}
			}
			if (isConducted) {
				info_fpwordmovement[0][i] = "---";
			}

			info_fpwordmovement[1][i] = sh_wordmovement.getRow(i + 1).getCell(1).toString();
			try {
				info_fpwordmovement[2][i] = sh_wordmovement.getRow(i + 1).getCell(2).toString();
			} catch (NullPointerException e) {
				info_fpwordmovement[2][i] = "na";
			}
			try {
				info_fpwordmovement[3][i] = sh_wordmovement.getRow(i + 1).getCell(3).toString();
			} catch (NullPointerException e) {
				info_fpwordmovement[3][i] = "";
			}
		}

		return info_fpwordmovement;
	}

	public String[][] readSearchResultForFruits(String type, String input) {
		int week = week_of_year + 3;
		int idx = Integer.valueOf(type.substring(2, type.length())) - 1;
		int idx_adj = (idx + 1) / 2;
		int input_length = input.length();
		int count = 0;
		for (int i = 0; i < num_fpattendance_attendance - 1; i++) {
			String t1 = sh_fpattendance_attendance.getRow(i + 1).getCell(0).toString();
			String t2 = sh_fpattendance_attendance.getRow(i + 1).getCell(3).toString();
			String t3 = sh_fpattendance_attendance.getRow(i + 1).getCell(week).toString();
			if (t1.length() >= input_length && t1.substring(0, input_length).equals(input) && t2.equals("fp" + idx)
					&& t3.equals("2.0")) {
				count++;
			}
		}
		for (int i = 0; i < wb_fpattendance.getSheetAt(idx_adj).getPhysicalNumberOfRows() - 1; i++) {
			String t4 = wb_fpattendance.getSheetAt(idx_adj).getRow(i + 1).getCell(0).toString();
			if (t4.length() >= input_length && t4.substring(0, input_length).equals(input)) {
				count++;
			}
		}

		String[][] info_searchresult;
		if (count > 5) {
			info_searchresult = new String[2][5];
		} else {
			info_searchresult = new String[2][count];
		}
		count = 0;

		for (int i = 0; i < num_fpattendance_attendance - 1; i++) {
			String t1 = sh_fpattendance_attendance.getRow(i + 1).getCell(0).toString();
			String t2 = sh_fpattendance_attendance.getRow(i + 1).getCell(3).toString();
			String t3 = sh_fpattendance_attendance.getRow(i + 1).getCell(week).toString();
			if (t1.length() >= input_length && t1.substring(0, input_length).equals(input) && t2.equals("fp" + idx)
					&& t3.equals("2.0")) {
				info_searchresult[0][count] = sh_fpattendance_attendance.getRow(i + 1).getCell(0).toString();
				info_searchresult[1][count] = sh_fpattendance_attendance.getRow(i + 1).getCell(1).toString();
				count++;
				if (count == 5) {
					break;
				}
			}
		}
		if (count != 5) {
			for (int i = 0; i < wb_fpattendance.getSheetAt(idx_adj).getPhysicalNumberOfRows() - 1; i++) {
				String t4 = wb_fpattendance.getSheetAt(idx_adj).getRow(i + 1).getCell(0).toString();
				if (t4.length() >= input_length && t4.substring(0, input_length).equals(input)) {
					info_searchresult[0][count] = wb_fpattendance.getSheetAt(idx_adj).getRow(i + 1).getCell(0)
							.toString();
					info_searchresult[1][count] = "기타";
					count++;
					if (count == 5) {
						break;
					}
				}
			}
		}

		return info_searchresult;
	}

	public void writeFpFruits(String type, String[][] believer, String[][] wordmovement) {
		int idx;
		if (type.length() == 3) {
			idx = Integer.valueOf(type.substring(2, 3));
		} else {
			idx = Integer.valueOf(type.substring(2, 4));
		}
		idx = (idx - 2) / 2;
		XSSFSheet sheet_b = wb_fpbeliever.getSheetAt(idx);
		XSSFSheet sheet_wm = wb_fpwordmovement.getSheetAt(idx);
		int num_sheet_b = sheet_b.getPhysicalNumberOfRows();
		int num_sheet_wm = sheet_wm.getPhysicalNumberOfRows();
		for (int i = 1; i < num_sheet_b; i++) {
			sheet_b.removeRow(sheet_b.getRow(i));
		}
		for (int i = 1; i < num_sheet_wm; i++) {
			sheet_wm.removeRow(sheet_wm.getRow(i));
		}
		for (int i = 0; i < believer[0].length; i++) {
			sheet_b.createRow(i + 1).createCell(0).setCellValue(believer[0][i]);
			sheet_b.getRow(i + 1).createCell(1).setCellValue(believer[1][i]);
			if (believer[2][i].equals("모름") || believer[2][i].equals("없음") || believer[2][i].equals("x")
					|| believer[2][i].equals("X") || believer[2][i].equals("-") || believer[2][i].equals(".")
					|| believer[2][i].equals(" ")) {
				sheet_b.getRow(i + 1).createCell(2).setCellValue("");
			} else {
				sheet_b.getRow(i + 1).createCell(2).setCellValue(believer[2][i]);
			}
			if (believer[3][i].equals("0")) {
				sheet_b.getRow(i + 1).createCell(3).setCellValue("");
			} else {
				sheet_b.getRow(i + 1).createCell(3).setCellValue(believer[3][i]);
			}
			if (believer[4][i].equals("0")) {
				sheet_b.getRow(i + 1).createCell(4).setCellValue("");
			} else {
				sheet_b.getRow(i + 1).createCell(4).setCellValue(believer[4][i]);
			}
			sheet_b.getRow(i + 1).createCell(5).setCellValue(believer[5][i]);
		}
		for (int i = 0; i < wordmovement[0].length; i++) {
			sheet_wm.createRow(i + 1).createCell(0).setCellValue(wordmovement[0][i]);
			sheet_wm.getRow(i + 1).createCell(1).setCellValue(wordmovement[1][i]);
			if (wordmovement[2][i].equals("0")) {
				sheet_wm.getRow(i + 1).createCell(2).setCellValue("");
			} else {
				sheet_wm.getRow(i + 1).createCell(2).setCellValue(wordmovement[2][i]);
			}
			if (wordmovement[3][i].equals("모름") || wordmovement[3][i].equals("없음") || wordmovement[3][i].equals("x")
					|| wordmovement[3][i].equals("X") || wordmovement[3][i].equals("-")
					|| wordmovement[3][i].equals(".") || wordmovement[3][i].equals(" ")) {
				sheet_wm.getRow(i + 1).createCell(3).setCellValue("");
			} else {
				sheet_wm.getRow(i + 1).createCell(3).setCellValue(wordmovement[3][i]);
			}
		}
	}

	public void writeFpStatus(String type) {
		int idx = Integer.valueOf(type.substring(2, type.length())) + 1;
		int week = week_of_year + 1;
		sh_status_fp.getRow(idx).getCell(week).setCellValue(false);
	}

	public String[] getConsideration(String type, String[] believer, String[] wordmovement) {
		int week = week_of_year + 3;
		int idx = Integer.valueOf(type.substring(2, type.length())) - 1;
		int idx_adj = (idx + 1) / 2;

		String[] consideration = new String[3];
		consideration[0] = "";
		consideration[1] = "";
		consideration[2] = "0";

		boolean csd = false;
		for (int i = 0; i < believer.length; i++) {
			csd = false;
			for (int j = 1; j < num_fpattendance_attendance; j++) {
				String t1 = sh_fpattendance_attendance.getRow(j).getCell(0).toString();
				if (t1.equals(believer[i])) {
					String t2 = sh_fpattendance_attendance.getRow(j).getCell(3).toString();
					String t3 = sh_fpattendance_attendance.getRow(j).getCell(week).toString();
					if (t2.equals("fp" + idx) && t3.equals("2.0")) {
						csd = true;
					} else {
						for (int k = 1; k < wb_fpattendance.getSheetAt(idx_adj).getPhysicalNumberOfRows(); k++) {
							String t4 = wb_fpattendance.getSheetAt(idx_adj).getRow(k).getCell(0).toString();
							if (t4.equals(believer[i])) {
								csd = true;
								break;
							}
						}
					}
					break;
				}
			}
			if (!csd) {
				for (int k = 1; k < wb_fpattendance.getSheetAt(idx_adj).getPhysicalNumberOfRows(); k++) {
					String t4 = wb_fpattendance.getSheetAt(idx_adj).getRow(k).getCell(0).toString();
					if (t4.equals(believer[i])) {
						csd = true;
						break;
					}
				}
			}
			if (!csd) {
				consideration[0] = believer[i];
				consideration[1] = "believer";
				consideration[2] = "1";
				break;
			}
		}

		if (csd) {
			for (int i = 0; i < wordmovement.length; i++) {
				csd = false;
				for (int j = 1; j < num_fpattendance_attendance; j++) {
					String t1 = sh_fpattendance_attendance.getRow(j).getCell(0).toString();
					if (t1.equals(wordmovement[i])) {
						String t2 = sh_fpattendance_attendance.getRow(j).getCell(3).toString();
						String t3 = sh_fpattendance_attendance.getRow(j).getCell(week).toString();
						if (t2.equals("fp" + idx) && t3.equals("2.0")) {
							csd = true;
						} else {
							for (int k = 1; k < wb_fpattendance.getSheetAt(idx_adj).getPhysicalNumberOfRows(); k++) {
								String t4 = wb_fpattendance.getSheetAt(idx_adj).getRow(k).getCell(0).toString();
								if (t4.equals(wordmovement[i])) {
									csd = true;
									break;
								}
							}
						}
						break;
					}
				}
				if (!csd) {
					for (int k = 1; k < wb_fpattendance.getSheetAt(idx_adj).getPhysicalNumberOfRows(); k++) {
						String t4 = wb_fpattendance.getSheetAt(idx_adj).getRow(k).getCell(0).toString();
						if (t4.equals(wordmovement[i])) {
							csd = true;
							break;
						}
					}
				}
				if (!csd) {
					consideration[0] = wordmovement[i];
					consideration[1] = "wordmovement";
					consideration[2] = "1";
					break;
				}
			}
		}

		return consideration;
	}

	public boolean getFpDuplication(String input) {
		boolean dup = false;
		for (int i = 0; i < num_fpattendance_attendance - 1; i++) {
			if (input.equals(sh_fpattendance_attendance.getRow(i + 1).getCell(0).toString())) {
				dup = true;
				break;
			}
		}

		return dup;
	}

	public void addFpAttendance(String type, String name, String group) {
		sh_fpattendance_attendance.createRow(num_fpattendance_attendance).createCell(0).setCellValue(name);
		sh_fpattendance_attendance.getRow(num_fpattendance_attendance).createCell(1).setCellValue(group);
		sh_fpattendance_attendance.getRow(num_fpattendance_attendance).createCell(2).setCellValue(type);
		sh_fpattendance_attendance.getRow(num_fpattendance_attendance).createCell(3).setCellValue(type);
		sh_fpattendance_attendance.getRow(num_fpattendance_attendance).createCell(4).setCellValue("");
		for (int i = 5; i <= 56; i++) {
			sh_fpattendance_attendance.getRow(num_fpattendance_attendance).createCell(i).setCellValue(0);
		}

		int idx = Integer.valueOf(type.substring(2, type.length()));
		if (idx < 14 || group.equals("교역자") || group.equals("장로")) {
			sh_fpattendance_attendance.getRow(num_fpattendance_attendance).createCell(57).setCellValue(group);
		} else if (idx == 15) {
			sh_fpattendance_attendance.getRow(num_fpattendance_attendance).createCell(57).setCellValue("거창지교회");
		} else if (idx == 17) {
			sh_fpattendance_attendance.getRow(num_fpattendance_attendance).createCell(57).setCellValue("상주지교회");
		} else if (idx == 19) {
			sh_fpattendance_attendance.getRow(num_fpattendance_attendance).createCell(57).setCellValue("경주지교회");
		} else if (idx == 21) {
			sh_fpattendance_attendance.getRow(num_fpattendance_attendance).createCell(57).setCellValue("성주지교회");
		} else if (idx == 23) {
			sh_fpattendance_attendance.getRow(num_fpattendance_attendance).createCell(57).setCellValue("영주지교회");
		} else if (idx == 25) {
			sh_fpattendance_attendance.getRow(num_fpattendance_attendance).createCell(57).setCellValue("북대구지교회");
		} else if (idx == 27) {
			sh_fpattendance_attendance.getRow(num_fpattendance_attendance).createCell(57).setCellValue("포항지교회");
		} else if (idx == 29) {
			sh_fpattendance_attendance.getRow(num_fpattendance_attendance).createCell(57).setCellValue("대전지교회");
		} else if (idx == 31) {
			sh_fpattendance_attendance.getRow(num_fpattendance_attendance).createCell(57).setCellValue("부여지교회");
		} else if (idx == 33) {
			sh_fpattendance_attendance.getRow(num_fpattendance_attendance).createCell(57).setCellValue("서대구지교회");
		} else if (idx == 35) {
			sh_fpattendance_attendance.getRow(num_fpattendance_attendance).createCell(57).setCellValue("창원지교회");
		} else if (idx == 37) {
			sh_fpattendance_attendance.getRow(num_fpattendance_attendance).createCell(57).setCellValue("북부산지교회");
		}

		num_fpattendance_attendance++;
	}

	public boolean getDuplication(String input) {
		boolean dup = false;
		for (int i = 0; i < num_attendance - 1; i++) {
			try {
				if (input.equals(sh_attendance.getRow(i + 1).getCell(2).toString())) {
					dup = true;
					break;
				}
			} catch (NullPointerException e) {
			}
		}

		return dup;
	}

	public void addAttendance(String type, String name, String group) {
		int nMain = num_attendance + 2;
		int nGeneral = num_attendance_general - 1;
		int nGroup = num_attendance_group - 1;
		int nBc = num_attendance_bc;
		int nBcRead = num_attendance_bc_read;

		int week = week_of_year + 1;
		int idx = 0;
		int nRow = 0;
		int cnt = 0;
		String prevName = "";
		boolean flag = true;

		switch (type.substring(0, 2)) {
		// general
		case "ge":
			idx = Integer.valueOf(type.substring("general".length(), type.length()));
			while (cnt < idx + 1) {
				nRow++;
				try {
					if (!sh_attendance_general.getRow(nRow).getCell(1).toString().equals("")) {
						cnt++;
					}
				} catch (NullPointerException e) {
				}
			}
			prevName = sh_attendance_general.getRow(nRow - 1).getCell(2).toString();

			num_attendance_general++;
			nGeneral++;
			for (int k = nGeneral; k > nRow; k--) {
				sh_attendance_general.createRow(k);
				for (int l = 0; l <= 55; l++) {
					try {
						sh_attendance_general.getRow(k).createCell(l)
								.setCellValue(sh_attendance_general.getRow(k - 1).getCell(l).toString());
					} catch (NullPointerException e) {
					}
				}
			}
			sh_attendance_general.createRow(nRow);
			sh_attendance_general.getRow(nRow).createCell(2).setCellValue(name);
			sh_attendance_general.getRow(nRow).createCell(week).setCellValue("○");
			break;

		// group
		case "gr":
			idx = Integer.valueOf(type.substring("group".length(), type.length()));
			while (cnt < idx) {
				nRow++;
				try {
					if (sh_attendance_group.getRow(nRow).getCell(0).toString().substring(1, 2).equals(" ")) {
						cnt++;
					}
				} catch (Exception e) {
				}
			}
			while (flag) {
				nRow++;
				try {
					if (sh_attendance_group.getRow(nRow).getCell(0).toString().equals(group)) {
						while (flag) {
							nRow++;
							try {
								if (!sh_attendance_group.getRow(nRow).getCell(0).toString().equals("")) {
									flag = false;
								}
							} catch (NullPointerException e) {
							}
						}
					}
				} catch (NullPointerException e) {
				}
			}
			prevName = sh_attendance_group.getRow(nRow - 1).getCell(2).toString();

			num_attendance_group++;
			nGroup++;
			for (int k = nGroup; k > nRow; k--) {
				sh_attendance_group.createRow(k);
				for (int l = 0; l <= 55; l++) {
					try {
						sh_attendance_group.getRow(k).createCell(l)
								.setCellValue(sh_attendance_group.getRow(k - 1).getCell(l).toString());
					} catch (NullPointerException e) {
					}
				}
			}
			sh_attendance_group.createRow(nRow);
			sh_attendance_group.getRow(nRow).createCell(2).setCellValue(name);
			sh_attendance_group.getRow(nRow).createCell(week).setCellValue("○");
			break;

		// bc
		case "bc":
			idx = Integer.valueOf(type.substring("bc".length(), type.length()));
			while (cnt < idx) {
				nRow++;
				try {
					if (sh_attendance_bc_read.getRow(nRow).getCell(0).toString().substring(1, 2).equals(" ")) {
						cnt++;
					}
				} catch (Exception e) {
				}
			}
			while (cnt < idx + 2) {
				nRow++;
				try {
					if (!sh_attendance_bc_read.getRow(nRow).getCell(1).toString().equals("")) {
						cnt++;
					}
				} catch (NullPointerException e) {
				}
			}
			prevName = sh_attendance_bc_read.getRow(nRow - 1).getCell(2).toString();

			num_attendance_bc_read++;
			nBcRead++;
			for (int k = nBcRead; k > nRow; k--) {
				if (k == nBcRead - 1) {
					sh_attendance_bc_read.removeRow(sh_attendance_bc_read.createRow(k));
				} else {
					sh_attendance_bc_read.createRow(k);
					for (int l = 0; l <= 55; l++) {
						try {
							sh_attendance_bc_read.getRow(k).createCell(l)
									.setCellValue(sh_attendance_bc_read.getRow(k - 1).getCell(l).toString());
						} catch (NullPointerException e) {
						}
					}
				}
			}
			sh_attendance_bc_read.createRow(nRow);
			sh_attendance_bc_read.getRow(nRow).createCell(2).setCellValue(name);
			sh_attendance_bc_read.getRow(nRow).createCell(week).setCellValue("○");

			nRow = 0;
			while (flag) {
				nRow++;
				try {
					if (sh_attendance_bc.getRow(nRow).getCell(2).toString().equals(prevName)) {
						nRow++;
						flag = false;
					}
				} catch (NullPointerException e) {
				}
			}

			num_attendance_bc++;
			nBc++;
			for (int k = nBc; k > nRow; k--) {
				if (k == nBc - 1) {
					sh_attendance_bc.removeRow(sh_attendance_bc.createRow(k));
				} else {
					sh_attendance_bc.createRow(k);
					for (int l = 0; l <= 55; l++) {
						try {
							sh_attendance_bc.getRow(k).createCell(l)
									.setCellValue(sh_attendance_bc.getRow(k - 1).getCell(l).toString());
						} catch (NullPointerException e) {
						}
					}
				}
			}
			sh_attendance_bc.createRow(nRow);
			sh_attendance_bc.getRow(nRow).createCell(2).setCellValue(name);
			sh_attendance_bc.getRow(nRow).createCell(week).setCellValue("○");
			break;
		}

		// main
		flag = true;
		nRow = 0;
		while (flag) {
			nRow++;
			try {
				if (sh_attendance.getRow(nRow).getCell(2).toString().equals(prevName)) {
					nRow++;
					flag = false;
				}
			} catch (NullPointerException e) {
			}
		}
		num_attendance++;
		nMain++;
		for (int k = nMain; k > nRow; k--) {
			if (k == nMain - 1 || k == nMain - 3 || k == nMain - 5) {
				sh_attendance.removeRow(sh_attendance.createRow(k));
			} else {
				sh_attendance.createRow(k);
				for (int l = 0; l <= 55; l++) {
					try {
						sh_attendance.getRow(k).createCell(l)
								.setCellValue(sh_attendance.getRow(k - 1).getCell(l).toString());
					} catch (NullPointerException e) {
					}
				}
			}
		}
		sh_attendance.createRow(nRow);
		sh_attendance.getRow(nRow).createCell(2).setCellValue(name);
		sh_attendance.getRow(nRow).createCell(week).setCellValue("○");
	}

	public void saveArchive() {
		FpFormatting fpfmt = new FpFormatting(wb_fpattendance, wb_fpbeliever, wb_fpwordmovement, path);
		int nrow = fpfmt.getNumRow1();
		new FpMonthlyFormatting(wb_fpattendance, path, nrow);
		new FpYearlyFormatting(wb_fpattendance, path, nrow);
		new Formatting(wb_attendance, path);
		new HomeFormatting(wb_attendance_home, path);
	}
}
