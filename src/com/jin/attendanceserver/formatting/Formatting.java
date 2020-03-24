package com.jin.attendanceserver.formatting;

import java.io.FileOutputStream;
import java.io.IOException;
import java.text.SimpleDateFormat;
import java.util.Calendar;
import java.util.Date;

import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.VerticalAlignment;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Formatting {

	final private XSSFWorkbook MAIN, WB;
	XSSFSheet mSH, SH;
	int nMain;
	String title, path;
	int[] first_of_week, cnt, cntGumi, cntDvdd, cntBC;
	String[] weeks;

	XSSFFont fBold, fNormal;
	XSSFCellStyle csBold1, csBold2;
	XSSFCellStyle[][] csNormal;

	public Formatting(XSSFWorkbook main, String directory_path) {
		MAIN = main;
		mSH = MAIN.getSheetAt(0);
		nMain = mSH.getLastRowNum() + 1;

		SimpleDateFormat formatter = new SimpleDateFormat("yyyyMMdd");
		String time = formatter.format(new Date());
		title = time + " 최종 출석부";
		path = directory_path + "\\archive\\Att\\" + title + ".xlsx";

		WB = new XSSFWorkbook();
		SH = WB.createSheet("최종 출석부");
		SH.setPrintGridlines(false);
		SH.setDisplayGridlines(false);

		setFont();
		setCellStyle();
		init();
		performance();

		try {
			WB.write(new FileOutputStream(path));
			WB.close();
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
	}

	public void setFont() {
		fBold = WB.createFont();
		fBold.setBold(true);
		fBold.setColor(IndexedColors.BLACK.getIndex());
		fBold.setFontName("맑은 고딕");
		fBold.setFontHeightInPoints((short) 10);

		fNormal = WB.createFont();
		fNormal.setBold(false);
		fNormal.setColor(IndexedColors.BLACK.getIndex());
		fNormal.setFontName("맑은 고딕");
		fNormal.setFontHeightInPoints((short) 10);
	}

	public void setCellStyle() {
		csBold1 = WB.createCellStyle();
		csBold1.setAlignment(HorizontalAlignment.CENTER);
		csBold1.setVerticalAlignment(VerticalAlignment.CENTER);
		csBold1.setBorderTop(BorderStyle.MEDIUM);
		csBold1.setTopBorderColor(IndexedColors.BLACK.getIndex());
		csBold1.setBorderBottom(BorderStyle.MEDIUM);
		csBold1.setBottomBorderColor(IndexedColors.BLACK.getIndex());
		csBold1.setBorderLeft(BorderStyle.MEDIUM);
		csBold1.setLeftBorderColor(IndexedColors.BLACK.getIndex());
		csBold1.setBorderRight(BorderStyle.MEDIUM);
		csBold1.setRightBorderColor(IndexedColors.BLACK.getIndex());
		csBold1.setFillForegroundColor(IndexedColors.WHITE.getIndex());
		csBold1.setFillPattern(FillPatternType.SOLID_FOREGROUND);
		csBold1.setFont(fBold);
		csBold1.setWrapText(true);

		csBold2 = WB.createCellStyle();
		csBold2.setAlignment(HorizontalAlignment.CENTER);
		csBold2.setVerticalAlignment(VerticalAlignment.CENTER);
		csBold2.setBorderTop(BorderStyle.THIN);
		csBold2.setTopBorderColor(IndexedColors.BLACK.getIndex());
		csBold2.setBorderBottom(BorderStyle.THIN);
		csBold2.setBottomBorderColor(IndexedColors.BLACK.getIndex());
		csBold2.setBorderLeft(BorderStyle.MEDIUM);
		csBold2.setLeftBorderColor(IndexedColors.BLACK.getIndex());
		csBold2.setBorderRight(BorderStyle.MEDIUM);
		csBold2.setRightBorderColor(IndexedColors.BLACK.getIndex());
		csBold2.setFillForegroundColor(IndexedColors.WHITE.getIndex());
		csBold2.setFillPattern(FillPatternType.SOLID_FOREGROUND);
		csBold2.setFont(fBold);
		csBold2.setWrapText(true);

		csNormal = new XSSFCellStyle[4][3];
		// 0: 상좌우, 좌우, 상하좌우 - 이름, 사유
		// 1: 상좌, 상, 상우 - top
		// 2: 좌, x, 우 - center
		// 3: 상하좌, 상하, 상하우 - sum

		for (int i = 0; i < 4; i++) {
			for (int j = 0; j < 3; j++) {
				csNormal[i][j] = WB.createCellStyle();
				csNormal[i][j].setAlignment(HorizontalAlignment.CENTER);
				csNormal[i][j].setVerticalAlignment(VerticalAlignment.CENTER);
				csNormal[i][j].setFillForegroundColor(IndexedColors.WHITE.getIndex());
				csNormal[i][j].setFillPattern(FillPatternType.SOLID_FOREGROUND);
				csNormal[i][j].setFont(fNormal);
				csNormal[i][j].setTopBorderColor(IndexedColors.BLACK.getIndex());
				csNormal[i][j].setBottomBorderColor(IndexedColors.BLACK.getIndex());
				csNormal[i][j].setLeftBorderColor(IndexedColors.BLACK.getIndex());
				csNormal[i][j].setRightBorderColor(IndexedColors.BLACK.getIndex());
			}
		}

		csNormal[0][0].setBorderTop(BorderStyle.MEDIUM);
		csNormal[0][0].setBorderBottom(BorderStyle.THIN);
		csNormal[0][0].setBorderLeft(BorderStyle.MEDIUM);
		csNormal[0][0].setBorderRight(BorderStyle.MEDIUM);
		csNormal[0][1].setBorderTop(BorderStyle.THIN);
		csNormal[0][1].setBorderBottom(BorderStyle.THIN);
		csNormal[0][1].setBorderLeft(BorderStyle.MEDIUM);
		csNormal[0][1].setBorderRight(BorderStyle.MEDIUM);
		csNormal[0][2].setBorderTop(BorderStyle.MEDIUM);
		csNormal[0][2].setBorderBottom(BorderStyle.MEDIUM);
		csNormal[0][2].setBorderLeft(BorderStyle.MEDIUM);
		csNormal[0][2].setBorderRight(BorderStyle.MEDIUM);

		csNormal[1][0].setBorderTop(BorderStyle.MEDIUM);
		csNormal[1][0].setBorderBottom(BorderStyle.THIN);
		csNormal[1][0].setBorderLeft(BorderStyle.MEDIUM);
		csNormal[1][0].setBorderRight(BorderStyle.THIN);
		csNormal[1][1].setBorderTop(BorderStyle.MEDIUM);
		csNormal[1][1].setBorderBottom(BorderStyle.THIN);
		csNormal[1][1].setBorderLeft(BorderStyle.THIN);
		csNormal[1][1].setBorderRight(BorderStyle.THIN);
		csNormal[1][2].setBorderTop(BorderStyle.MEDIUM);
		csNormal[1][2].setBorderBottom(BorderStyle.THIN);
		csNormal[1][2].setBorderLeft(BorderStyle.THIN);
		csNormal[1][2].setBorderRight(BorderStyle.MEDIUM);

		csNormal[2][0].setBorderTop(BorderStyle.THIN);
		csNormal[2][0].setBorderBottom(BorderStyle.THIN);
		csNormal[2][0].setBorderLeft(BorderStyle.MEDIUM);
		csNormal[2][0].setBorderRight(BorderStyle.THIN);
		csNormal[2][1].setBorderTop(BorderStyle.THIN);
		csNormal[2][1].setBorderBottom(BorderStyle.THIN);
		csNormal[2][1].setBorderLeft(BorderStyle.THIN);
		csNormal[2][1].setBorderRight(BorderStyle.THIN);
		csNormal[2][2].setBorderTop(BorderStyle.THIN);
		csNormal[2][2].setBorderBottom(BorderStyle.THIN);
		csNormal[2][2].setBorderLeft(BorderStyle.THIN);
		csNormal[2][2].setBorderRight(BorderStyle.MEDIUM);

		csNormal[3][0].setBorderTop(BorderStyle.MEDIUM);
		csNormal[3][0].setBorderBottom(BorderStyle.MEDIUM);
		csNormal[3][0].setBorderLeft(BorderStyle.MEDIUM);
		csNormal[3][0].setBorderRight(BorderStyle.THIN);
		csNormal[3][1].setBorderTop(BorderStyle.MEDIUM);
		csNormal[3][1].setBorderBottom(BorderStyle.MEDIUM);
		csNormal[3][1].setBorderLeft(BorderStyle.THIN);
		csNormal[3][1].setBorderRight(BorderStyle.THIN);
		csNormal[3][2].setBorderTop(BorderStyle.MEDIUM);
		csNormal[3][2].setBorderBottom(BorderStyle.MEDIUM);
		csNormal[3][2].setBorderLeft(BorderStyle.THIN);
		csNormal[3][2].setBorderRight(BorderStyle.MEDIUM);
	}

	public void init() {
		first_of_week = new int[] { 1, 5, 9, 14, 18, 23, 28, 32, 37, 41, 45, 49, 53 };
		weeks = new String[] { "1/5", "1/12", "1/19", "1/26", "2/2", "2/9", "2/16", "2/23", "3/1", "3/8", "3/15",
				"3/22", "3/29", "4/5", "4/12", "4/19", "4/26", "5/3", "5/10", "5/17", "5/24", "5/31", "6/7", "6/14",
				"6/21", "6/28", "7/5", "7/12", "7/19", "7/26", "8/2", "8/9", "8/16", "8/23", "8/30", "9/6", "9/13",
				"9/20", "9/27", "10/4", "10/11", "10/18", "10/25", "11/1", "11/8", "11/15", "11/22", "11/29", "12/6",
				"12/13", "12/20", "12/27" };
		cnt = new int[53];
		cntGumi = new int[53];
		cntDvdd = new int[53];
		cntBC = new int[53];
		for (int i = 0; i < cnt.length; i++) {
			cnt[i] = 0;
			cntGumi[i] = 0;
			cntDvdd[i] = 0;
			cntBC[i] = 0;
		}
	}

	public void performance() {
		SH.setColumnWidth(0, 1500);
		SH.setColumnWidth(1, 2000);
		SH.setColumnWidth(2, 3000);
		for (int i = 3; i < 55; i++) {
			SH.setColumnWidth(i, 1300);
		}
		SH.setColumnWidth(55, 4000);

		SH.createFreezePane(0, 1, 0, 1);

		int week = Calendar.getInstance().get(Calendar.WEEK_OF_YEAR);
		if (week == 1) {
			if (Calendar.getInstance().get(Calendar.MONTH) == 11) {
				week = 53;
			}
		}
		week += 1;

		for (int i = week + 1; i < 55; i++) {
			SH.setColumnHidden(i, true);
		}
		for (int i = 3; i < first_of_week.length; i++) {
			if (first_of_week[i] + 2 > week) {
				for (int j = 1; j < first_of_week[i - 3]; j++) {
					SH.setColumnHidden(j + 2, true);
				}
				break;
			}
		}

		int idx = 0;
		for (int i = 0; i < nMain; i++) {
			SH.createRow(i);
			for (int j = 0; j <= 55; j++) {
				SH.getRow(i).createCell(j);
				if (i == 0 && j >= 3 && j < 55) {
					SH.getRow(i).getCell(j).setCellValue(weeks[idx]);
					idx++;
				} else {
					try {
						String value = mSH.getRow(i).getCell(j).toString();
						if (value.equals("NULL")) SH.getRow(i).getCell(j).setCellValue("");
						else SH.getRow(i).getCell(j).setCellValue(mSH.getRow(i).getCell(j).toString());
					} catch (NullPointerException e) {
						SH.getRow(i).getCell(j).setCellValue("");
					}
				}
			}
		}

		int nRowPrev = 0;
		int nRow = 0;
		int count = 0;
		boolean flag = false;
		boolean flag_school = false;

		SH.getRow(0).getCell(0).setCellStyle(csBold1);
		SH.getRow(0).getCell(1).setCellStyle(csBold1);
		SH.getRow(0).getCell(2).setCellStyle(csNormal[0][2]);
		for (int n = 3; n < 55; n++) {
			if (first_of_week[count] == n - 2) {
				SH.getRow(0).getCell(n).setCellStyle(csNormal[3][0]);
				count++;
			} else if (first_of_week[count] == n - 1) {
				SH.getRow(0).getCell(n).setCellStyle(csNormal[3][2]);
			} else {
				SH.getRow(0).getCell(n).setCellStyle(csNormal[3][1]);
			}
		}
		SH.getRow(0).getCell(55).setCellStyle(csNormal[0][2]);

		// General, Group
		for (int i = 1; i < nMain; i++) {
			nRow = i;
			SH.getRow(i).getCell(0).setCellStyle(csBold1);
			SH.getRow(i).getCell(1).setCellStyle(csBold1);
			count = 0;

			if (SH.getRow(i).getCell(0).toString().equals("소계")) {
				flag = false;
				if (flag_school) {
					SH.addMergedRegion(new CellRangeAddress(nRowPrev, nRow - 1, 0, 1));
				} else {
					SH.addMergedRegion(new CellRangeAddress(nRowPrev, nRow - 1, 0, 0));
					SH.addMergedRegion(new CellRangeAddress(nRowPrev, nRow - 1, 1, 1));
				}
				nRowPrev = i;
				SH.getRow(i).getCell(2).setCellValue(cnt[0]);
				SH.getRow(i).getCell(2).setCellStyle(csNormal[0][2]);
				cntGumi[0] += cnt[0];
				for (int n = 3; n < 55; n++) {
					SH.getRow(i).getCell(n).setCellValue(cnt[n - 2]);
					cntGumi[n - 2] += cnt[n - 2];
					if (first_of_week[count] == n - 2) {
						SH.getRow(i).getCell(n).setCellStyle(csNormal[3][0]);
						count++;
					} else if (first_of_week[count] == n - 1) {
						SH.getRow(i).getCell(n).setCellStyle(csNormal[3][2]);
					} else {
						SH.getRow(i).getCell(n).setCellStyle(csNormal[3][1]);
					}
				}
				SH.getRow(i).getCell(55).setCellStyle(csNormal[0][2]);
				SH.addMergedRegion(new CellRangeAddress(i, i, 0, 1));
				i++;
				for (int n = 0; n <= 55; n++) {
					SH.getRow(i).getCell(n).setCellStyle(csBold1);
				}
				for (int n = 0; n < cnt.length; n++) {
					cnt[n] = 0;
				}
				SH.addMergedRegion(new CellRangeAddress(i, i, 0, 55));
				if (SH.getRow(i).getCell(0).toString().equals("거 창 지 교 회")) {
					nRow = i + 1;
					nRowPrev = i + 1;
					break;
				} else if (SH.getRow(i).getCell(0).toString().equals("고 등 부")) {
					flag_school = true;
				}

			} else if (!SH.getRow(i).getCell(0).toString().equals("")) {
				if (flag) {
					if (flag_school) {
						SH.addMergedRegion(new CellRangeAddress(nRowPrev, nRow - 1, 0, 1));
					} else {
						SH.addMergedRegion(new CellRangeAddress(nRowPrev, nRow - 1, 0, 0));
						SH.addMergedRegion(new CellRangeAddress(nRowPrev, nRow - 1, 1, 1));
					}
				}
				flag = true;
				nRowPrev = i;
				SH.getRow(i).getCell(2).setCellStyle(csNormal[0][0]);
				cnt[0]++;
				for (int n = 3; n < 55; n++) {
					if (SH.getRow(i).getCell(n).toString().equals("○")) {
						cnt[n - 2]++;
					}
					if (first_of_week[count] == n - 2) {
						SH.getRow(i).getCell(n).setCellStyle(csNormal[1][0]);
						count++;
					} else if (first_of_week[count] == n - 1) {
						SH.getRow(i).getCell(n).setCellStyle(csNormal[1][2]);
					} else {
						SH.getRow(i).getCell(n).setCellStyle(csNormal[1][1]);
					}
				}
				SH.getRow(i).getCell(55).setCellStyle(csNormal[0][0]);

			} else {
				SH.getRow(i).getCell(2).setCellStyle(csNormal[0][1]);
				cnt[0]++;
				for (int n = 3; n < 55; n++) {
					if (SH.getRow(i).getCell(n).toString().equals("○")) {
						cnt[n - 2]++;
					}
					if (first_of_week[count] == n - 2) {
						SH.getRow(i).getCell(n).setCellStyle(csNormal[2][0]);
						count++;
					} else if (first_of_week[count] == n - 1) {
						SH.getRow(i).getCell(n).setCellStyle(csNormal[2][2]);
					} else {
						SH.getRow(i).getCell(n).setCellStyle(csNormal[2][1]);
					}
				}
				SH.getRow(i).getCell(55).setCellStyle(csNormal[0][1]);

			}
		}

		// BC
		int nRowPrevBC = nRowPrev;
		for (int i = nRow; i < nMain - 6; i++) {
			nRow = i;
			SH.getRow(i).getCell(0).setCellStyle(csBold1);
			count = 0;
			flag = true;

			if (SH.getRow(i).getCell(1).toString().equals("소계")) {
				SH.getRow(i).getCell(2).setCellValue(cnt[0]);
				cntDvdd[0] += cnt[0];
				for (int n = 3; n < 55; n++) {
					SH.getRow(i).getCell(n).setCellValue(cnt[n - 2]);
					cntDvdd[n - 2] += cnt[n - 2];
				}
				SH.getRow(i).getCell(1).setCellStyle(csBold2);
				SH.getRow(i).getCell(2).setCellStyle(csNormal[0][1]);
				SH.getRow(i).getCell(55).setCellStyle(csNormal[0][1]);
				if (nRow - nRowPrev != 1) {
					SH.addMergedRegion(new CellRangeAddress(nRowPrev, nRow - 1, 1, 1));
				}
				nRowPrev = i + 1;
				for (int n = 2; n < 55; n++) {
					cnt[n - 2] = 0;
				}

			} else if (SH.getRow(i).getCell(1).toString().equals("총계")) {
				flag = false;
				if (SH.getRow(nRowPrevBC - 1).getCell(0).toString().equals("포 항 지 교 회")) {
					for (int n = 2; n < 55; n++) {
						cntDvdd[n - 2] += cnt[n - 2];
					}
				}
				SH.getRow(i).getCell(2).setCellValue(cntDvdd[0]);
				for (int n = 3; n < 55; n++) {
					SH.getRow(i).getCell(n).setCellValue(cntDvdd[n - 2]);
				}

				SH.getRow(i).getCell(1).setCellStyle(csBold1);
				SH.getRow(i).getCell(2).setCellStyle(csNormal[0][2]);
				SH.getRow(i).getCell(55).setCellStyle(csNormal[0][2]);

				for (int n = 3; n < 55; n++) {
					if (first_of_week[count] == n - 2) {
						SH.getRow(i).getCell(n).setCellStyle(csNormal[3][0]);
						count++;
					} else if (first_of_week[count] == n - 1) {
						SH.getRow(i).getCell(n).setCellStyle(csNormal[3][2]);
					} else {
						SH.getRow(i).getCell(n).setCellStyle(csNormal[3][1]);
					}
				}

				if (nRow - nRowPrevBC != 1) {
					if (SH.getRow(nRowPrevBC - 1).getCell(0).toString().equals("포 항 지 교 회")) {
						if (nRow - nRowPrev != 1) {
							SH.addMergedRegion(new CellRangeAddress(nRowPrev, nRow - 1, 1, 1));
						}
					}
					SH.addMergedRegion(new CellRangeAddress(nRowPrevBC, nRow, 0, 0));
				}
				nRowPrev = i + 2;
				nRowPrevBC = i + 2;
				if (i != nMain - 7) {
					i++;
					SH.getRow(i).getCell(1).setCellStyle(csBold1);
					SH.getRow(i).getCell(2).setCellStyle(csNormal[0][2]);
					SH.getRow(i).getCell(55).setCellStyle(csNormal[0][2]);
					for (int n = 0; n <= 55; n++) {
						SH.getRow(i).getCell(n).setCellStyle(csBold1);
					}
					SH.addMergedRegion(new CellRangeAddress(i, i, 0, 55));
				}
				for (int n = 2; n < 55; n++) {
					cntBC[n - 2] += cntDvdd[n - 2];
					cntDvdd[n - 2] = 0;
					cnt[n - 2] = 0;
				}

			} else {
				SH.getRow(i).getCell(1).setCellStyle(csBold2);
				SH.getRow(i).getCell(2).setCellStyle(csNormal[0][1]);
				SH.getRow(i).getCell(55).setCellStyle(csNormal[0][1]);
				cnt[0]++;
				for (int n = 3; n < 55; n++) {
					if (SH.getRow(i).getCell(n).toString().equals("○")) {
						cnt[n - 2]++;
					}
				}

			}

			if (flag) {
				for (int n = 3; n < 55; n++) {
					if (first_of_week[count] == n - 2) {
						SH.getRow(i).getCell(n).setCellStyle(csNormal[2][0]);
						count++;
					} else if (first_of_week[count] == n - 1) {
						SH.getRow(i).getCell(n).setCellStyle(csNormal[2][2]);
					} else {
						SH.getRow(i).getCell(n).setCellStyle(csNormal[2][1]);
					}
				}
			}
		}

		SH.addMergedRegion(new CellRangeAddress(nMain - 5, nMain - 5, 0, 1));
		SH.addMergedRegion(new CellRangeAddress(nMain - 3, nMain - 3, 0, 1));
		SH.addMergedRegion(new CellRangeAddress(nMain - 1, nMain - 1, 0, 1));
		for (int i = 1; i <= 5; i += 2) {
			SH.getRow(nMain - i).getCell(0).setCellStyle(csBold1);
			SH.getRow(nMain - i).getCell(1).setCellStyle(csBold1);
			SH.getRow(nMain - i).getCell(2).setCellStyle(csNormal[0][2]);
			SH.getRow(nMain - i).getCell(55).setCellStyle(csNormal[0][2]);
		}

		count = 0;
		SH.getRow(nMain - 5).getCell(2).setCellValue(cntBC[0]);
		for (int n = 3; n < 55; n++) {
			SH.getRow(nMain - 5).getCell(n).setCellValue(cntBC[n - 2]);
			if (first_of_week[count] == n - 2) {
				SH.getRow(nMain - 5).getCell(n).setCellStyle(csNormal[3][0]);
				count++;
			} else if (first_of_week[count] == n - 1) {
				SH.getRow(nMain - 5).getCell(n).setCellStyle(csNormal[3][2]);
			} else {
				SH.getRow(nMain - 5).getCell(n).setCellStyle(csNormal[3][1]);
			}
		}

		count = 0;
		SH.getRow(nMain - 3).getCell(2).setCellValue(cntGumi[0]);
		for (int n = 3; n < 55; n++) {
			SH.getRow(nMain - 3).getCell(n).setCellValue(cntGumi[n - 2]);
			if (first_of_week[count] == n - 2) {
				SH.getRow(nMain - 3).getCell(n).setCellStyle(csNormal[3][0]);
				count++;
			} else if (first_of_week[count] == n - 1) {
				SH.getRow(nMain - 3).getCell(n).setCellStyle(csNormal[3][2]);
			} else {
				SH.getRow(nMain - 3).getCell(n).setCellStyle(csNormal[3][1]);
			}
		}

		count = 0;
		SH.getRow(nMain - 1).getCell(2).setCellValue(cntGumi[0] + cntBC[0]);
		for (int n = 3; n < 55; n++) {
			SH.getRow(nMain - 1).getCell(n).setCellValue(cntGumi[n - 2] + cntBC[n - 2]);
			if (first_of_week[count] == n - 2) {
				SH.getRow(nMain - 1).getCell(n).setCellStyle(csNormal[3][0]);
				count++;
			} else if (first_of_week[count] == n - 1) {
				SH.getRow(nMain - 1).getCell(n).setCellStyle(csNormal[3][2]);
			} else {
				SH.getRow(nMain - 1).getCell(n).setCellStyle(csNormal[3][1]);
			}
		}

	}
}
