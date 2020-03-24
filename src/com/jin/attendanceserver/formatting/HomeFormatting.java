package com.jin.attendanceserver.formatting;

import java.io.FileOutputStream;
import java.io.IOException;
import java.text.SimpleDateFormat;
import java.util.Calendar;

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

public class HomeFormatting {

	final private XSSFWorkbook MAIN, WB;
	XSSFSheet mSH, SH;
	int nMain;
	String title, path;
	int[] first_of_week, cnt, cntGumi, cntDvdd, cntBC;

	XSSFFont fTitle, fBold, fNormal;
	XSSFCellStyle csTitle, csBold1, csBold2;
	XSSFCellStyle[][] csNormal;

	public HomeFormatting(XSSFWorkbook main, String directory_path) {
		MAIN = main;
		mSH = MAIN.getSheetAt(0);
		nMain = mSH.getLastRowNum() + 1;

		SimpleDateFormat formatter = new SimpleDateFormat("M월 F주차");
		Calendar c = Calendar.getInstance();
		c.set(Calendar.DAY_OF_WEEK, Calendar.SUNDAY);
		String time = formatter.format(c.getTime());
		title = time + " 출석부";
		path = directory_path + "\\archive\\HomeAtt\\" + title + ".xlsx";

		WB = new XSSFWorkbook();
		SH = WB.createSheet("출석부");
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
		fTitle = WB.createFont();
		fTitle.setBold(true);
		fTitle.setColor(IndexedColors.BLACK.getIndex());
		fTitle.setFontName("맑은 고딕");
		fTitle.setFontHeightInPoints((short) 20);

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
		csTitle = WB.createCellStyle();
		csTitle.setAlignment(HorizontalAlignment.CENTER);
		csTitle.setVerticalAlignment(VerticalAlignment.TOP);
		csTitle.setFont(fTitle);

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
		SH.addMergedRegion(new CellRangeAddress(0, 2, 0, 12));
		SH.createRow(0).createCell(0).setCellValue(title);
		SH.getRow(0).getCell(0).setCellStyle(csTitle);

		SH.setColumnWidth(0, 1500);
		SH.setColumnWidth(1, 2000);
		SH.setColumnWidth(2, 3000);
		for (int n = 3; n <= 12; n++) {
			if (n % 2 == 1) {
				SH.setColumnWidth(n, 1300);
				SH.addMergedRegion(new CellRangeAddress(3, 3, n, n + 1));
			} else {
				SH.setColumnWidth(n, 4000);
			}
		}

		SH.setColumnBreak(12);
		SH.createFreezePane(0, 4);

		for (int i = 0; i < nMain; i++) {
			SH.createRow(i + 3);
			for (int j = 0; j <= 12; j++) {
				SH.getRow(i + 3).createCell(j);
				try {
					String value = mSH.getRow(i).getCell(j).toString();
					if (value.equals("NULL"))
						SH.getRow(i + 3).getCell(j).setCellValue("");
					else
						SH.getRow(i + 3).getCell(j).setCellValue(mSH.getRow(i).getCell(j).toString());
				} catch (NullPointerException e) {
					SH.getRow(i + 3).getCell(j).setCellValue("");
				}
			}
		}

		for (int n = 0; n <= 12; n++) {
			SH.getRow(3).getCell(n).setCellStyle(csBold1);
		}

		int nRowPrev = 0;
		int nRow = 0;
		boolean flag = false;
		boolean flag_school = false;

		// General, Group
		for (int i = 4; i < nMain + 3; i++) {
			nRow = i;
			SH.getRow(i).getCell(0).setCellStyle(csBold1);
			SH.getRow(i).getCell(1).setCellStyle(csBold1);
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
				for (int n = 3; n <= 12; n++) {
					if (n % 2 == 1) {
						SH.addMergedRegion(new CellRangeAddress(i, i, n, n + 1));
						SH.getRow(i).getCell(n).setCellValue(
								cnt[n - 2] + " (" + Math.round((float) cnt[n - 2] / (float) cnt[0] * 100f) + "%)");
						cntGumi[n - 2] += cnt[n - 2];
						SH.getRow(i).getCell(n).setCellStyle(csNormal[3][0]);
					} else {
						SH.getRow(i).getCell(n).setCellStyle(csNormal[3][2]);
					}
				}
				SH.addMergedRegion(new CellRangeAddress(i, i, 0, 1));
				i++;
				for (int n = 0; n <= 12; n++) {
					SH.getRow(i).getCell(n).setCellStyle(csBold1);
				}
				for (int n = 0; n < cnt.length; n++) {
					cnt[n] = 0;
				}
				SH.addMergedRegion(new CellRangeAddress(i, i, 0, 12));
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
				for (int n = 3; n <= 12; n++) {
					if (SH.getRow(i).getCell(n).toString().equals("○")) {
						cnt[n - 2]++;
					}
					if (n % 2 == 1) {
						SH.getRow(i).getCell(n).setCellStyle(csNormal[1][0]);
					} else {
						SH.getRow(i).getCell(n).setCellStyle(csNormal[1][2]);
					}
				}

			} else {
				SH.getRow(i).getCell(2).setCellStyle(csNormal[0][1]);
				cnt[0]++;
				for (int n = 3; n <= 12; n++) {
					if (SH.getRow(i).getCell(n).toString().equals("○")) {
						cnt[n - 2]++;
					}
					if (n % 2 == 1) {
						SH.getRow(i).getCell(n).setCellStyle(csNormal[2][0]);
					} else {
						SH.getRow(i).getCell(n).setCellStyle(csNormal[2][2]);
					}
				}
			}
		}

		// BC
		int nRowPrevBC = nRowPrev;
		for (int i = nRow; i < nMain - 3; i++) {
			nRow = i;
			SH.getRow(i).getCell(0).setCellStyle(csBold1);
			flag = true;

			if (SH.getRow(i).getCell(1).toString().equals("소계")) {
				SH.getRow(i).getCell(2).setCellValue(cnt[0]);
				cntDvdd[0] += cnt[0];
				for (int n = 3; n <= 12; n++) {
					if (n % 2 == 1) {
						SH.addMergedRegion(new CellRangeAddress(i, i, n, n + 1));
						SH.getRow(i).getCell(n).setCellValue(
								cnt[n - 2] + " (" + Math.round((float) cnt[n - 2] / (float) cnt[0] * 100f) + "%)");
						cntDvdd[n - 2] += cnt[n - 2];
					}
				}
				SH.getRow(i).getCell(1).setCellStyle(csBold2);
				SH.getRow(i).getCell(2).setCellStyle(csNormal[0][1]);
				if (nRow - nRowPrev != 1) {
					SH.addMergedRegion(new CellRangeAddress(nRowPrev, nRow - 1, 1, 1));
				}
				nRowPrev = i + 1;
				for (int n = 2; n <= 12; n++) {
					cnt[n - 2] = 0;
				}

			} else if (SH.getRow(i).getCell(1).toString().equals("총계")) {
				flag = false;
				if (SH.getRow(nRowPrevBC - 1).getCell(0).toString().equals("포 항 지 교 회")) {
					for (int n = 2; n <= 12; n++) {
						cntDvdd[n - 2] += cnt[n - 2];
					}
				}
				SH.getRow(i).getCell(2).setCellValue(cntDvdd[0]);
				for (int n = 3; n <= 12; n++) {
					if (n % 2 == 1) {
						SH.addMergedRegion(new CellRangeAddress(i, i, n, n + 1));
						SH.getRow(i).getCell(n).setCellValue(cntDvdd[n - 2] + " ("
								+ Math.round((float) cntDvdd[n - 2] / (float) cntDvdd[0] * 100f) + "%)");
					}
				}

				SH.getRow(i).getCell(1).setCellStyle(csBold1);
				SH.getRow(i).getCell(2).setCellStyle(csNormal[0][2]);

				for (int n = 3; n <= 12; n++) {
					if (n % 2 == 1) {
						SH.getRow(i).getCell(n).setCellStyle(csNormal[3][0]);
					} else {
						SH.getRow(i).getCell(n).setCellStyle(csNormal[3][2]);
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
				if (i != nMain - 4) {
					i++;
					SH.getRow(i).getCell(1).setCellStyle(csBold1);
					SH.getRow(i).getCell(2).setCellStyle(csNormal[0][2]);
					for (int n = 0; n <= 12; n++) {
						SH.getRow(i).getCell(n).setCellStyle(csBold1);
					}
					SH.addMergedRegion(new CellRangeAddress(i, i, 0, 12));
				}
				for (int n = 2; n <= 12; n++) {
					cntBC[n - 2] += cntDvdd[n - 2];
					cntDvdd[n - 2] = 0;
					cnt[n - 2] = 0;
				}

			} else {
				SH.getRow(i).getCell(1).setCellStyle(csBold2);
				SH.getRow(i).getCell(2).setCellStyle(csNormal[0][1]);
				cnt[0]++;
				for (int n = 3; n <= 12; n++) {
					if (SH.getRow(i).getCell(n).toString().equals("○")) {
						cnt[n - 2]++;
					}
				}

			}

			if (flag) {
				for (int n = 3; n <= 12; n++) {
					if (n % 2 == 1) {
						SH.getRow(i).getCell(n).setCellStyle(csNormal[2][0]);
					} else {
						SH.getRow(i).getCell(n).setCellStyle(csNormal[2][2]);
					}
				}
			}
		}

		SH.addMergedRegion(new CellRangeAddress(nMain - 2, nMain - 2, 0, 1));
		SH.addMergedRegion(new CellRangeAddress(nMain, nMain, 0, 1));
		SH.addMergedRegion(new CellRangeAddress(nMain + 2, nMain + 2, 0, 1));
		for (int i = -2; i <= 2; i += 2) {
			SH.getRow(nMain - i).getCell(0).setCellStyle(csBold1);
			SH.getRow(nMain - i).getCell(1).setCellStyle(csBold1);
			SH.getRow(nMain - i).getCell(2).setCellStyle(csNormal[0][2]);
		}

		SH.getRow(nMain - 2).getCell(2).setCellValue(cntBC[0]);
		for (int n = 3; n <= 12; n++) {
			if (n % 2 == 1) {
				SH.addMergedRegion(new CellRangeAddress(nMain - 2, nMain - 2, n, n + 1));
				SH.getRow(nMain - 2).getCell(n).setCellValue(
						cntBC[n - 2] + " (" + Math.round(((float) cntBC[n - 2] / (float) cntBC[0] * 100f)) + "%)");
				SH.getRow(nMain - 2).getCell(n).setCellStyle(csNormal[3][0]);
			} else {
				SH.getRow(nMain - 2).getCell(n).setCellStyle(csNormal[3][2]);
			}
		}

		SH.getRow(nMain).getCell(2).setCellValue(cntGumi[0]);
		for (int n = 3; n <= 12; n++) {
			if (n % 2 == 1) {
				SH.addMergedRegion(new CellRangeAddress(nMain, nMain, n, n + 1));
				SH.getRow(nMain).getCell(n).setCellValue(cntGumi[n - 2] + " ("
						+ Math.round(((float) cntGumi[n - 2] / (float) cntGumi[0] * 100f)) + "%)");
				SH.getRow(nMain).getCell(n).setCellStyle(csNormal[3][0]);
			} else {
				SH.getRow(nMain).getCell(n).setCellStyle(csNormal[3][2]);
			}
		}

		SH.getRow(nMain + 2).getCell(2).setCellValue(cntGumi[0] + cntBC[0]);
		for (int n = 3; n <= 12; n++) {
			if (n % 2 == 1) {
				SH.addMergedRegion(new CellRangeAddress(nMain + 2, nMain + 2, n, n + 1));
				SH.getRow(nMain + 2).getCell(n)
						.setCellValue(cntGumi[n - 2] + cntBC[n - 2] + " (" + Math.round(
								((float) (cntGumi[n - 2] + cntBC[n - 2]) / (float) (cntGumi[0] + cntBC[0]) * 100f))
								+ "%)");
				SH.getRow(nMain + 2).getCell(n).setCellStyle(csNormal[3][0]);
			} else {
				SH.getRow(nMain + 2).getCell(n).setCellStyle(csNormal[3][2]);
			}
		}
	}
}