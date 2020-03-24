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

public class FpYearlyFormatting {
	final private XSSFWorkbook ATTENDANCE, WB;
	XSSFSheet attSH, SH;
	int nAtt, nRow;
	String title, path;
	String[] category1 = { "교역자", "장로", "안수집사", "권사", "서리집사(남)", "서리집사(여)", "권찰", "성도(남)", "성도(여)", "청년", "대학", "거창지교회",
			"상주지교회", "경주지교회", "성주지교회", "영주지교회", "북대구지교회", "포항지교회", "대전지교회", "부여지교회", "서대구지교회", "창원지교회", "북부산지교회" };
	String[][] column;

	XSSFFont fTitle, fNormal;
	XSSFCellStyle csTitle, csCategoryLeft, csCategoryCenter, csCategoryRight, csNormalLeft, csNormalCenter,
			csNormalRight, csNormalLastLeft, csNormalLastCenter, csNormalLastRight;

	short cCategory = IndexedColors.PALE_BLUE.getIndex();
	short cNormal = IndexedColors.WHITE.getIndex();

	public FpYearlyFormatting(XSSFWorkbook attendance, String directory_path, int nrow) {
		ATTENDANCE = attendance;
		attSH = ATTENDANCE.getSheetAt(0);
		nAtt = attSH.getPhysicalNumberOfRows();
		nRow = nrow;

		SimpleDateFormat formatter = new SimpleDateFormat("yyyy년");
		String time = formatter.format(new Date());
		title = time + " 전교인 주일 현장전도";
		path = directory_path + "\\archive\\FpYearlyAtt\\" + title + ".xlsx";

		WB = new XSSFWorkbook();
		SH = WB.createSheet("출석");
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

		csCategoryLeft = WB.createCellStyle();
		csCategoryLeft.setAlignment(HorizontalAlignment.CENTER);
		csCategoryLeft.setVerticalAlignment(VerticalAlignment.CENTER);
		csCategoryLeft.setBorderTop(BorderStyle.THICK);
		csCategoryLeft.setTopBorderColor(IndexedColors.BLACK.getIndex());
		csCategoryLeft.setBorderBottom(BorderStyle.THIN);
		csCategoryLeft.setBottomBorderColor(IndexedColors.BLACK.getIndex());
		csCategoryLeft.setBorderLeft(BorderStyle.THICK);
		csCategoryLeft.setLeftBorderColor(IndexedColors.BLACK.getIndex());
		csCategoryLeft.setBorderRight(BorderStyle.THIN);
		csCategoryLeft.setRightBorderColor(IndexedColors.BLACK.getIndex());
		csCategoryLeft.setFillForegroundColor(cCategory);
		csCategoryLeft.setFillPattern(FillPatternType.SOLID_FOREGROUND);
		csCategoryLeft.setFont(fNormal);

		csCategoryCenter = WB.createCellStyle();
		csCategoryCenter.setAlignment(HorizontalAlignment.CENTER);
		csCategoryCenter.setVerticalAlignment(VerticalAlignment.CENTER);
		csCategoryCenter.setBorderTop(BorderStyle.THICK);
		csCategoryCenter.setTopBorderColor(IndexedColors.BLACK.getIndex());
		csCategoryCenter.setBorderBottom(BorderStyle.THIN);
		csCategoryCenter.setBottomBorderColor(IndexedColors.BLACK.getIndex());
		csCategoryCenter.setBorderLeft(BorderStyle.THIN);
		csCategoryCenter.setLeftBorderColor(IndexedColors.BLACK.getIndex());
		csCategoryCenter.setFillForegroundColor(cCategory);
		csCategoryCenter.setFillPattern(FillPatternType.SOLID_FOREGROUND);
		csCategoryCenter.setFont(fNormal);

		csCategoryRight = WB.createCellStyle();
		csCategoryRight.setAlignment(HorizontalAlignment.CENTER);
		csCategoryRight.setVerticalAlignment(VerticalAlignment.CENTER);
		csCategoryRight.setBorderTop(BorderStyle.THICK);
		csCategoryRight.setTopBorderColor(IndexedColors.BLACK.getIndex());
		csCategoryRight.setBorderBottom(BorderStyle.THIN);
		csCategoryRight.setBottomBorderColor(IndexedColors.BLACK.getIndex());
		csCategoryRight.setBorderRight(BorderStyle.THICK);
		csCategoryRight.setRightBorderColor(IndexedColors.BLACK.getIndex());
		csCategoryRight.setFillForegroundColor(cCategory);
		csCategoryRight.setFillPattern(FillPatternType.SOLID_FOREGROUND);
		csCategoryRight.setFont(fNormal);

		csNormalLeft = WB.createCellStyle();
		csNormalLeft.setAlignment(HorizontalAlignment.CENTER);
		csNormalLeft.setVerticalAlignment(VerticalAlignment.CENTER);
		csNormalLeft.setBorderTop(BorderStyle.THIN);
		csNormalLeft.setTopBorderColor(IndexedColors.BLACK.getIndex());
		csNormalLeft.setBorderBottom(BorderStyle.THIN);
		csNormalLeft.setBottomBorderColor(IndexedColors.BLACK.getIndex());
		csNormalLeft.setBorderLeft(BorderStyle.THICK);
		csNormalLeft.setLeftBorderColor(IndexedColors.BLACK.getIndex());
		csNormalLeft.setBorderRight(BorderStyle.THIN);
		csNormalLeft.setRightBorderColor(IndexedColors.BLACK.getIndex());
		csNormalLeft.setFillForegroundColor(cNormal);
		csNormalLeft.setFillPattern(FillPatternType.SOLID_FOREGROUND);
		csNormalLeft.setFont(fNormal);

		csNormalCenter = WB.createCellStyle();
		csNormalCenter.setAlignment(HorizontalAlignment.CENTER);
		csNormalCenter.setVerticalAlignment(VerticalAlignment.CENTER);
		csNormalCenter.setBorderTop(BorderStyle.THIN);
		csNormalCenter.setTopBorderColor(IndexedColors.BLACK.getIndex());
		csNormalCenter.setBorderBottom(BorderStyle.THIN);
		csNormalCenter.setBottomBorderColor(IndexedColors.BLACK.getIndex());
		csNormalCenter.setBorderLeft(BorderStyle.THIN);
		csNormalCenter.setLeftBorderColor(IndexedColors.BLACK.getIndex());
		csNormalCenter.setBorderRight(BorderStyle.THIN);
		csNormalCenter.setRightBorderColor(IndexedColors.BLACK.getIndex());
		csNormalCenter.setFillForegroundColor(cNormal);
		csNormalCenter.setFillPattern(FillPatternType.SOLID_FOREGROUND);
		csNormalCenter.setFont(fNormal);

		csNormalRight = WB.createCellStyle();
		csNormalRight.setAlignment(HorizontalAlignment.CENTER);
		csNormalRight.setVerticalAlignment(VerticalAlignment.CENTER);
		csNormalRight.setBorderTop(BorderStyle.THIN);
		csNormalRight.setTopBorderColor(IndexedColors.BLACK.getIndex());
		csNormalRight.setBorderBottom(BorderStyle.THIN);
		csNormalRight.setBottomBorderColor(IndexedColors.BLACK.getIndex());
		csNormalRight.setBorderLeft(BorderStyle.THIN);
		csNormalRight.setLeftBorderColor(IndexedColors.BLACK.getIndex());
		csNormalRight.setBorderRight(BorderStyle.THICK);
		csNormalRight.setRightBorderColor(IndexedColors.BLACK.getIndex());
		csNormalRight.setFillForegroundColor(cNormal);
		csNormalRight.setFillPattern(FillPatternType.SOLID_FOREGROUND);
		csNormalRight.setFont(fNormal);

		csNormalLastLeft = WB.createCellStyle();
		csNormalLastLeft.setAlignment(HorizontalAlignment.CENTER);
		csNormalLastLeft.setVerticalAlignment(VerticalAlignment.CENTER);
		csNormalLastLeft.setBorderTop(BorderStyle.THIN);
		csNormalLastLeft.setTopBorderColor(IndexedColors.BLACK.getIndex());
		csNormalLastLeft.setBorderBottom(BorderStyle.THICK);
		csNormalLastLeft.setBottomBorderColor(IndexedColors.BLACK.getIndex());
		csNormalLastLeft.setBorderLeft(BorderStyle.THICK);
		csNormalLastLeft.setLeftBorderColor(IndexedColors.BLACK.getIndex());
		csNormalLastLeft.setBorderRight(BorderStyle.THIN);
		csNormalLastLeft.setRightBorderColor(IndexedColors.BLACK.getIndex());
		csNormalLastLeft.setFillForegroundColor(cNormal);
		csNormalLastLeft.setFillPattern(FillPatternType.SOLID_FOREGROUND);
		csNormalLastLeft.setFont(fNormal);

		csNormalLastCenter = WB.createCellStyle();
		csNormalLastCenter.setAlignment(HorizontalAlignment.CENTER);
		csNormalLastCenter.setVerticalAlignment(VerticalAlignment.CENTER);
		csNormalLastCenter.setBorderTop(BorderStyle.THIN);
		csNormalLastCenter.setTopBorderColor(IndexedColors.BLACK.getIndex());
		csNormalLastCenter.setBorderBottom(BorderStyle.THICK);
		csNormalLastCenter.setBottomBorderColor(IndexedColors.BLACK.getIndex());
		csNormalLastCenter.setBorderLeft(BorderStyle.THIN);
		csNormalLastCenter.setLeftBorderColor(IndexedColors.BLACK.getIndex());
		csNormalLastCenter.setBorderRight(BorderStyle.THIN);
		csNormalLastCenter.setRightBorderColor(IndexedColors.BLACK.getIndex());
		csNormalLastCenter.setFillForegroundColor(cNormal);
		csNormalLastCenter.setFillPattern(FillPatternType.SOLID_FOREGROUND);
		csNormalLastCenter.setFont(fNormal);

		csNormalLastRight = WB.createCellStyle();
		csNormalLastRight.setAlignment(HorizontalAlignment.CENTER);
		csNormalLastRight.setVerticalAlignment(VerticalAlignment.CENTER);
		csNormalLastRight.setBorderTop(BorderStyle.THIN);
		csNormalLastRight.setTopBorderColor(IndexedColors.BLACK.getIndex());
		csNormalLastRight.setBorderBottom(BorderStyle.THICK);
		csNormalLastRight.setBottomBorderColor(IndexedColors.BLACK.getIndex());
		csNormalLastRight.setBorderLeft(BorderStyle.THIN);
		csNormalLastRight.setLeftBorderColor(IndexedColors.BLACK.getIndex());
		csNormalLastRight.setBorderRight(BorderStyle.THICK);
		csNormalLastRight.setRightBorderColor(IndexedColors.BLACK.getIndex());
		csNormalLastRight.setFillForegroundColor(cNormal);
		csNormalLastRight.setFillPattern(FillPatternType.SOLID_FOREGROUND);
		csNormalLastRight.setFont(fNormal);
	}

	public void init() {
		String[][][] people = new String[category1.length][2][];
		int week = Calendar.getInstance().get(Calendar.WEEK_OF_YEAR);
		if (week == 1) {
			if (Calendar.getInstance().get(Calendar.MONTH) == 11) {
				week = 53;
			}
		}
		week += 3;
		int[] N = new int[category1.length];
		for (int i = 0; i < N.length; i++) {
			N[i] = 0;
		}
		for (int i = 1; i < nAtt; i++) {
			String t1 = attSH.getRow(i).getCell(57).toString();
			for (int j = 0; j < category1.length; j++) {
				if (category1[j].equals(t1)) {
					N[j]++;
					break;
				}
			}
		}
		for (int i = 0; i < people.length; i++) {
			people[i][0] = new String[N[i]];
			people[i][1] = new String[N[i]];
			N[i] = 0;
		}
		for (int i = 1; i < nAtt; i++) {
			String t1 = attSH.getRow(i).getCell(57).toString();
			String t2 = attSH.getRow(i).getCell(0).toString();
			for (int j = 0; j < category1.length; j++) {
				if (category1[j].equals(t1)) {
					people[j][0][N[j]] = t2;
					int cntYear1 = 0;
					int cntYear2 = 0;
					for (int k = 5; k <= week; k++) {
						if (attSH.getRow(i).getCell(k).toString().equals("2.0")) {
							cntYear1++;
						} else if (attSH.getRow(i).getCell(k).toString().equals("1.0")) {
							cntYear2++;
						}
					}
					people[j][1][N[j]] = String.valueOf(cntYear1) + " / " + String.valueOf(cntYear2);
					N[j]++;
					break;
				}
			}
		}

		column = new String[nRow * 3 - 3][16];
		int idx1 = 0;
		int idx2 = 0;
		int idxCategory = 0;
		int n1 = 0;
		int n2 = nRow - 3;
		int n3 = nRow * 2 - 3;
		int n4 = nRow * 3 - 3;

		for (int i = 0; i < people.length; i++) {
			int idxNum = 1;

			if (idx1 == n2 - 1 || idx1 == n3 - 1 || idx1 == n4 - 1) {
				column[idx1][idx2 * 4] = "0.0";
				column[idx1][idx2 * 4 + 1] = "";
				column[idx1][idx2 * 4 + 2] = "";
				column[idx1][idx2 * 4 + 3] = "";
				idx1++;
				if (idx1 == n2) {
					idx2++;
					if (idx2 == 4) {
						idx2 = 0;
					} else {
						idx1 = n1;
					}
				} else if (idx1 == n3) {
					idx2++;
					if (idx2 == 4) {
						idx2 = 0;
					} else {
						idx1 = n2;
					}
				} else if (idx1 == n4) {
					idx2++;
					if (idx2 == 4) {
						idx2 = 0;
					} else {
						idx1 = n3;
					}
				}
			} else if (idx1 == n2 - 2 || idx1 == n3 - 2 || idx1 == n4 - 2) {
				column[idx1][idx2 * 4] = "0.0";
				column[idx1][idx2 * 4 + 1] = "";
				column[idx1][idx2 * 4 + 2] = "";
				column[idx1][idx2 * 4 + 3] = "";
				idx1++;
				column[idx1][idx2 * 4] = "0.0";
				column[idx1][idx2 * 4 + 1] = "";
				column[idx1][idx2 * 4 + 2] = "";
				column[idx1][idx2 * 4 + 3] = "";
				idx1++;
				if (idx1 == n2) {
					idx2++;
					if (idx2 == 4) {
						idx2 = 0;
					} else {
						idx1 = n1;
					}
				} else if (idx1 == n3) {
					idx2++;
					if (idx2 == 4) {
						idx2 = 0;
					} else {
						idx1 = n2;
					}
				} else if (idx1 == n4) {
					idx2++;
					if (idx2 == 4) {
						idx2 = 0;
					} else {
						idx1 = n3;
					}
				}
			} else if (idx1 == n2 - 3 || idx1 == n3 - 3 || idx1 == n4 - 3) {
				column[idx1][idx2 * 4] = "0.0";
				column[idx1][idx2 * 4 + 1] = "";
				column[idx1][idx2 * 4 + 2] = "";
				column[idx1][idx2 * 4 + 3] = "";
				idx1++;
				column[idx1][idx2 * 4] = "0.0";
				column[idx1][idx2 * 4 + 1] = "";
				column[idx1][idx2 * 4 + 2] = "";
				column[idx1][idx2 * 4 + 3] = "";
				idx1++;
				column[idx1][idx2 * 4] = "0.0";
				column[idx1][idx2 * 4 + 1] = "";
				column[idx1][idx2 * 4 + 2] = "";
				column[idx1][idx2 * 4 + 3] = "";
				idx1++;
				if (idx1 == n2) {
					idx2++;
					if (idx2 == 4) {
						idx2 = 0;
					} else {
						idx1 = n1;
					}
				} else if (idx1 == n3) {
					idx2++;
					if (idx2 == 4) {
						idx2 = 0;
					} else {
						idx1 = n2;
					}
				} else if (idx1 == n4) {
					idx2++;
					if (idx2 == 4) {
						idx2 = 0;
					} else {
						idx1 = n3;
					}
				}
			}

			column[idx1][idx2 * 4] = "3.0";
			column[idx1][idx2 * 4 + 1] = "NO";
			column[idx1][idx2 * 4 + 2] = category1[idxCategory];
			column[idx1][idx2 * 4 + 3] = "";
			idx1++;

			for (int j = 0; j < people[i][0].length; j++) {
				column[idx1][idx2 * 4] = "0.0";
				column[idx1][idx2 * 4 + 1] = String.valueOf(idxNum);
				column[idx1][idx2 * 4 + 2] = people[i][0][j];
				column[idx1][idx2 * 4 + 3] = people[i][1][j];
				idxNum++;

				idx1++;
				if (idx1 == n2) {
					idx2++;
					if (idx2 == 4) {
						idx2 = 0;
					} else {
						idx1 = n1;
					}
					if (j != people[i][0].length - 1) {
						column[idx1][idx2 * 4] = "3.0";
						column[idx1][idx2 * 4 + 1] = "NO";
						column[idx1][idx2 * 4 + 2] = category1[idxCategory];
						column[idx1][idx2 * 4 + 3] = "";
						idx1++;
					}
				} else if (idx1 == n3) {
					idx2++;
					if (idx2 == 4) {
						idx2 = 0;
					} else {
						idx1 = n2;
					}
					if (j != people[i][0].length - 1) {
						column[idx1][idx2 * 4] = "3.0";
						column[idx1][idx2 * 4 + 1] = "NO";
						column[idx1][idx2 * 4 + 2] = category1[idxCategory];
						column[idx1][idx2 * 4 + 3] = "";
						idx1++;
					}
				} else if (idx1 == n4) {
					idx2++;
					if (idx2 == 4) {
						idx2 = 0;
					} else {
						idx1 = n3;
					}
					if (j != people[i][0].length - 1) {
						column[idx1][idx2 * 4] = "3.0";
						column[idx1][idx2 * 4 + 1] = "NO";
						column[idx1][idx2 * 4 + 2] = category1[idxCategory];
						column[idx1][idx2 * 4 + 3] = "";
						idx1++;
					}
				}

			}
			idxCategory++;
		}

		while (idx1 < n4) {
			column[idx1][idx2 * 4] = "0.0";
			column[idx1][idx2 * 4 + 1] = "";
			column[idx1][idx2 * 4 + 2] = "";
			column[idx1][idx2 * 4 + 3] = "";
			idx1++;
			if (idx1 == n2) {
				idx2++;
				if (idx2 == 4) {
					idx2 = 0;
				} else {
					idx1 = n1;
				}
			} else if (idx1 == n3) {
				idx2++;
				if (idx2 == 4) {
					idx2 = 0;
				} else {
					idx1 = n2;
				}
			} else if (idx1 == n4) {
				idx2++;
				if (idx2 == 4) {
					idx2 = 0;
				} else {
					idx1 = n3;
				}
			}
		}
	}

	public void performance() {
		SH.addMergedRegion(new CellRangeAddress(0, 2, 0, 11));
		SH.createRow(0).createCell(0).setCellValue(title);
		SH.getRow(0).getCell(0).setCellStyle(csTitle);

		SH.setColumnWidth(0, 1000);
		SH.setColumnWidth(3, 1000);
		SH.setColumnWidth(6, 1000);
		SH.setColumnWidth(9, 1000);
		SH.setColumnWidth(1, 4000);
		SH.setColumnWidth(2, 4000);
		SH.setColumnWidth(4, 4000);
		SH.setColumnWidth(5, 4000);
		SH.setColumnWidth(7, 4000);
		SH.setColumnWidth(8, 4000);
		SH.setColumnWidth(10, 4000);
		SH.setColumnWidth(11, 4000);

		SH.setAutobreaks(false);
		SH.setRowBreak(nRow - 1);
		SH.setRowBreak(nRow * 2 - 1);
		SH.setRowBreak(nRow * 3 - 1);
		SH.setColumnBreak(11);

		for (int k = 3; k < nRow * 3; k++) {
			SH.createRow(k);
			for (int i = 0; i < 4; i++) {
				SH.getRow(k).createCell(i * 3).setCellValue(column[k - 3][i * 4 + 1]);
				SH.getRow(k).createCell(i * 3 + 1).setCellValue(column[k - 3][i * 4 + 2]);
				SH.getRow(k).createCell(i * 3 + 2).setCellValue(column[k - 3][i * 4 + 3]);
				if (column[k - 3][i * 4].equals("0.0")) {
					// Normal
					if (k == nRow - 1 || k == nRow * 2 - 1 || k == nRow * 3 - 1) {
						SH.getRow(k).getCell(i * 3).setCellStyle(csNormalLastLeft);
						SH.getRow(k).getCell(i * 3 + 1).setCellStyle(csNormalLastCenter);
						SH.getRow(k).getCell(i * 3 + 2).setCellStyle(csNormalLastRight);
					} else {
						SH.getRow(k).getCell(i * 3).setCellStyle(csNormalLeft);
						SH.getRow(k).getCell(i * 3 + 1).setCellStyle(csNormalCenter);
						SH.getRow(k).getCell(i * 3 + 2).setCellStyle(csNormalRight);
					}
				} else {
					// Category
					SH.addMergedRegion(new CellRangeAddress(k, k, i * 3 + 1, i * 3 + 2));
					SH.getRow(k).getCell(i * 3).setCellStyle(csCategoryLeft);
					SH.getRow(k).getCell(i * 3 + 1).setCellStyle(csCategoryCenter);
					SH.getRow(k).getCell(i * 3 + 2).setCellStyle(csCategoryRight);
				}
			}
		}
	}

}
