package com.jin.attendanceserver.formatting;

import java.io.FileOutputStream;
import java.io.IOException;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
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

public class FpFormatting {

	final private XSSFWorkbook ATTENDANCE, BELIEVER, WORDMOVEMENT, WB1, WB2;

	String path1, path2, title1, title2;
	String[] category1 = { "교역자", "장로", "안수집사", "권사", "서리집사(남)", "서리집사(여)", "권찰", "성도(남)", "성도(여)", "청년", "대학", "기타",
			"거창지교회", "상주지교회", "경주지교회", "성주지교회", "영주지교회", "북대구지교회", "포항지교회", "대전지교회", "부여지교회", "서대구지교회", "창원지교회",
			"북부산지교회", "지교회기타" };
	XSSFSheet SH1, SH2, attendance;
	int num_attendance;

	int[] N, row, listNum, attNum, bN, wmN;
	int[][] volume;
	int nRow1, nRow2, nLeft, nRight;
	String[][][] people, list, bList, wmList;
	String[][] column;
	ArrayList<String> people1, people2;
	ArrayList<Integer> peopleIdx1, peopleIdx2;

	XSSFFont fTitle1, fNormal1, fTitle2, fNormal2;
	XSSFCellStyle csTitle1, csCategoryLeft, csCategoryCenter, csCategoryRight, csNormalLeft, csNormalCenter,
			csNormalRight, csDedicationLeft, csDedicationCenter, csDedicationRight, csAttendanceLeft,
			csAttendanceCenter, csAttendanceRight, csNormalLastLeft, csNormalLastCenter, csNormalLastRight,
			csDedicationLastLeft, csDedicationLastCenter, csDedicationLastRight, csAttendanceLastLeft,
			csAttendanceLastCenter, csAttendanceLastRight, csSummaryTitleLeft, csSummaryTitleCenter,
			csSummaryTitleRight, csNormalYellowLeft, csNormalYellowCenter, csNormalYellowRight;
	XSSFCellStyle csTitle2, csIdxBc, csIdxRemeet, csIdxBeliever, csMenuAboveLeft, csMenuAboveCenter, csMenuAboveRight,
			csMenuBelowLeft, csMenuBelowCenter, csMenuBelowRight, csMenuNormalLeft, csMenuNormalCenter,
			csMenuNormalRight, csBC, csAT, csBCf, csATf, csNormalAboveLeft, csNormalAboveCenter, csNormalAboveRight,
			csRemeet, csRemeetAbove, csSum, csYellow, csLine;

	short cCategory = IndexedColors.PALE_BLUE.getIndex();
	short cNormal = IndexedColors.WHITE.getIndex();
	short cDedication = IndexedColors.LIGHT_ORANGE.getIndex();
	short cAttendance = IndexedColors.LIME.getIndex();
	short cYellow = IndexedColors.YELLOW.getIndex();

	public FpFormatting(XSSFWorkbook attendance, XSSFWorkbook believer, XSSFWorkbook wordmovement, String directory_path) {
		ATTENDANCE = attendance;
		BELIEVER = believer;
		WORDMOVEMENT = wordmovement;

		SimpleDateFormat formatter = new SimpleDateFormat("yyyy년 MM월 dd일");
		String time = formatter.format(new Date());
		title1 = time + " 전교인 주일 현장전도";
		title2 = time + " 전교인 주일 현장전도 보고서";

		path1 = directory_path + "\\archive\\FpAtt\\" + time + " 출석.xlsx";
		path2 = directory_path + "\\archive\\FpAtt\\" + time + " 열매.xlsx";

		WB1 = new XSSFWorkbook();
		SH1 = WB1.createSheet("출석");
		WB2 = new XSSFWorkbook();
		SH2 = WB2.createSheet("열매");
		SH1.setPrintGridlines(false);
		SH1.setDisplayGridlines(false);
		SH2.setPrintGridlines(false);
		SH2.setDisplayGridlines(false);

		setFont();
		setCellStyle1();
		init1();
		performance1();
		setCellStyle2();
		init2();
		performance2();

		try {
			WB1.write(new FileOutputStream(path1));
			WB1.close();
			WB2.write(new FileOutputStream(path2));
			WB2.close();
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
	}

	public void setFont() {
		fTitle1 = WB1.createFont();
		fTitle1.setBold(true);
		fTitle1.setColor(IndexedColors.BLACK.getIndex());
		fTitle1.setFontName("맑은 고딕");
		fTitle1.setFontHeightInPoints((short) 20);

		fNormal1 = WB1.createFont();
		fNormal1.setBold(false);
		fNormal1.setColor(IndexedColors.BLACK.getIndex());
		fNormal1.setFontName("맑은 고딕");
		fNormal1.setFontHeightInPoints((short) 10);

		fTitle2 = WB2.createFont();
		fTitle2.setBold(true);
		fTitle2.setColor(IndexedColors.BLACK.getIndex());
		fTitle2.setFontName("맑은 고딕");
		fTitle2.setFontHeightInPoints((short) 20);

		fNormal2 = WB2.createFont();
		fNormal2.setBold(false);
		fNormal2.setColor(IndexedColors.BLACK.getIndex());
		fNormal2.setFontName("맑은 고딕");
		fNormal2.setFontHeightInPoints((short) 9);
	}

	public void setCellStyle1() {
		csTitle1 = WB1.createCellStyle();
		csTitle1.setAlignment(HorizontalAlignment.CENTER);
		csTitle1.setVerticalAlignment(VerticalAlignment.TOP);
		csTitle1.setFont(fTitle1);

		csCategoryLeft = WB1.createCellStyle();
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
		csCategoryLeft.setFont(fNormal1);

		csCategoryCenter = WB1.createCellStyle();
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
		csCategoryCenter.setFont(fNormal1);

		csCategoryRight = WB1.createCellStyle();
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
		csCategoryRight.setFont(fNormal1);

		csNormalLeft = WB1.createCellStyle();
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
		csNormalLeft.setFont(fNormal1);

		csNormalCenter = WB1.createCellStyle();
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
		csNormalCenter.setFont(fNormal1);

		csNormalRight = WB1.createCellStyle();
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
		csNormalRight.setFont(fNormal1);

		csDedicationLeft = WB1.createCellStyle();
		csDedicationLeft.setAlignment(HorizontalAlignment.CENTER);
		csDedicationLeft.setVerticalAlignment(VerticalAlignment.CENTER);
		csDedicationLeft.setBorderTop(BorderStyle.THIN);
		csDedicationLeft.setTopBorderColor(IndexedColors.BLACK.getIndex());
		csDedicationLeft.setBorderBottom(BorderStyle.THIN);
		csDedicationLeft.setBottomBorderColor(IndexedColors.BLACK.getIndex());
		csDedicationLeft.setBorderLeft(BorderStyle.THICK);
		csDedicationLeft.setLeftBorderColor(IndexedColors.BLACK.getIndex());
		csDedicationLeft.setBorderRight(BorderStyle.THIN);
		csDedicationLeft.setRightBorderColor(IndexedColors.BLACK.getIndex());
		csDedicationLeft.setFillForegroundColor(cDedication);
		csDedicationLeft.setFillPattern(FillPatternType.SOLID_FOREGROUND);
		csDedicationLeft.setFont(fNormal1);

		csDedicationCenter = WB1.createCellStyle();
		csDedicationCenter.setAlignment(HorizontalAlignment.CENTER);
		csDedicationCenter.setVerticalAlignment(VerticalAlignment.CENTER);
		csDedicationCenter.setBorderTop(BorderStyle.THIN);
		csDedicationCenter.setTopBorderColor(IndexedColors.BLACK.getIndex());
		csDedicationCenter.setBorderBottom(BorderStyle.THIN);
		csDedicationCenter.setBottomBorderColor(IndexedColors.BLACK.getIndex());
		csDedicationCenter.setBorderLeft(BorderStyle.THIN);
		csDedicationCenter.setLeftBorderColor(IndexedColors.BLACK.getIndex());
		csDedicationCenter.setBorderRight(BorderStyle.THIN);
		csDedicationCenter.setRightBorderColor(IndexedColors.BLACK.getIndex());
		csDedicationCenter.setFillForegroundColor(cDedication);
		csDedicationCenter.setFillPattern(FillPatternType.SOLID_FOREGROUND);
		csDedicationCenter.setFont(fNormal1);

		csDedicationRight = WB1.createCellStyle();
		csDedicationRight.setAlignment(HorizontalAlignment.CENTER);
		csDedicationRight.setVerticalAlignment(VerticalAlignment.CENTER);
		csDedicationRight.setBorderTop(BorderStyle.THIN);
		csDedicationRight.setTopBorderColor(IndexedColors.BLACK.getIndex());
		csDedicationRight.setBorderBottom(BorderStyle.THIN);
		csDedicationRight.setBottomBorderColor(IndexedColors.BLACK.getIndex());
		csDedicationRight.setBorderLeft(BorderStyle.THIN);
		csDedicationRight.setLeftBorderColor(IndexedColors.BLACK.getIndex());
		csDedicationRight.setBorderRight(BorderStyle.THICK);
		csDedicationRight.setRightBorderColor(IndexedColors.BLACK.getIndex());
		csDedicationRight.setFillForegroundColor(cDedication);
		csDedicationRight.setFillPattern(FillPatternType.SOLID_FOREGROUND);
		csDedicationRight.setFont(fNormal1);

		csAttendanceLeft = WB1.createCellStyle();
		csAttendanceLeft.setAlignment(HorizontalAlignment.CENTER);
		csAttendanceLeft.setVerticalAlignment(VerticalAlignment.CENTER);
		csAttendanceLeft.setBorderTop(BorderStyle.THIN);
		csAttendanceLeft.setTopBorderColor(IndexedColors.BLACK.getIndex());
		csAttendanceLeft.setBorderBottom(BorderStyle.THIN);
		csAttendanceLeft.setBottomBorderColor(IndexedColors.BLACK.getIndex());
		csAttendanceLeft.setBorderLeft(BorderStyle.THICK);
		csAttendanceLeft.setLeftBorderColor(IndexedColors.BLACK.getIndex());
		csAttendanceLeft.setBorderRight(BorderStyle.THIN);
		csAttendanceLeft.setRightBorderColor(IndexedColors.BLACK.getIndex());
		csAttendanceLeft.setFillForegroundColor(cAttendance);
		csAttendanceLeft.setFillPattern(FillPatternType.SOLID_FOREGROUND);
		csAttendanceLeft.setFont(fNormal1);

		csAttendanceCenter = WB1.createCellStyle();
		csAttendanceCenter.setAlignment(HorizontalAlignment.CENTER);
		csAttendanceCenter.setVerticalAlignment(VerticalAlignment.CENTER);
		csAttendanceCenter.setBorderTop(BorderStyle.THIN);
		csAttendanceCenter.setTopBorderColor(IndexedColors.BLACK.getIndex());
		csAttendanceCenter.setBorderBottom(BorderStyle.THIN);
		csAttendanceCenter.setBottomBorderColor(IndexedColors.BLACK.getIndex());
		csAttendanceCenter.setBorderLeft(BorderStyle.THIN);
		csAttendanceCenter.setLeftBorderColor(IndexedColors.BLACK.getIndex());
		csAttendanceCenter.setBorderRight(BorderStyle.THIN);
		csAttendanceCenter.setRightBorderColor(IndexedColors.BLACK.getIndex());
		csAttendanceCenter.setFillForegroundColor(cAttendance);
		csAttendanceCenter.setFillPattern(FillPatternType.SOLID_FOREGROUND);
		csAttendanceCenter.setFont(fNormal1);

		csAttendanceRight = WB1.createCellStyle();
		csAttendanceRight.setAlignment(HorizontalAlignment.CENTER);
		csAttendanceRight.setVerticalAlignment(VerticalAlignment.CENTER);
		csAttendanceRight.setBorderTop(BorderStyle.THIN);
		csAttendanceRight.setTopBorderColor(IndexedColors.BLACK.getIndex());
		csAttendanceRight.setBorderBottom(BorderStyle.THIN);
		csAttendanceRight.setBottomBorderColor(IndexedColors.BLACK.getIndex());
		csAttendanceRight.setBorderLeft(BorderStyle.THIN);
		csAttendanceRight.setLeftBorderColor(IndexedColors.BLACK.getIndex());
		csAttendanceRight.setBorderRight(BorderStyle.THICK);
		csAttendanceRight.setRightBorderColor(IndexedColors.BLACK.getIndex());
		csAttendanceRight.setFillForegroundColor(cAttendance);
		csAttendanceRight.setFillPattern(FillPatternType.SOLID_FOREGROUND);
		csAttendanceRight.setFont(fNormal1);

		csNormalLastLeft = WB1.createCellStyle();
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
		csNormalLastLeft.setFont(fNormal1);

		csNormalLastCenter = WB1.createCellStyle();
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
		csNormalLastCenter.setFont(fNormal1);

		csNormalLastRight = WB1.createCellStyle();
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
		csNormalLastRight.setFont(fNormal1);

		csDedicationLastLeft = WB1.createCellStyle();
		csDedicationLastLeft.setAlignment(HorizontalAlignment.CENTER);
		csDedicationLastLeft.setVerticalAlignment(VerticalAlignment.CENTER);
		csDedicationLastLeft.setBorderTop(BorderStyle.THIN);
		csDedicationLastLeft.setTopBorderColor(IndexedColors.BLACK.getIndex());
		csDedicationLastLeft.setBorderBottom(BorderStyle.THICK);
		csDedicationLastLeft.setBottomBorderColor(IndexedColors.BLACK.getIndex());
		csDedicationLastLeft.setBorderLeft(BorderStyle.THICK);
		csDedicationLastLeft.setLeftBorderColor(IndexedColors.BLACK.getIndex());
		csDedicationLastLeft.setBorderRight(BorderStyle.THIN);
		csDedicationLastLeft.setRightBorderColor(IndexedColors.BLACK.getIndex());
		csDedicationLastLeft.setFillForegroundColor(cDedication);
		csDedicationLastLeft.setFillPattern(FillPatternType.SOLID_FOREGROUND);
		csDedicationLastLeft.setFont(fNormal1);

		csDedicationLastCenter = WB1.createCellStyle();
		csDedicationLastCenter.setAlignment(HorizontalAlignment.CENTER);
		csDedicationLastCenter.setVerticalAlignment(VerticalAlignment.CENTER);
		csDedicationLastCenter.setBorderTop(BorderStyle.THIN);
		csDedicationLastCenter.setTopBorderColor(IndexedColors.BLACK.getIndex());
		csDedicationLastCenter.setBorderBottom(BorderStyle.THICK);
		csDedicationLastCenter.setBottomBorderColor(IndexedColors.BLACK.getIndex());
		csDedicationLastCenter.setBorderLeft(BorderStyle.THIN);
		csDedicationLastCenter.setLeftBorderColor(IndexedColors.BLACK.getIndex());
		csDedicationLastCenter.setBorderRight(BorderStyle.THIN);
		csDedicationLastCenter.setRightBorderColor(IndexedColors.BLACK.getIndex());
		csDedicationLastCenter.setFillForegroundColor(cDedication);
		csDedicationLastCenter.setFillPattern(FillPatternType.SOLID_FOREGROUND);
		csDedicationLastCenter.setFont(fNormal1);

		csDedicationLastRight = WB1.createCellStyle();
		csDedicationLastRight.setAlignment(HorizontalAlignment.CENTER);
		csDedicationLastRight.setVerticalAlignment(VerticalAlignment.CENTER);
		csDedicationLastRight.setBorderTop(BorderStyle.THIN);
		csDedicationLastRight.setTopBorderColor(IndexedColors.BLACK.getIndex());
		csDedicationLastRight.setBorderBottom(BorderStyle.THICK);
		csDedicationLastRight.setBottomBorderColor(IndexedColors.BLACK.getIndex());
		csDedicationLastRight.setBorderLeft(BorderStyle.THIN);
		csDedicationLastRight.setLeftBorderColor(IndexedColors.BLACK.getIndex());
		csDedicationLastRight.setBorderRight(BorderStyle.THICK);
		csDedicationLastRight.setRightBorderColor(IndexedColors.BLACK.getIndex());
		csDedicationLastRight.setFillForegroundColor(cDedication);
		csDedicationLastRight.setFillPattern(FillPatternType.SOLID_FOREGROUND);
		csDedicationLastRight.setFont(fNormal1);

		csAttendanceLastLeft = WB1.createCellStyle();
		csAttendanceLastLeft.setAlignment(HorizontalAlignment.CENTER);
		csAttendanceLastLeft.setVerticalAlignment(VerticalAlignment.CENTER);
		csAttendanceLastLeft.setBorderTop(BorderStyle.THIN);
		csAttendanceLastLeft.setTopBorderColor(IndexedColors.BLACK.getIndex());
		csAttendanceLastLeft.setBorderBottom(BorderStyle.THICK);
		csAttendanceLastLeft.setBottomBorderColor(IndexedColors.BLACK.getIndex());
		csAttendanceLastLeft.setBorderLeft(BorderStyle.THICK);
		csAttendanceLastLeft.setLeftBorderColor(IndexedColors.BLACK.getIndex());
		csAttendanceLastLeft.setBorderRight(BorderStyle.THIN);
		csAttendanceLastLeft.setRightBorderColor(IndexedColors.BLACK.getIndex());
		csAttendanceLastLeft.setFillForegroundColor(cAttendance);
		csAttendanceLastLeft.setFillPattern(FillPatternType.SOLID_FOREGROUND);
		csAttendanceLastLeft.setFont(fNormal1);

		csAttendanceLastCenter = WB1.createCellStyle();
		csAttendanceLastCenter.setAlignment(HorizontalAlignment.CENTER);
		csAttendanceLastCenter.setVerticalAlignment(VerticalAlignment.CENTER);
		csAttendanceLastCenter.setBorderTop(BorderStyle.THIN);
		csAttendanceLastCenter.setTopBorderColor(IndexedColors.BLACK.getIndex());
		csAttendanceLastCenter.setBorderBottom(BorderStyle.THICK);
		csAttendanceLastCenter.setBottomBorderColor(IndexedColors.BLACK.getIndex());
		csAttendanceLastCenter.setBorderLeft(BorderStyle.THIN);
		csAttendanceLastCenter.setLeftBorderColor(IndexedColors.BLACK.getIndex());
		csAttendanceLastCenter.setBorderRight(BorderStyle.THIN);
		csAttendanceLastCenter.setRightBorderColor(IndexedColors.BLACK.getIndex());
		csAttendanceLastCenter.setFillForegroundColor(cAttendance);
		csAttendanceLastCenter.setFillPattern(FillPatternType.SOLID_FOREGROUND);
		csAttendanceLastCenter.setFont(fNormal1);

		csAttendanceLastRight = WB1.createCellStyle();
		csAttendanceLastRight.setAlignment(HorizontalAlignment.CENTER);
		csAttendanceLastRight.setVerticalAlignment(VerticalAlignment.CENTER);
		csAttendanceLastRight.setBorderTop(BorderStyle.THIN);
		csAttendanceLastRight.setTopBorderColor(IndexedColors.BLACK.getIndex());
		csAttendanceLastRight.setBorderBottom(BorderStyle.THICK);
		csAttendanceLastRight.setBottomBorderColor(IndexedColors.BLACK.getIndex());
		csAttendanceLastRight.setBorderLeft(BorderStyle.THIN);
		csAttendanceLastRight.setLeftBorderColor(IndexedColors.BLACK.getIndex());
		csAttendanceLastRight.setBorderRight(BorderStyle.THICK);
		csAttendanceLastRight.setRightBorderColor(IndexedColors.BLACK.getIndex());
		csAttendanceLastRight.setFillForegroundColor(cAttendance);
		csAttendanceLastRight.setFillPattern(FillPatternType.SOLID_FOREGROUND);
		csAttendanceLastRight.setFont(fNormal1);

		csSummaryTitleLeft = WB1.createCellStyle();
		csSummaryTitleLeft.setAlignment(HorizontalAlignment.CENTER);
		csSummaryTitleLeft.setVerticalAlignment(VerticalAlignment.CENTER);
		csSummaryTitleLeft.setBorderTop(BorderStyle.THICK);
		csSummaryTitleLeft.setTopBorderColor(IndexedColors.BLACK.getIndex());
		csSummaryTitleLeft.setBorderBottom(BorderStyle.THICK);
		csSummaryTitleLeft.setBottomBorderColor(IndexedColors.BLACK.getIndex());
		csSummaryTitleLeft.setBorderLeft(BorderStyle.THICK);
		csSummaryTitleLeft.setLeftBorderColor(IndexedColors.BLACK.getIndex());
		csSummaryTitleLeft.setBorderRight(BorderStyle.THIN);
		csSummaryTitleLeft.setRightBorderColor(IndexedColors.BLACK.getIndex());
		csSummaryTitleLeft.setFillForegroundColor(cNormal);
		csSummaryTitleLeft.setFillPattern(FillPatternType.SOLID_FOREGROUND);
		csSummaryTitleLeft.setFont(fNormal1);

		csSummaryTitleCenter = WB1.createCellStyle();
		csSummaryTitleCenter.setAlignment(HorizontalAlignment.CENTER);
		csSummaryTitleCenter.setVerticalAlignment(VerticalAlignment.CENTER);
		csSummaryTitleCenter.setBorderTop(BorderStyle.THICK);
		csSummaryTitleCenter.setTopBorderColor(IndexedColors.BLACK.getIndex());
		csSummaryTitleCenter.setBorderBottom(BorderStyle.THICK);
		csSummaryTitleCenter.setBottomBorderColor(IndexedColors.BLACK.getIndex());
		csSummaryTitleCenter.setBorderLeft(BorderStyle.THIN);
		csSummaryTitleCenter.setLeftBorderColor(IndexedColors.BLACK.getIndex());
		csSummaryTitleCenter.setBorderRight(BorderStyle.THIN);
		csSummaryTitleCenter.setRightBorderColor(IndexedColors.BLACK.getIndex());
		csSummaryTitleCenter.setFillForegroundColor(cNormal);
		csSummaryTitleCenter.setFillPattern(FillPatternType.SOLID_FOREGROUND);
		csSummaryTitleCenter.setFont(fNormal1);

		csSummaryTitleRight = WB1.createCellStyle();
		csSummaryTitleRight.setAlignment(HorizontalAlignment.CENTER);
		csSummaryTitleRight.setVerticalAlignment(VerticalAlignment.CENTER);
		csSummaryTitleRight.setBorderTop(BorderStyle.THICK);
		csSummaryTitleRight.setTopBorderColor(IndexedColors.BLACK.getIndex());
		csSummaryTitleRight.setBorderBottom(BorderStyle.THICK);
		csSummaryTitleRight.setBottomBorderColor(IndexedColors.BLACK.getIndex());
		csSummaryTitleRight.setBorderLeft(BorderStyle.THIN);
		csSummaryTitleRight.setLeftBorderColor(IndexedColors.BLACK.getIndex());
		csSummaryTitleRight.setBorderRight(BorderStyle.THICK);
		csSummaryTitleRight.setRightBorderColor(IndexedColors.BLACK.getIndex());
		csSummaryTitleRight.setFillForegroundColor(cYellow);
		csSummaryTitleRight.setFillPattern(FillPatternType.SOLID_FOREGROUND);
		csSummaryTitleRight.setFont(fNormal1);

		csNormalYellowLeft = WB1.createCellStyle();
		csNormalYellowLeft.setAlignment(HorizontalAlignment.CENTER);
		csNormalYellowLeft.setVerticalAlignment(VerticalAlignment.CENTER);
		csNormalYellowLeft.setBorderTop(BorderStyle.THIN);
		csNormalYellowLeft.setTopBorderColor(IndexedColors.BLACK.getIndex());
		csNormalYellowLeft.setBorderBottom(BorderStyle.THIN);
		csNormalYellowLeft.setBottomBorderColor(IndexedColors.BLACK.getIndex());
		csNormalYellowLeft.setBorderLeft(BorderStyle.THICK);
		csNormalYellowLeft.setLeftBorderColor(IndexedColors.BLACK.getIndex());
		csNormalYellowLeft.setBorderRight(BorderStyle.THIN);
		csNormalYellowLeft.setRightBorderColor(IndexedColors.BLACK.getIndex());
		csNormalYellowLeft.setFillForegroundColor(cYellow);
		csNormalYellowLeft.setFillPattern(FillPatternType.SOLID_FOREGROUND);
		csNormalYellowLeft.setFont(fNormal1);

		csNormalYellowCenter = WB1.createCellStyle();
		csNormalYellowCenter.setAlignment(HorizontalAlignment.CENTER);
		csNormalYellowCenter.setVerticalAlignment(VerticalAlignment.CENTER);
		csNormalYellowCenter.setBorderTop(BorderStyle.THIN);
		csNormalYellowCenter.setTopBorderColor(IndexedColors.BLACK.getIndex());
		csNormalYellowCenter.setBorderBottom(BorderStyle.THIN);
		csNormalYellowCenter.setBottomBorderColor(IndexedColors.BLACK.getIndex());
		csNormalYellowCenter.setBorderLeft(BorderStyle.THIN);
		csNormalYellowCenter.setLeftBorderColor(IndexedColors.BLACK.getIndex());
		csNormalYellowCenter.setBorderRight(BorderStyle.THIN);
		csNormalYellowCenter.setRightBorderColor(IndexedColors.BLACK.getIndex());
		csNormalYellowCenter.setFillForegroundColor(cYellow);
		csNormalYellowCenter.setFillPattern(FillPatternType.SOLID_FOREGROUND);
		csNormalYellowCenter.setFont(fNormal1);

		csNormalYellowRight = WB1.createCellStyle();
		csNormalYellowRight.setAlignment(HorizontalAlignment.CENTER);
		csNormalYellowRight.setVerticalAlignment(VerticalAlignment.CENTER);
		csNormalYellowRight.setBorderTop(BorderStyle.THIN);
		csNormalYellowRight.setTopBorderColor(IndexedColors.BLACK.getIndex());
		csNormalYellowRight.setBorderBottom(BorderStyle.THIN);
		csNormalYellowRight.setBottomBorderColor(IndexedColors.BLACK.getIndex());
		csNormalYellowRight.setBorderLeft(BorderStyle.THIN);
		csNormalYellowRight.setLeftBorderColor(IndexedColors.BLACK.getIndex());
		csNormalYellowRight.setBorderRight(BorderStyle.THICK);
		csNormalYellowRight.setRightBorderColor(IndexedColors.BLACK.getIndex());
		csNormalYellowRight.setFillForegroundColor(cDedication);
		csNormalYellowRight.setFillPattern(FillPatternType.SOLID_FOREGROUND);
		csNormalYellowRight.setFont(fNormal1);
	}

	public void init1() {
		attendance = ATTENDANCE.getSheetAt(0);
		num_attendance = attendance.getPhysicalNumberOfRows();
		N = new int[25];
		// 모든 category는 category title을 포함함
		int N_sum = 12;
		for (int i = 0; i < N.length; i++) {
			N[i] = 0;
		}
		for (int i = 1; i < num_attendance; i++) {
			String c = attendance.getRow(i).getCell(57).toString();
			for (int j = 0; j < category1.length; j++) {
				if (c.equals(category1[j])) {
					N[j]++;
					N_sum++;
					break;
				}
			}
		}
		people = new String[N.length][4][];
		int week = Calendar.getInstance().get(Calendar.WEEK_OF_YEAR);
		if (week == 1) {
			if (Calendar.getInstance().get(Calendar.MONTH) == 11) {
				week = 53;
			}
		}
		week += 3;
		for (int i = 0; i < N.length; i++) {
			people[i][0] = new String[N[i]];
			people[i][1] = new String[N[i]];
			people[i][2] = new String[N[i]];
			people[i][3] = new String[N[i]];
			N[i] = 0;
		}
		volume = new int[25][3];
		for (int i = 0; i < volume.length; i++) {
			volume[i][0] = 0;
			volume[i][1] = 0;
			volume[i][2] = 0;
		}
		for (int i = 1; i < num_attendance; i++) {
			String c = attendance.getRow(i).getCell(57).toString();
			for (int j = 0; j < category1.length; j++) {
				if (c.equals(category1[j])) {
					// NAME
					people[j][0][N[j]] = attendance.getRow(i).getCell(0).toString();
					// GROUP
					people[j][1][N[j]] = attendance.getRow(i).getCell(1).toString();
					// CONTENT
					try {
						people[j][2][N[j]] = attendance.getRow(i).getCell(4).toString();
					} catch (NullPointerException e) {
						people[j][2][N[j]] = "";
					}
					// CHECK
					people[j][3][N[j]] = attendance.getRow(i).getCell(week).toString();
					if (people[j][3][N[j]].equals("0.0")) {
						volume[j][0]++;
					} else if (people[j][3][N[j]].equals("1.0")) {
						volume[j][1]++;
					} else {
						volume[j][2]++;
					}
					N[j]++;
					break;
				}
			}
		}

		people1 = new ArrayList<String>();
		people2 = new ArrayList<String>();
		peopleIdx1 = new ArrayList<Integer>();
		peopleIdx2 = new ArrayList<Integer>();

		for (int i = 0; i < 19; i++) {
			XSSFSheet etc = ATTENDANCE.getSheetAt(i + 1);
			int num = etc.getPhysicalNumberOfRows() - 1;
			for (int j = 0; j < num; j++) {
				if (i < 7) {
					people1.add(etc.getRow(j + 1).getCell(0).toString());
					peopleIdx1.add(i);
				} else {
					people2.add(etc.getRow(j + 1).getCell(0).toString());
					peopleIdx2.add(i);
				}
			}
		}

		int idx1 = 0;
		while (idx1 < people1.size()) {
			int idx2 = 0;
			while (idx2 < people1.size()) {
				if (idx1 != idx2 && people1.get(idx1).equals(people1.get(idx2))) {
					people1.remove(idx2);
					peopleIdx1.remove(idx2);
				} else {
					idx2++;
				}
			}
			idx1++;
		}
		idx1 = 0;
		while (idx1 < people2.size()) {
			int idx2 = 0;
			while (idx2 < people2.size()) {
				if (idx1 != idx2 && people2.get(idx1).equals(people2.get(idx2))) {
					people2.remove(idx2);
					peopleIdx2.remove(idx2);
				} else {
					idx2++;
				}
			}
			idx1++;
		}
		idx1 = 0;
		while (idx1 < people2.size()) {
			int idx2 = 0;
			while (idx2 < people1.size()) {
				if (people2.get(idx1).equals(people1.get(idx2))) {
					people1.remove(idx2);
					peopleIdx1.remove(idx2);
				} else {
					idx2++;
				}
			}
			idx1++;
		}

		people[11][0] = new String[(people1.size() + 1) / 2];
		people[24][0] = new String[(people2.size() + 1) / 2];

		// sum + (12 - 1) + 50 <= 12n - 80
		nRow1 = (N_sum + (people1.size() + 1) / 2 + (people2.size() + 1) / 2 + 141) / 12;

		column = new String[nRow1 * 3 - 20][16];
		idx1 = 0;
		int idx2 = 0;
		int idxCategory = 0;
		int n1 = 0;
		int n2 = nRow1 - 3;
		int n3 = nRow1 * 2 - 3;
		int n4 = nRow1 * 3 - 20;

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
				if (idxCategory < 11) {
					// 구미
					column[idx1][idx2 * 4] = people[i][3][j];
					column[idx1][idx2 * 4 + 1] = String.valueOf(idxNum);
					column[idx1][idx2 * 4 + 2] = people[i][0][j];
					column[idx1][idx2 * 4 + 3] = people[i][2][j];
					idxNum++;
				} else if (idxCategory == 11) {
					// 구미 기타
					column[idx1][idx2 * 4] = "2.0";
					column[idx1][idx2 * 4 + 1] = String.valueOf(idxNum);
					column[idx1][idx2 * 4 + 2] = people1.get(idxNum - 1);
					if (idxNum != people1.size()) {
						column[idx1][idx2 * 4 + 3] = people1.get(idxNum);
					} else {
						column[idx1][idx2 * 4 + 3] = "";
					}
					idxNum += 2;
				} else if (idxCategory >= 12 && idxCategory < 24) {
					// 지교회
					column[idx1][idx2 * 4] = people[i][3][j];
					column[idx1][idx2 * 4 + 1] = String.valueOf(idxNum);
					column[idx1][idx2 * 4 + 2] = people[i][0][j];
					column[idx1][idx2 * 4 + 3] = people[i][1][j];
					idxNum++;
				} else {
					// 지교회 기타
					column[idx1][idx2 * 4] = "2.0";
					column[idx1][idx2 * 4 + 1] = String.valueOf(idxNum);
					column[idx1][idx2 * 4 + 2] = people2.get(idxNum - 1);
					if (idxNum != people2.size()) {
						column[idx1][idx2 * 4 + 3] = people2.get(idxNum);
					} else {
						column[idx1][idx2 * 4 + 3] = "";
					}
					idxNum += 2;
				}
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

	public void performance1() {
		SH1.addMergedRegion(new CellRangeAddress(0, 2, 0, 11));
		SH1.createRow(0).createCell(0).setCellValue(title1);
		SH1.getRow(0).getCell(0).setCellStyle(csTitle1);

		SH1.setColumnWidth(0, 1000);
		SH1.setColumnWidth(3, 1000);
		SH1.setColumnWidth(6, 1000);
		SH1.setColumnWidth(9, 1000);
		SH1.setColumnWidth(1, 4000);
		SH1.setColumnWidth(2, 4000);
		SH1.setColumnWidth(4, 4000);
		SH1.setColumnWidth(5, 4000);
		SH1.setColumnWidth(7, 4000);
		SH1.setColumnWidth(8, 4000);
		SH1.setColumnWidth(10, 4000);
		SH1.setColumnWidth(11, 4000);

		SH1.setAutobreaks(false);
		SH1.setRowBreak(nRow1 - 1);
		SH1.setRowBreak(nRow1 * 2 - 1);
		SH1.setRowBreak(nRow1 * 3 - 1);
		SH1.setColumnBreak(11);

		for (int k = 3; k < nRow1 * 3 - 17; k++) {
			SH1.createRow(k);
			for (int i = 0; i < 4; i++) {
				SH1.getRow(k).createCell(i * 3).setCellValue(column[k - 3][i * 4 + 1]);
				SH1.getRow(k).createCell(i * 3 + 1).setCellValue(column[k - 3][i * 4 + 2]);
				SH1.getRow(k).createCell(i * 3 + 2).setCellValue(column[k - 3][i * 4 + 3]);
				if (column[k - 3][i * 4].equals("0.0")) {
					// Normal
					if (k == nRow1 - 1 || k == nRow1 * 2 - 1 || k == nRow1 * 3 - 18) {
						SH1.getRow(k).getCell(i * 3).setCellStyle(csNormalLastLeft);
						SH1.getRow(k).getCell(i * 3 + 1).setCellStyle(csNormalLastCenter);
						SH1.getRow(k).getCell(i * 3 + 2).setCellStyle(csNormalLastRight);
					} else {
						SH1.getRow(k).getCell(i * 3).setCellStyle(csNormalLeft);
						SH1.getRow(k).getCell(i * 3 + 1).setCellStyle(csNormalCenter);
						SH1.getRow(k).getCell(i * 3 + 2).setCellStyle(csNormalRight);
					}
				} else if (column[k - 3][i * 4].equals("1.0")) {
					// Dedication
					if (k == nRow1 - 1 || k == nRow1 * 2 - 1 || k == nRow1 * 3 - 18) {
						SH1.getRow(k).getCell(i * 3).setCellStyle(csDedicationLastLeft);
						SH1.getRow(k).getCell(i * 3 + 1).setCellStyle(csDedicationLastCenter);
						SH1.getRow(k).getCell(i * 3 + 2).setCellStyle(csDedicationLastRight);
					} else {
						SH1.getRow(k).getCell(i * 3).setCellStyle(csDedicationLeft);
						SH1.getRow(k).getCell(i * 3 + 1).setCellStyle(csDedicationCenter);
						SH1.getRow(k).getCell(i * 3 + 2).setCellStyle(csDedicationRight);
					}
				} else if (column[k - 3][i * 4].equals("2.0")) {
					// Attendance
					if (k == nRow1 - 1 || k == nRow1 * 2 - 1 || k == nRow1 * 3 - 18) {
						SH1.getRow(k).getCell(i * 3).setCellStyle(csAttendanceLastLeft);
						SH1.getRow(k).getCell(i * 3 + 1).setCellStyle(csAttendanceLastCenter);
						SH1.getRow(k).getCell(i * 3 + 2).setCellStyle(csAttendanceLastRight);
					} else {
						SH1.getRow(k).getCell(i * 3).setCellStyle(csAttendanceLeft);
						SH1.getRow(k).getCell(i * 3 + 1).setCellStyle(csAttendanceCenter);
						SH1.getRow(k).getCell(i * 3 + 2).setCellStyle(csAttendanceRight);
					}
				} else {
					// Category
					SH1.addMergedRegion(new CellRangeAddress(k, k, i * 3 + 1, i * 3 + 2));
					SH1.getRow(k).getCell(i * 3).setCellStyle(csCategoryLeft);
					SH1.getRow(k).getCell(i * 3 + 1).setCellStyle(csCategoryCenter);
					SH1.getRow(k).getCell(i * 3 + 2).setCellStyle(csCategoryRight);
				}
			}
		}

		SH1.addMergedRegion(new CellRangeAddress(nRow1 * 3 - 17, nRow1 * 3 - 17, 0, 11));
		SH1.createRow(nRow1 * 3 - 17).createCell(0).setCellStyle(csSummaryTitleLeft);
		for (int i = 1; i < 11; i++) {
			SH1.getRow(nRow1 * 3 - 17).createCell(i).setCellStyle(csSummaryTitleCenter);
		}
		SH1.getRow(nRow1 * 3 - 17).createCell(11).setCellStyle(csSummaryTitleRight);

		for (int i = 0; i < 16; i++) {
			int row = nRow1 * 3 - 16 + i;
			SH1.createRow(row);
			if (i < 12) {
				SH1.addMergedRegion(new CellRangeAddress(row, row, 2, 3));
				SH1.addMergedRegion(new CellRangeAddress(row, row, 5, 6));
				SH1.addMergedRegion(new CellRangeAddress(row, row, 8, 9));
				SH1.addMergedRegion(new CellRangeAddress(row, row, 10, 11));
			} else if (i < 14) {
				SH1.addMergedRegion(new CellRangeAddress(row, row, 2, 11));
			} else {
				SH1.addMergedRegion(new CellRangeAddress(row, row, 0, 11));
			}
		}

		for (int k = 0; k < 16; k++) {
			if (k == 11) {
				SH1.getRow(nRow1 * 3 - 16 + k).createCell(0).setCellStyle(csNormalYellowLeft);
				for (int i = 1; i < 11; i++) {
					SH1.getRow(nRow1 * 3 - 16 + k).createCell(i).setCellStyle(csNormalYellowCenter);
				}
				SH1.getRow(nRow1 * 3 - 16 + k).createCell(11).setCellStyle(csNormalYellowRight);
			} else if (k == 0 || k == 14 || k == 15) {
				SH1.getRow(nRow1 * 3 - 16 + k).createCell(0).setCellStyle(csSummaryTitleLeft);
				for (int i = 1; i < 11; i++) {
					SH1.getRow(nRow1 * 3 - 16 + k).createCell(i).setCellStyle(csSummaryTitleCenter);
				}
				SH1.getRow(nRow1 * 3 - 16 + k).createCell(11).setCellStyle(csSummaryTitleRight);
			} else {
				SH1.getRow(nRow1 * 3 - 16 + k).createCell(0).setCellStyle(csNormalLeft);
				for (int i = 1; i < 11; i++) {
					SH1.getRow(nRow1 * 3 - 16 + k).createCell(i).setCellStyle(csNormalCenter);
				}
				SH1.getRow(nRow1 * 3 - 16 + k).createCell(11).setCellStyle(csNormalRight);
			}
		}

		SH1.getRow(nRow1 * 3 - 16).getCell(0).setCellValue("구분");
		SH1.getRow(nRow1 * 3 - 16).getCell(1).setCellValue("직분");
		SH1.getRow(nRow1 * 3 - 16).getCell(2).setCellValue("참석");
		SH1.getRow(nRow1 * 3 - 16).getCell(4).setCellValue("헌신");
		SH1.getRow(nRow1 * 3 - 16).getCell(5).setCellValue("불참");
		SH1.getRow(nRow1 * 3 - 16).getCell(7).setCellValue("전도 참석율");
		SH1.getRow(nRow1 * 3 - 16).getCell(8).setCellValue("헌신 참석율");
		SH1.getRow(nRow1 * 3 - 16).getCell(10).setCellValue("총 참석율");

		SH1.getRow(nRow1 * 3 - 15).getCell(0).setCellValue("1");
		SH1.getRow(nRow1 * 3 - 15).getCell(1).setCellValue("교역자");
		SH1.getRow(nRow1 * 3 - 15).getCell(2).setCellValue(String.valueOf(volume[0][2]));
		SH1.getRow(nRow1 * 3 - 15).getCell(4).setCellValue(String.valueOf(volume[0][1]));
		SH1.getRow(nRow1 * 3 - 15).getCell(5).setCellValue(String.valueOf(volume[0][0]));
		SH1.getRow(nRow1 * 3 - 15).getCell(7).setCellValue((int) ((float) volume[0][2] / (float) N[0] * 100) + "%");
		SH1.getRow(nRow1 * 3 - 15).getCell(8).setCellValue((int) ((float) volume[0][1] / (float) N[0] * 100) + "%");
		SH1.getRow(nRow1 * 3 - 15).getCell(10)
				.setCellValue((int) (((float) volume[0][1] + (float) volume[0][2]) / (float) N[0] * 100) + "%");

		SH1.getRow(nRow1 * 3 - 14).getCell(0).setCellValue("2");
		SH1.getRow(nRow1 * 3 - 14).getCell(1).setCellValue("장로");
		SH1.getRow(nRow1 * 3 - 14).getCell(2).setCellValue(String.valueOf(volume[1][2]));
		SH1.getRow(nRow1 * 3 - 14).getCell(4).setCellValue(String.valueOf(volume[1][1]));
		SH1.getRow(nRow1 * 3 - 14).getCell(5).setCellValue(String.valueOf(volume[1][0]));
		SH1.getRow(nRow1 * 3 - 14).getCell(7).setCellValue((int) ((float) volume[1][2] / (float) N[1] * 100) + "%");
		SH1.getRow(nRow1 * 3 - 14).getCell(8).setCellValue((int) ((float) volume[1][1] / (float) N[1] * 100) + "%");
		SH1.getRow(nRow1 * 3 - 14).getCell(10)
				.setCellValue((int) (((float) volume[1][1] + (float) volume[1][2]) / (float) N[1] * 100) + "%");

		SH1.getRow(nRow1 * 3 - 13).getCell(0).setCellValue("3");
		SH1.getRow(nRow1 * 3 - 13).getCell(1).setCellValue("안수집사");
		SH1.getRow(nRow1 * 3 - 13).getCell(2).setCellValue(String.valueOf(volume[2][2]));
		SH1.getRow(nRow1 * 3 - 13).getCell(4).setCellValue(String.valueOf(volume[2][1]));
		SH1.getRow(nRow1 * 3 - 13).getCell(5).setCellValue(String.valueOf(volume[2][0]));
		SH1.getRow(nRow1 * 3 - 13).getCell(7).setCellValue((int) ((float) volume[2][2] / (float) N[2] * 100) + "%");
		SH1.getRow(nRow1 * 3 - 13).getCell(8).setCellValue((int) ((float) volume[2][1] / (float) N[2] * 100) + "%");
		SH1.getRow(nRow1 * 3 - 13).getCell(10)
				.setCellValue((int) (((float) volume[2][1] + (float) volume[2][2]) / (float) N[2] * 100) + "%");

		SH1.getRow(nRow1 * 3 - 12).getCell(0).setCellValue("4");
		SH1.getRow(nRow1 * 3 - 12).getCell(1).setCellValue("권사");
		SH1.getRow(nRow1 * 3 - 12).getCell(2).setCellValue(String.valueOf(volume[3][2]));
		SH1.getRow(nRow1 * 3 - 12).getCell(4).setCellValue(String.valueOf(volume[3][1]));
		SH1.getRow(nRow1 * 3 - 12).getCell(5).setCellValue(String.valueOf(volume[3][0]));
		SH1.getRow(nRow1 * 3 - 12).getCell(7).setCellValue((int) ((float) volume[3][2] / (float) N[3] * 100) + "%");
		SH1.getRow(nRow1 * 3 - 12).getCell(8).setCellValue((int) ((float) volume[3][1] / (float) N[3] * 100) + "%");
		SH1.getRow(nRow1 * 3 - 12).getCell(10)
				.setCellValue((int) (((float) volume[3][1] + (float) volume[3][2]) / (float) N[3] * 100) + "%");

		SH1.getRow(nRow1 * 3 - 11).getCell(0).setCellValue("5");
		SH1.getRow(nRow1 * 3 - 11).getCell(1).setCellValue("서리집사(남)");
		SH1.getRow(nRow1 * 3 - 11).getCell(2).setCellValue(String.valueOf(volume[4][2]));
		SH1.getRow(nRow1 * 3 - 11).getCell(4).setCellValue(String.valueOf(volume[4][1]));
		SH1.getRow(nRow1 * 3 - 11).getCell(5).setCellValue(String.valueOf(volume[4][0]));
		SH1.getRow(nRow1 * 3 - 11).getCell(7).setCellValue((int) ((float) volume[4][2] / (float) N[4] * 100) + "%");
		SH1.getRow(nRow1 * 3 - 11).getCell(8).setCellValue((int) ((float) volume[4][1] / (float) N[4] * 100) + "%");
		SH1.getRow(nRow1 * 3 - 11).getCell(10)
				.setCellValue((int) (((float) volume[4][1] + (float) volume[4][2]) / (float) N[4] * 100) + "%");

		SH1.getRow(nRow1 * 3 - 10).getCell(0).setCellValue("6");
		SH1.getRow(nRow1 * 3 - 10).getCell(1).setCellValue("서리집사(여)");
		SH1.getRow(nRow1 * 3 - 10).getCell(2).setCellValue(String.valueOf(volume[5][2]));
		SH1.getRow(nRow1 * 3 - 10).getCell(4).setCellValue(String.valueOf(volume[5][1]));
		SH1.getRow(nRow1 * 3 - 10).getCell(5).setCellValue(String.valueOf(volume[5][0]));
		SH1.getRow(nRow1 * 3 - 10).getCell(7).setCellValue((int) ((float) volume[5][2] / (float) N[5] * 100) + "%");
		SH1.getRow(nRow1 * 3 - 10).getCell(8).setCellValue((int) ((float) volume[5][1] / (float) N[5] * 100) + "%");
		SH1.getRow(nRow1 * 3 - 10).getCell(10)
				.setCellValue((int) (((float) volume[5][1] + (float) volume[5][2]) / (float) N[5] * 100) + "%");

		SH1.getRow(nRow1 * 3 - 9).getCell(0).setCellValue("7");
		SH1.getRow(nRow1 * 3 - 9).getCell(1).setCellValue("권찰");
		SH1.getRow(nRow1 * 3 - 9).getCell(2).setCellValue(String.valueOf(volume[6][2]));
		SH1.getRow(nRow1 * 3 - 9).getCell(4).setCellValue(String.valueOf(volume[6][1]));
		SH1.getRow(nRow1 * 3 - 9).getCell(5).setCellValue(String.valueOf(volume[6][0]));
		SH1.getRow(nRow1 * 3 - 9).getCell(7).setCellValue((int) ((float) volume[6][2] / (float) N[6] * 100) + "%");
		SH1.getRow(nRow1 * 3 - 9).getCell(8).setCellValue((int) ((float) volume[6][1] / (float) N[6] * 100) + "%");
		SH1.getRow(nRow1 * 3 - 9).getCell(10)
				.setCellValue((int) (((float) volume[6][1] + (float) volume[6][2]) / (float) N[6] * 100) + "%");

		SH1.getRow(nRow1 * 3 - 8).getCell(0).setCellValue("8");
		SH1.getRow(nRow1 * 3 - 8).getCell(1).setCellValue("성도");
		SH1.getRow(nRow1 * 3 - 8).getCell(2).setCellValue(String.valueOf(volume[7][2] + volume[8][2]));
		SH1.getRow(nRow1 * 3 - 8).getCell(4).setCellValue(String.valueOf(volume[7][1] + volume[8][1]));
		SH1.getRow(nRow1 * 3 - 8).getCell(5).setCellValue(String.valueOf(volume[7][0] + volume[8][0]));
		SH1.getRow(nRow1 * 3 - 8).getCell(7).setCellValue(
				(int) (((float) volume[7][2] + (float) volume[8][2]) / ((float) N[7] + (float) N[8]) * 100) + "%");
		SH1.getRow(nRow1 * 3 - 8).getCell(8).setCellValue(
				(int) (((float) volume[7][1] + (float) volume[8][1]) / ((float) N[7] + (float) N[8]) * 100) + "%");
		SH1.getRow(nRow1 * 3 - 8).getCell(10).setCellValue(
				(int) (((float) volume[7][1] + (float) volume[7][2] + (float) volume[8][1] + (float) volume[8][2])
						/ ((float) N[7] + (float) N[8]) * 100) + "%");

		SH1.getRow(nRow1 * 3 - 7).getCell(0).setCellValue("9");
		SH1.getRow(nRow1 * 3 - 7).getCell(1).setCellValue("청년");
		SH1.getRow(nRow1 * 3 - 7).getCell(2).setCellValue(String.valueOf(volume[9][2]));
		SH1.getRow(nRow1 * 3 - 7).getCell(4).setCellValue(String.valueOf(volume[9][1]));
		SH1.getRow(nRow1 * 3 - 7).getCell(5).setCellValue(String.valueOf(volume[9][0]));
		SH1.getRow(nRow1 * 3 - 7).getCell(7).setCellValue((int) ((float) volume[9][2] / (float) N[9] * 100) + "%");
		SH1.getRow(nRow1 * 3 - 7).getCell(8).setCellValue((int) ((float) volume[9][1] / (float) N[9] * 100) + "%");
		SH1.getRow(nRow1 * 3 - 7).getCell(10)
				.setCellValue((int) (((float) volume[9][1] + (float) volume[9][2]) / (float) N[9] * 100) + "%");

		SH1.getRow(nRow1 * 3 - 6).getCell(0).setCellValue("10");
		SH1.getRow(nRow1 * 3 - 6).getCell(1).setCellValue("대학");
		SH1.getRow(nRow1 * 3 - 6).getCell(2).setCellValue(String.valueOf(volume[10][2]));
		SH1.getRow(nRow1 * 3 - 6).getCell(4).setCellValue(String.valueOf(volume[10][1]));
		SH1.getRow(nRow1 * 3 - 6).getCell(5).setCellValue(String.valueOf(volume[10][0]));
		SH1.getRow(nRow1 * 3 - 6).getCell(7).setCellValue((int) ((float) volume[10][2] / (float) N[10] * 100) + "%");
		SH1.getRow(nRow1 * 3 - 6).getCell(8).setCellValue((int) ((float) volume[10][1] / (float) N[10] * 100) + "%");
		SH1.getRow(nRow1 * 3 - 6).getCell(10)
				.setCellValue((int) (((float) volume[10][1] + (float) volume[10][2]) / (float) N[10] * 100) + "%");

		int s = 0;
		int s1 = 0;
		int s2 = 0;
		int s3 = 0;
		for (int i = 0; i < 11; i++) {
			s += N[i];
			s1 += volume[i][2];
			s2 += volume[i][1];
			s3 += volume[i][0];
		}
		SH1.getRow(nRow1 * 3 - 5).getCell(0).setCellValue("11");
		SH1.getRow(nRow1 * 3 - 5).getCell(1).setCellValue("총계");
		SH1.getRow(nRow1 * 3 - 5).getCell(2).setCellValue(String.valueOf(s1));
		SH1.getRow(nRow1 * 3 - 5).getCell(4).setCellValue(String.valueOf(s2));
		SH1.getRow(nRow1 * 3 - 5).getCell(5).setCellValue(String.valueOf(s3));
		SH1.getRow(nRow1 * 3 - 5).getCell(7).setCellValue((int) ((float) s1 / (float) s * 100) + "%");
		SH1.getRow(nRow1 * 3 - 5).getCell(8).setCellValue((int) ((float) s2 / (float) s * 100) + "%");
		SH1.getRow(nRow1 * 3 - 5).getCell(10).setCellValue((int) (((float) s1 + (float) s2) / (float) s * 100) + "%");

		int s4 = 0;
		for (int i = 12; i < 24; i++) {
			s4 += volume[i][2];
		}
		s4 += people2.size();

		SH1.getRow(nRow1 * 3 - 4).getCell(0).setCellValue("12");
		SH1.getRow(nRow1 * 3 - 4).getCell(1).setCellValue("지교회");
		SH1.getRow(nRow1 * 3 - 4).getCell(2).setCellValue(String.valueOf(s4));

		SH1.getRow(nRow1 * 3 - 3).getCell(0).setCellValue("13");
		SH1.getRow(nRow1 * 3 - 3).getCell(1).setCellValue("기타");
		SH1.getRow(nRow1 * 3 - 3).getCell(2).setCellValue(String.valueOf(people1.size()));

		int s5 = s1 + s4 + people1.size();

		SH1.getRow(nRow1 * 3 - 2).getCell(0).setCellValue("총: " + s5 + "명");

		SH1.getRow(nRow1 * 3 - 1).getCell(0).setCellValue("* 기타 = 유치, 유년, 초등, 중등, 고등, 성도 (참석숫자만)");
	}

	public void setCellStyle2() {
		csTitle2 = WB2.createCellStyle();
		csTitle2.setAlignment(HorizontalAlignment.CENTER);
		csTitle2.setVerticalAlignment(VerticalAlignment.TOP);
		csTitle2.setFont(fTitle2);

		csIdxBc = WB2.createCellStyle();
		csIdxBc.setAlignment(HorizontalAlignment.CENTER);
		csIdxBc.setVerticalAlignment(VerticalAlignment.CENTER);
		csIdxBc.setBorderTop(BorderStyle.THIN);
		csIdxBc.setTopBorderColor(IndexedColors.BLACK.getIndex());
		csIdxBc.setBorderBottom(BorderStyle.THIN);
		csIdxBc.setBottomBorderColor(IndexedColors.BLACK.getIndex());
		csIdxBc.setBorderLeft(BorderStyle.THIN);
		csIdxBc.setLeftBorderColor(IndexedColors.BLACK.getIndex());
		csIdxBc.setBorderRight(BorderStyle.THIN);
		csIdxBc.setRightBorderColor(IndexedColors.BLACK.getIndex());
		csIdxBc.setFillForegroundColor(cDedication);
		csIdxBc.setFillPattern(FillPatternType.SOLID_FOREGROUND);
		csIdxBc.setFont(fNormal2);

		csIdxRemeet = WB2.createCellStyle();
		csIdxRemeet.setAlignment(HorizontalAlignment.CENTER);
		csIdxRemeet.setVerticalAlignment(VerticalAlignment.CENTER);
		csIdxRemeet.setBorderTop(BorderStyle.THIN);
		csIdxRemeet.setTopBorderColor(IndexedColors.BLACK.getIndex());
		csIdxRemeet.setBorderBottom(BorderStyle.THIN);
		csIdxRemeet.setBottomBorderColor(IndexedColors.BLACK.getIndex());
		csIdxRemeet.setBorderLeft(BorderStyle.THIN);
		csIdxRemeet.setLeftBorderColor(IndexedColors.BLACK.getIndex());
		csIdxRemeet.setBorderRight(BorderStyle.THIN);
		csIdxRemeet.setRightBorderColor(IndexedColors.BLACK.getIndex());
		csIdxRemeet.setFillForegroundColor(cAttendance);
		csIdxRemeet.setFillPattern(FillPatternType.SOLID_FOREGROUND);
		csIdxRemeet.setFont(fNormal2);

		csIdxBeliever = WB2.createCellStyle();
		csIdxBeliever.setAlignment(HorizontalAlignment.CENTER);
		csIdxBeliever.setVerticalAlignment(VerticalAlignment.CENTER);
		csIdxBeliever.setBorderTop(BorderStyle.THIN);
		csIdxBeliever.setTopBorderColor(IndexedColors.BLACK.getIndex());
		csIdxBeliever.setBorderBottom(BorderStyle.THIN);
		csIdxBeliever.setBottomBorderColor(IndexedColors.BLACK.getIndex());
		csIdxBeliever.setBorderLeft(BorderStyle.THIN);
		csIdxBeliever.setLeftBorderColor(IndexedColors.BLACK.getIndex());
		csIdxBeliever.setBorderRight(BorderStyle.THIN);
		csIdxBeliever.setRightBorderColor(IndexedColors.BLACK.getIndex());
		csIdxBeliever.setFillForegroundColor(IndexedColors.WHITE.getIndex());
		csIdxBeliever.setFillPattern(FillPatternType.SOLID_FOREGROUND);
		csIdxBeliever.setFont(fNormal2);

		csMenuAboveLeft = WB2.createCellStyle();
		csMenuAboveLeft.setAlignment(HorizontalAlignment.CENTER);
		csMenuAboveLeft.setVerticalAlignment(VerticalAlignment.CENTER);
		csMenuAboveLeft.setBorderTop(BorderStyle.THICK);
		csMenuAboveLeft.setTopBorderColor(IndexedColors.BLACK.getIndex());
		csMenuAboveLeft.setBorderBottom(BorderStyle.THIN);
		csMenuAboveLeft.setBottomBorderColor(IndexedColors.BLACK.getIndex());
		csMenuAboveLeft.setBorderLeft(BorderStyle.THICK);
		csMenuAboveLeft.setLeftBorderColor(IndexedColors.BLACK.getIndex());
		csMenuAboveLeft.setFillForegroundColor(cCategory);
		csMenuAboveLeft.setFillPattern(FillPatternType.SOLID_FOREGROUND);
		csMenuAboveLeft.setFont(fNormal2);

		csMenuAboveCenter = WB2.createCellStyle();
		csMenuAboveCenter.setAlignment(HorizontalAlignment.CENTER);
		csMenuAboveCenter.setVerticalAlignment(VerticalAlignment.CENTER);
		csMenuAboveCenter.setBorderTop(BorderStyle.THICK);
		csMenuAboveCenter.setTopBorderColor(IndexedColors.BLACK.getIndex());
		csMenuAboveCenter.setBorderBottom(BorderStyle.THIN);
		csMenuAboveCenter.setBottomBorderColor(IndexedColors.BLACK.getIndex());
		csMenuAboveCenter.setFillForegroundColor(cCategory);
		csMenuAboveCenter.setFillPattern(FillPatternType.SOLID_FOREGROUND);
		csMenuAboveCenter.setFont(fNormal2);

		csMenuAboveRight = WB2.createCellStyle();
		csMenuAboveRight.setAlignment(HorizontalAlignment.CENTER);
		csMenuAboveRight.setVerticalAlignment(VerticalAlignment.CENTER);
		csMenuAboveRight.setBorderTop(BorderStyle.THICK);
		csMenuAboveRight.setTopBorderColor(IndexedColors.BLACK.getIndex());
		csMenuAboveRight.setBorderBottom(BorderStyle.THIN);
		csMenuAboveRight.setBottomBorderColor(IndexedColors.BLACK.getIndex());
		csMenuAboveRight.setBorderLeft(BorderStyle.THIN);
		csMenuAboveRight.setLeftBorderColor(IndexedColors.BLACK.getIndex());
		csMenuAboveRight.setBorderRight(BorderStyle.THICK);
		csMenuAboveRight.setRightBorderColor(IndexedColors.BLACK.getIndex());
		csMenuAboveRight.setFillForegroundColor(cCategory);
		csMenuAboveRight.setFillPattern(FillPatternType.SOLID_FOREGROUND);
		csMenuAboveRight.setFont(fNormal2);

		csMenuBelowLeft = WB2.createCellStyle();
		csMenuBelowLeft.setAlignment(HorizontalAlignment.CENTER);
		csMenuBelowLeft.setVerticalAlignment(VerticalAlignment.CENTER);
		csMenuBelowLeft.setBorderTop(BorderStyle.THIN);
		csMenuBelowLeft.setTopBorderColor(IndexedColors.BLACK.getIndex());
		csMenuBelowLeft.setBorderBottom(BorderStyle.THICK);
		csMenuBelowLeft.setBottomBorderColor(IndexedColors.BLACK.getIndex());
		csMenuBelowLeft.setBorderLeft(BorderStyle.THICK);
		csMenuBelowLeft.setLeftBorderColor(IndexedColors.BLACK.getIndex());
		csMenuBelowLeft.setBorderRight(BorderStyle.THIN);
		csMenuBelowLeft.setRightBorderColor(IndexedColors.BLACK.getIndex());
		csMenuBelowLeft.setFillForegroundColor(cCategory);
		csMenuBelowLeft.setFillPattern(FillPatternType.SOLID_FOREGROUND);
		csMenuBelowLeft.setFont(fNormal2);

		csMenuBelowCenter = WB2.createCellStyle();
		csMenuBelowCenter.setAlignment(HorizontalAlignment.CENTER);
		csMenuBelowCenter.setVerticalAlignment(VerticalAlignment.CENTER);
		csMenuBelowCenter.setBorderTop(BorderStyle.THIN);
		csMenuBelowCenter.setTopBorderColor(IndexedColors.BLACK.getIndex());
		csMenuBelowCenter.setBorderBottom(BorderStyle.THICK);
		csMenuBelowCenter.setBottomBorderColor(IndexedColors.BLACK.getIndex());
		csMenuBelowCenter.setBorderLeft(BorderStyle.THIN);
		csMenuBelowCenter.setLeftBorderColor(IndexedColors.BLACK.getIndex());
		csMenuBelowCenter.setBorderRight(BorderStyle.THIN);
		csMenuBelowCenter.setRightBorderColor(IndexedColors.BLACK.getIndex());
		csMenuBelowCenter.setFillForegroundColor(cCategory);
		csMenuBelowCenter.setFillPattern(FillPatternType.SOLID_FOREGROUND);
		csMenuBelowCenter.setFont(fNormal2);

		csMenuBelowRight = WB2.createCellStyle();
		csMenuBelowRight.setAlignment(HorizontalAlignment.CENTER);
		csMenuBelowRight.setVerticalAlignment(VerticalAlignment.CENTER);
		csMenuBelowRight.setBorderTop(BorderStyle.THIN);
		csMenuBelowRight.setTopBorderColor(IndexedColors.BLACK.getIndex());
		csMenuBelowRight.setBorderBottom(BorderStyle.THICK);
		csMenuBelowRight.setBottomBorderColor(IndexedColors.BLACK.getIndex());
		csMenuBelowRight.setBorderLeft(BorderStyle.THIN);
		csMenuBelowRight.setLeftBorderColor(IndexedColors.BLACK.getIndex());
		csMenuBelowRight.setBorderRight(BorderStyle.THICK);
		csMenuBelowRight.setRightBorderColor(IndexedColors.BLACK.getIndex());
		csMenuBelowRight.setFillForegroundColor(cCategory);
		csMenuBelowRight.setFillPattern(FillPatternType.SOLID_FOREGROUND);
		csMenuBelowRight.setFont(fNormal2);

		csMenuNormalLeft = WB2.createCellStyle();
		csMenuNormalLeft.setAlignment(HorizontalAlignment.CENTER);
		csMenuNormalLeft.setVerticalAlignment(VerticalAlignment.CENTER);
		csMenuNormalLeft.setBorderTop(BorderStyle.THIN);
		csMenuNormalLeft.setTopBorderColor(IndexedColors.BLACK.getIndex());
		csMenuNormalLeft.setBorderBottom(BorderStyle.THIN);
		csMenuNormalLeft.setBottomBorderColor(IndexedColors.BLACK.getIndex());
		csMenuNormalLeft.setBorderLeft(BorderStyle.THICK);
		csMenuNormalLeft.setLeftBorderColor(IndexedColors.BLACK.getIndex());
		csMenuNormalLeft.setBorderRight(BorderStyle.THIN);
		csMenuNormalLeft.setRightBorderColor(IndexedColors.BLACK.getIndex());
		csMenuNormalLeft.setFillForegroundColor(cCategory);
		csMenuNormalLeft.setFillPattern(FillPatternType.SOLID_FOREGROUND);
		csMenuNormalLeft.setFont(fNormal2);

		csMenuNormalCenter = WB2.createCellStyle();
		csMenuNormalCenter.setAlignment(HorizontalAlignment.CENTER);
		csMenuNormalCenter.setVerticalAlignment(VerticalAlignment.CENTER);
		csMenuNormalCenter.setBorderTop(BorderStyle.THIN);
		csMenuNormalCenter.setTopBorderColor(IndexedColors.BLACK.getIndex());
		csMenuNormalCenter.setBorderBottom(BorderStyle.THIN);
		csMenuNormalCenter.setBottomBorderColor(IndexedColors.BLACK.getIndex());
		csMenuNormalCenter.setBorderLeft(BorderStyle.THIN);
		csMenuNormalCenter.setLeftBorderColor(IndexedColors.BLACK.getIndex());
		csMenuNormalCenter.setBorderRight(BorderStyle.THIN);
		csMenuNormalCenter.setRightBorderColor(IndexedColors.BLACK.getIndex());
		csMenuNormalCenter.setFillForegroundColor(cCategory);
		csMenuNormalCenter.setFillPattern(FillPatternType.SOLID_FOREGROUND);
		csMenuNormalCenter.setFont(fNormal2);

		csMenuNormalRight = WB2.createCellStyle();
		csMenuNormalRight.setAlignment(HorizontalAlignment.CENTER);
		csMenuNormalRight.setVerticalAlignment(VerticalAlignment.CENTER);
		csMenuNormalRight.setBorderTop(BorderStyle.THIN);
		csMenuNormalRight.setTopBorderColor(IndexedColors.BLACK.getIndex());
		csMenuNormalRight.setBorderBottom(BorderStyle.THIN);
		csMenuNormalRight.setBottomBorderColor(IndexedColors.BLACK.getIndex());
		csMenuNormalRight.setBorderLeft(BorderStyle.THIN);
		csMenuNormalRight.setLeftBorderColor(IndexedColors.BLACK.getIndex());
		csMenuNormalRight.setBorderRight(BorderStyle.THICK);
		csMenuNormalRight.setRightBorderColor(IndexedColors.BLACK.getIndex());
		csMenuNormalRight.setFillForegroundColor(cCategory);
		csMenuNormalRight.setFillPattern(FillPatternType.SOLID_FOREGROUND);
		csMenuNormalRight.setFont(fNormal2);

		csNormalLeft = WB2.createCellStyle();
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
		csNormalLeft.setFont(fNormal2);

		csNormalCenter = WB2.createCellStyle();
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
		csNormalCenter.setFont(fNormal2);

		csNormalRight = WB2.createCellStyle();
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
		csNormalRight.setFont(fNormal2);

		csBC = WB2.createCellStyle();
		csBC.setAlignment(HorizontalAlignment.CENTER);
		csBC.setVerticalAlignment(VerticalAlignment.CENTER);
		csBC.setBorderTop(BorderStyle.NONE);
		csBC.setTopBorderColor(IndexedColors.BLACK.getIndex());
		csBC.setBorderBottom(BorderStyle.NONE);
		csBC.setBottomBorderColor(IndexedColors.BLACK.getIndex());
		csBC.setBorderLeft(BorderStyle.NONE);
		csBC.setLeftBorderColor(IndexedColors.BLACK.getIndex());
		csBC.setBorderRight(BorderStyle.NONE);
		csBC.setRightBorderColor(IndexedColors.BLACK.getIndex());
		csBC.setFillForegroundColor(cDedication);
		csBC.setFillPattern(FillPatternType.SOLID_FOREGROUND);
		csBC.setFont(fNormal2);

		csAT = WB2.createCellStyle();
		csAT.setAlignment(HorizontalAlignment.CENTER);
		csAT.setVerticalAlignment(VerticalAlignment.CENTER);
		csAT.setBorderTop(BorderStyle.NONE);
		csAT.setTopBorderColor(IndexedColors.BLACK.getIndex());
		csAT.setBorderBottom(BorderStyle.NONE);
		csAT.setBottomBorderColor(IndexedColors.BLACK.getIndex());
		csAT.setBorderLeft(BorderStyle.NONE);
		csAT.setLeftBorderColor(IndexedColors.BLACK.getIndex());
		csAT.setBorderRight(BorderStyle.NONE);
		csAT.setRightBorderColor(IndexedColors.BLACK.getIndex());
		csAT.setFillForegroundColor(IndexedColors.WHITE.getIndex());
		csAT.setFillPattern(FillPatternType.SOLID_FOREGROUND);
		csAT.setFont(fNormal2);

		csBCf = WB2.createCellStyle();
		csBCf.setAlignment(HorizontalAlignment.CENTER);
		csBCf.setVerticalAlignment(VerticalAlignment.CENTER);
		csBCf.setBorderTop(BorderStyle.THIN);
		csBCf.setTopBorderColor(IndexedColors.BLACK.getIndex());
		csBCf.setBorderBottom(BorderStyle.NONE);
		csBCf.setBottomBorderColor(IndexedColors.BLACK.getIndex());
		csBCf.setBorderLeft(BorderStyle.NONE);
		csBCf.setLeftBorderColor(IndexedColors.BLACK.getIndex());
		csBCf.setBorderRight(BorderStyle.NONE);
		csBCf.setRightBorderColor(IndexedColors.BLACK.getIndex());
		csBCf.setFillForegroundColor(cDedication);
		csBCf.setFillPattern(FillPatternType.SOLID_FOREGROUND);
		csBCf.setFont(fNormal2);

		csATf = WB2.createCellStyle();
		csATf.setAlignment(HorizontalAlignment.CENTER);
		csATf.setVerticalAlignment(VerticalAlignment.CENTER);
		csATf.setBorderTop(BorderStyle.THIN);
		csATf.setTopBorderColor(IndexedColors.BLACK.getIndex());
		csATf.setBorderBottom(BorderStyle.NONE);
		csATf.setBottomBorderColor(IndexedColors.BLACK.getIndex());
		csATf.setBorderLeft(BorderStyle.NONE);
		csATf.setLeftBorderColor(IndexedColors.BLACK.getIndex());
		csATf.setBorderRight(BorderStyle.NONE);
		csATf.setRightBorderColor(IndexedColors.BLACK.getIndex());
		csATf.setFillForegroundColor(IndexedColors.WHITE.getIndex());
		csATf.setFillPattern(FillPatternType.SOLID_FOREGROUND);
		csATf.setFont(fNormal2);

		csNormalAboveLeft = WB2.createCellStyle();
		csNormalAboveLeft.setAlignment(HorizontalAlignment.CENTER);
		csNormalAboveLeft.setVerticalAlignment(VerticalAlignment.CENTER);
		csNormalAboveLeft.setBorderTop(BorderStyle.THICK);
		csNormalAboveLeft.setTopBorderColor(IndexedColors.BLACK.getIndex());
		csNormalAboveLeft.setBorderBottom(BorderStyle.THIN);
		csNormalAboveLeft.setBottomBorderColor(IndexedColors.BLACK.getIndex());
		csNormalAboveLeft.setBorderLeft(BorderStyle.THICK);
		csNormalAboveLeft.setLeftBorderColor(IndexedColors.BLACK.getIndex());
		csNormalAboveLeft.setBorderRight(BorderStyle.THIN);
		csNormalAboveLeft.setRightBorderColor(IndexedColors.BLACK.getIndex());
		csNormalAboveLeft.setFillForegroundColor(cNormal);
		csNormalAboveLeft.setFillPattern(FillPatternType.SOLID_FOREGROUND);
		csNormalAboveLeft.setFont(fNormal2);

		csNormalAboveCenter = WB2.createCellStyle();
		csNormalAboveCenter.setAlignment(HorizontalAlignment.CENTER);
		csNormalAboveCenter.setVerticalAlignment(VerticalAlignment.CENTER);
		csNormalAboveCenter.setBorderTop(BorderStyle.THICK);
		csNormalAboveCenter.setTopBorderColor(IndexedColors.BLACK.getIndex());
		csNormalAboveCenter.setBorderBottom(BorderStyle.THIN);
		csNormalAboveCenter.setBottomBorderColor(IndexedColors.BLACK.getIndex());
		csNormalAboveCenter.setBorderLeft(BorderStyle.THIN);
		csNormalAboveCenter.setLeftBorderColor(IndexedColors.BLACK.getIndex());
		csNormalAboveCenter.setBorderRight(BorderStyle.THIN);
		csNormalAboveCenter.setRightBorderColor(IndexedColors.BLACK.getIndex());
		csNormalAboveCenter.setFillForegroundColor(cNormal);
		csNormalAboveCenter.setFillPattern(FillPatternType.SOLID_FOREGROUND);
		csNormalAboveCenter.setFont(fNormal2);

		csNormalAboveRight = WB2.createCellStyle();
		csNormalAboveRight.setAlignment(HorizontalAlignment.CENTER);
		csNormalAboveRight.setVerticalAlignment(VerticalAlignment.CENTER);
		csNormalAboveRight.setBorderTop(BorderStyle.THICK);
		csNormalAboveRight.setTopBorderColor(IndexedColors.BLACK.getIndex());
		csNormalAboveRight.setBorderBottom(BorderStyle.THIN);
		csNormalAboveRight.setBottomBorderColor(IndexedColors.BLACK.getIndex());
		csNormalAboveRight.setBorderLeft(BorderStyle.THIN);
		csNormalAboveRight.setLeftBorderColor(IndexedColors.BLACK.getIndex());
		csNormalAboveRight.setBorderRight(BorderStyle.THICK);
		csNormalAboveRight.setRightBorderColor(IndexedColors.BLACK.getIndex());
		csNormalAboveRight.setFillForegroundColor(cNormal);
		csNormalAboveRight.setFillPattern(FillPatternType.SOLID_FOREGROUND);
		csNormalAboveRight.setFont(fNormal2);

		csRemeet = WB2.createCellStyle();
		csRemeet.setAlignment(HorizontalAlignment.CENTER);
		csRemeet.setVerticalAlignment(VerticalAlignment.CENTER);
		csRemeet.setBorderTop(BorderStyle.THIN);
		csRemeet.setTopBorderColor(IndexedColors.BLACK.getIndex());
		csRemeet.setBorderBottom(BorderStyle.THIN);
		csRemeet.setBottomBorderColor(IndexedColors.BLACK.getIndex());
		csRemeet.setBorderLeft(BorderStyle.THIN);
		csRemeet.setLeftBorderColor(IndexedColors.BLACK.getIndex());
		csRemeet.setBorderRight(BorderStyle.THIN);
		csRemeet.setRightBorderColor(IndexedColors.BLACK.getIndex());
		csRemeet.setFillForegroundColor(cAttendance);
		csRemeet.setFillPattern(FillPatternType.SOLID_FOREGROUND);
		csRemeet.setFont(fNormal2);

		csRemeetAbove = WB2.createCellStyle();
		csRemeetAbove.setAlignment(HorizontalAlignment.CENTER);
		csRemeetAbove.setVerticalAlignment(VerticalAlignment.CENTER);
		csRemeetAbove.setBorderTop(BorderStyle.THICK);
		csRemeetAbove.setTopBorderColor(IndexedColors.BLACK.getIndex());
		csRemeetAbove.setBorderBottom(BorderStyle.THIN);
		csRemeetAbove.setBottomBorderColor(IndexedColors.BLACK.getIndex());
		csRemeetAbove.setBorderLeft(BorderStyle.THIN);
		csRemeetAbove.setLeftBorderColor(IndexedColors.BLACK.getIndex());
		csRemeetAbove.setBorderRight(BorderStyle.THIN);
		csRemeetAbove.setRightBorderColor(IndexedColors.BLACK.getIndex());
		csRemeetAbove.setFillForegroundColor(cAttendance);
		csRemeetAbove.setFillPattern(FillPatternType.SOLID_FOREGROUND);
		csRemeetAbove.setFont(fNormal2);

		csSum = WB2.createCellStyle();
		csSum.setAlignment(HorizontalAlignment.CENTER);
		csSum.setVerticalAlignment(VerticalAlignment.CENTER);
		csSum.setBorderTop(BorderStyle.THICK);
		csSum.setTopBorderColor(IndexedColors.BLACK.getIndex());
		csSum.setBorderBottom(BorderStyle.THICK);
		csSum.setBottomBorderColor(IndexedColors.BLACK.getIndex());
		csSum.setBorderLeft(BorderStyle.THICK);
		csSum.setLeftBorderColor(IndexedColors.BLACK.getIndex());
		csSum.setBorderRight(BorderStyle.THICK);
		csSum.setRightBorderColor(IndexedColors.BLACK.getIndex());
		csSum.setFillForegroundColor(IndexedColors.WHITE.getIndex());
		csSum.setFillPattern(FillPatternType.SOLID_FOREGROUND);
		csSum.setFont(fNormal2);

		csYellow = WB2.createCellStyle();
		csYellow.setAlignment(HorizontalAlignment.CENTER);
		csYellow.setVerticalAlignment(VerticalAlignment.CENTER);
		csYellow.setBorderTop(BorderStyle.THICK);
		csYellow.setTopBorderColor(IndexedColors.BLACK.getIndex());
		csYellow.setBorderBottom(BorderStyle.THICK);
		csYellow.setBottomBorderColor(IndexedColors.BLACK.getIndex());
		csYellow.setBorderLeft(BorderStyle.THICK);
		csYellow.setLeftBorderColor(IndexedColors.BLACK.getIndex());
		csYellow.setBorderRight(BorderStyle.THICK);
		csYellow.setRightBorderColor(IndexedColors.BLACK.getIndex());
		csYellow.setFillForegroundColor(cYellow);
		csYellow.setFillPattern(FillPatternType.SOLID_FOREGROUND);
		csYellow.setFont(fNormal2);

		csLine = WB2.createCellStyle();
		csLine.setAlignment(HorizontalAlignment.CENTER);
		csLine.setVerticalAlignment(VerticalAlignment.CENTER);
		csLine.setBorderTop(BorderStyle.NONE);
		csLine.setTopBorderColor(IndexedColors.BLACK.getIndex());
		csLine.setBorderBottom(BorderStyle.NONE);
		csLine.setBottomBorderColor(IndexedColors.BLACK.getIndex());
		csLine.setBorderLeft(BorderStyle.NONE);
		csLine.setLeftBorderColor(IndexedColors.BLACK.getIndex());
		csLine.setBorderRight(BorderStyle.THICK);
		csLine.setRightBorderColor(IndexedColors.BLACK.getIndex());
		csLine.setFillForegroundColor(cNormal);
		csLine.setFillPattern(FillPatternType.SOLID_FOREGROUND);
		csLine.setFont(fNormal2);

	}

	public void init2() {
		list = new String[10][2][];
		String[][][] list1 = new String[10][2][];
		String[][][] list2 = new String[10][2][];
		listNum = new int[10];
		int[] listNum1 = new int[10];
		int[] listNum2 = new int[10];
		for (int i = 0; i < listNum.length; i++) {
			listNum[i] = 0;
			listNum1[i] = 0;
			listNum2[i] = 0;
		}
		attNum = new int[8];
		for (int i = 0; i < attNum.length; i++) {
			attNum[i] = 0;
		}
		listNum[9] = people1.size() + people2.size();
		attNum[7] = people2.size();
		for (int i = 0; i < peopleIdx1.size(); i++) {
			attNum[peopleIdx1.get(i)]++;
		}
		int week = Calendar.getInstance().get(Calendar.WEEK_OF_YEAR);
		if (week == 1) {
			if (Calendar.getInstance().get(Calendar.MONTH) == 11) {
				week = 53;
			}
		}
		week += 3;
		for (int i = 1; i < num_attendance; i++) {
			String a = attendance.getRow(i).getCell(week).toString();
			if (a.equals("2.0")) {
				String b = attendance.getRow(i).getCell(3).toString();
				switch (b) {
				case "fp1":
					attNum[0]++;
					break;
				case "fp3":
					attNum[1]++;
					break;
				case "fp5":
					attNum[2]++;
					break;
				case "fp7":
					attNum[3]++;
					break;
				case "fp9":
					attNum[4]++;
					break;
				case "fp11":
					attNum[5]++;
					break;
				case "fp13":
					attNum[6]++;
					break;
				default:
					attNum[7]++;
					break;
				}
				String c = attendance.getRow(i).getCell(57).toString();
				for (int j = 0; j < category1.length; j++) {
					if (c.equals(category1[j])) {
						switch (j) {
						case 0:
							listNum[0]++;
							listNum1[0]++;
							break;
						case 1:
							listNum[1]++;
							listNum1[1]++;
							break;
						case 2:
							listNum[2]++;
							listNum1[2]++;
							break;
						case 3:
							listNum[3]++;
							listNum1[3]++;
							break;
						case 4:
							listNum[4]++;
							listNum1[4]++;
							break;
						case 5:
							listNum[5]++;
							listNum1[5]++;
							break;
						case 6:
							listNum[6]++;
							listNum1[6]++;
							break;
						case 7:
							listNum[6]++;
							listNum1[6]++;
							break;
						case 8:
							listNum[6]++;
							listNum1[6]++;
							break;
						case 9:
							listNum[7]++;
							listNum1[7]++;
							break;
						case 10:
							listNum[8]++;
							listNum1[8]++;
							break;
						default:
							String d = attendance.getRow(i).getCell(1).toString();
							switch (d) {
							case "교역자":
								listNum[0]++;
								listNum2[0]++;
								break;
							case "장로":
								listNum[1]++;
								listNum2[1]++;
								break;
							case "안수집사":
								listNum[2]++;
								listNum2[2]++;
								break;
							case "권사":
								listNum[3]++;
								listNum2[3]++;
								break;
							case "서리집사(남)":
								listNum[4]++;
								listNum2[4]++;
								break;
							case "서리집사(여)":
								listNum[5]++;
								listNum2[5]++;
								break;
							case "권찰":
								listNum[6]++;
								listNum2[6]++;
								break;
							case "성도(남)":
								listNum[6]++;
								listNum2[6]++;
								break;
							case "성도(여)":
								listNum[6]++;
								listNum2[6]++;
								break;
							case "청년":
								listNum[7]++;
								listNum2[7]++;
								break;
							case "대학":
								listNum[8]++;
								listNum2[8]++;
								break;
							}
							break;
						}

						break;
					}
				}
			}
		}

		for (int i = 0; i < list.length; i++) {
			if (listNum[i] != 0 && listNum[i] % 10 == 0) {
				list[i][0] = new String[listNum[i]];
				list[i][1] = new String[listNum[i]];
			} else {
				list[i][0] = new String[listNum[i] + 10 - (listNum[i] % 10)];
				list[i][1] = new String[listNum[i] + 10 - (listNum[i] % 10)];
			}
			list1[i][0] = new String[listNum1[i]];
			list1[i][1] = new String[listNum1[i]];
			list2[i][0] = new String[listNum2[i]];
			list2[i][1] = new String[listNum2[i]];
			if (i != 9) {
				listNum1[i] = 0;
				listNum2[i] = 0;

			}
		}

		for (int i = 1; i < num_attendance; i++) {
			String a = attendance.getRow(i).getCell(week).toString();
			if (a.equals("2.0")) {
				String c = attendance.getRow(i).getCell(57).toString();
				for (int j = 0; j < category1.length; j++) {
					if (c.equals(category1[j])) {
						switch (j) {
						case 0:
							list1[0][0][listNum1[0]] = attendance.getRow(i).getCell(0).toString();
							list1[0][1][listNum1[0]] = "1";
							listNum1[0]++;
							break;
						case 1:
							list1[1][0][listNum1[1]] = attendance.getRow(i).getCell(0).toString();
							list1[1][1][listNum1[1]] = "1";
							listNum1[1]++;
							break;
						case 2:
							list1[2][0][listNum1[2]] = attendance.getRow(i).getCell(0).toString();
							list1[2][1][listNum1[2]] = "1";
							listNum1[2]++;
							break;
						case 3:
							list1[3][0][listNum1[3]] = attendance.getRow(i).getCell(0).toString();
							list1[3][1][listNum1[3]] = "1";
							listNum1[3]++;
							break;
						case 4:
							list1[4][0][listNum1[4]] = attendance.getRow(i).getCell(0).toString();
							list1[4][1][listNum1[4]] = "1";
							listNum1[4]++;
							break;
						case 5:
							list1[5][0][listNum1[5]] = attendance.getRow(i).getCell(0).toString();
							list1[5][1][listNum1[5]] = "1";
							listNum1[5]++;
							break;
						case 6:
							list1[6][0][listNum1[6]] = attendance.getRow(i).getCell(0).toString();
							list1[6][1][listNum1[6]] = "1";
							listNum1[6]++;
							break;
						case 7:
							list1[6][0][listNum1[6]] = attendance.getRow(i).getCell(0).toString();
							list1[6][1][listNum1[6]] = "1";
							listNum1[6]++;
							break;
						case 8:
							list1[6][0][listNum1[6]] = attendance.getRow(i).getCell(0).toString();
							list1[6][1][listNum1[6]] = "1";
							listNum1[6]++;
							break;
						case 9:
							list1[7][0][listNum1[7]] = attendance.getRow(i).getCell(0).toString();
							list1[7][1][listNum1[7]] = "1";
							listNum1[7]++;
							break;
						case 10:
							list1[8][0][listNum1[8]] = attendance.getRow(i).getCell(0).toString();
							list1[8][1][listNum1[8]] = "1";
							listNum1[8]++;
							break;
						default:
							String d = attendance.getRow(i).getCell(1).toString();
							switch (d) {
							case "교역자":
								list2[0][0][listNum2[0]] = attendance.getRow(i).getCell(0).toString();
								list2[0][1][listNum2[0]] = "2";
								listNum2[0]++;
								break;
							case "장로":
								list2[1][0][listNum2[1]] = attendance.getRow(i).getCell(0).toString();
								list2[1][1][listNum2[1]] = "2";
								listNum2[1]++;
								break;
							case "안수집사":
								list2[2][0][listNum2[2]] = attendance.getRow(i).getCell(0).toString();
								list2[2][1][listNum2[2]] = "2";
								listNum2[2]++;
								break;
							case "권사":
								list2[3][0][listNum2[3]] = attendance.getRow(i).getCell(0).toString();
								list2[3][1][listNum2[3]] = "2";
								listNum2[3]++;
								break;
							case "서리집사(남)":
								list2[4][0][listNum2[4]] = attendance.getRow(i).getCell(0).toString();
								list2[4][1][listNum2[4]] = "2";
								listNum2[4]++;
								break;
							case "서리집사(여)":
								list2[5][0][listNum2[5]] = attendance.getRow(i).getCell(0).toString();
								list2[5][1][listNum2[5]] = "2";
								listNum2[5]++;
								break;
							case "권찰":
								list2[6][0][listNum2[6]] = attendance.getRow(i).getCell(0).toString();
								list2[6][1][listNum2[6]] = "2";
								listNum2[6]++;
								break;
							case "성도(남)":
								list2[6][0][listNum2[6]] = attendance.getRow(i).getCell(0).toString();
								list2[6][1][listNum2[6]] = "2";
								listNum2[6]++;
								break;
							case "성도(여)":
								list2[6][0][listNum2[6]] = attendance.getRow(i).getCell(0).toString();
								list2[6][1][listNum2[6]] = "2";
								listNum2[6]++;
								break;
							case "청년":
								list2[7][0][listNum2[7]] = attendance.getRow(i).getCell(0).toString();
								list2[7][1][listNum2[7]] = "2";
								listNum2[7]++;
								break;
							case "대학":
								list2[8][0][listNum2[8]] = attendance.getRow(i).getCell(0).toString();
								list2[8][1][listNum2[8]] = "2";
								listNum2[8]++;
								break;
							}
							break;
						}

						break;
					}
				}
			}
		}

		int[] listCount = new int[10];
		for (int i = 0; i < list.length - 1; i++) {
			listCount[i] = 0;
			if (listNum[i] != 0) {
				for (int j = 0; j < list1[i][0].length; j++) {
					list[i][0][listCount[i]] = list1[i][0][j];
					list[i][1][listCount[i]] = list1[i][1][j];
					listCount[i]++;
				}
				for (int j = 0; j < list2[i][0].length; j++) {
					list[i][0][listCount[i]] = list2[i][0][j];
					list[i][1][listCount[i]] = list2[i][1][j];
					listCount[i]++;
				}
				while (listCount[i] % 10 != 0) {
					list[i][0][listCount[i]] = "";
					list[i][1][listCount[i]] = "1";
					listCount[i]++;
				}
			} else {
				for (int j = 0; j < 10; j++) {
					list[i][0][listCount[i]] = "";
					list[i][1][listCount[i]] = "1";
					listCount[i]++;
				}
			}
		}
		listCount[9] = 0;
		if (listNum[9] != 0) {
			for (int j = 0; j < people1.size(); j++) {
				list[9][0][listCount[9]] = people1.get(j);
				list[9][1][listCount[9]] = "1";
				listCount[9]++;
			}
			for (int j = 0; j < people2.size(); j++) {
				list[9][0][listCount[9]] = people2.get(j);
				list[9][1][listCount[9]] = "2";
				listCount[9]++;
			}
			while (listCount[9] % 10 != 0) {
				list[9][0][listCount[9]] = "";
				list[9][1][listCount[9]] = "1";
				listCount[9]++;
			}
		} else {
			for (int j = 0; j < 10; j++) {
				list[9][0][listCount[9]] = "";
				list[9][1][listCount[9]] = "1";
				listCount[9]++;
			}
		}

		nLeft = 9;
		for (int i = 0; i < listNum.length; i++) {
			if (listNum[i] != 0 && listNum[i] % 10 == 0) {
				nLeft += listNum[i] / 10;
			} else {
				nLeft += listNum[i] / 10 + 1;
			}
		}

		bList = new String[19][6][];
		wmList = new String[19][4][];
		bN = new int[19];
		wmN = new int[19];
		nRight = 12;
		for (int i = 0; i < 19; i++) {
			XSSFSheet b = BELIEVER.getSheetAt(i);
			int bNum = b.getPhysicalNumberOfRows() - 1;
			bN[i] = bNum;
			XSSFSheet wm = WORDMOVEMENT.getSheetAt(i);
			int wmNum = wm.getPhysicalNumberOfRows() - 1;
			wmN[i] = wmNum;

			if (bNum == 0) {
				bList[i][0] = new String[1];
				bList[i][0][0] = "0";
			} else {
				int cc = 4 - bNum % 4;
				if (cc == 4) {
					cc = 0;
				}
				nRight += (bNum + cc) / 4;

				for (int j = 0; j < 6; j++) {
					bList[i][j] = new String[bNum + cc];
				}
				for (int j = 0; j < bNum; j++) {
					bList[i][0][j] = b.getRow(j + 1).getCell(0).toString();
					bList[i][1][j] = b.getRow(j + 1).getCell(1).toString();
					bList[i][2][j] = b.getRow(j + 1).getCell(2).toString();
					bList[i][3][j] = b.getRow(j + 1).getCell(3).toString();
					bList[i][4][j] = b.getRow(j + 1).getCell(4).toString();
					bList[i][5][j] = b.getRow(j + 1).getCell(5).toString();
				}
				for (int j = bNum; j < bNum + cc; j++) {
					for (int k = 0; k < 6; k++) {
						bList[i][k][j] = "";
					}
				}
			}

			if (wmNum == 0) {
				wmList[i][0] = new String[1];
				wmList[i][0][0] = "0";
			} else {
				int cc = 4 - wmNum % 4;
				if (cc == 4) {
					cc = 0;
				}
				nRight += (wmNum + cc) / 4;

				for (int j = 0; j < 4; j++) {
					wmList[i][j] = new String[wmNum + cc];
				}
				for (int j = 0; j < wmNum; j++) {
					wmList[i][0][j] = wm.getRow(j + 1).getCell(0).toString();
					wmList[i][1][j] = wm.getRow(j + 1).getCell(1).toString();
					wmList[i][2][j] = wm.getRow(j + 1).getCell(2).toString();
					wmList[i][3][j] = wm.getRow(j + 1).getCell(3).toString();
				}
				for (int j = wmNum; j < wmNum + cc; j++) {
					for (int k = 0; k < 4; k++) {
						wmList[i][k][j] = "";
					}
				}
			}
		}

		nRow2 = Math.max(nLeft, nRight);
	}

	public void performance2() {
		SH2.addMergedRegion(new CellRangeAddress(0, 2, 0, 29));
		SH2.createRow(0).createCell(0).setCellValue(title2);
		SH2.getRow(0).getCell(0).setCellStyle(csTitle2);

		for (int i = 0; i < 30; i++) {
			if (i == 0) {
				SH2.setColumnWidth(i, 2700);
			} else if (i == 11 || i == 29) {
				SH2.setColumnWidth(i, 1000);
			} else if (i == 16 || i == 20 || i == 24 || i == 28) {
				SH2.setColumnWidth(i, 3500);
			} else {
				SH2.setColumnWidth(i, 1700);
			}
		}

		SH2.setAutobreaks(false);
		SH2.setRowBreak(nRow2 - 1);
		SH2.setColumnBreak(29);

		SH2.createRow(4);
		SH2.getRow(4).createCell(15).setCellStyle(csIdxBc);
		SH2.getRow(4).createCell(16).setCellValue("지교회");
		SH2.getRow(4).createCell(19).setCellStyle(csIdxRemeet);
		SH2.getRow(4).createCell(20).setCellValue("영접+재만남");
		SH2.getRow(4).createCell(23).setCellStyle(csIdxBeliever);
		SH2.getRow(4).createCell(24).setCellValue("영접");

		SH2.createRow(6);
		SH2.addMergedRegion(new CellRangeAddress(6, 6, 0, 10));
		SH2.addMergedRegion(new CellRangeAddress(6, 6, 12, 28));
		SH2.addMergedRegion(new CellRangeAddress(6, 7, 29, 29));
		for (int i = 0; i < 30; i++) {
			if (i == 0 || i == 12) {
				SH2.getRow(6).createCell(i).setCellStyle(csMenuAboveLeft);
			} else if (i == 11 || i == 28 || i == 29) {
				SH2.getRow(6).createCell(i).setCellStyle(csMenuAboveRight);
			} else {
				SH2.getRow(6).createCell(i).setCellStyle(csMenuAboveCenter);
			}
		}
		SH2.getRow(6).getCell(0).setCellValue("현 장 동 역 자");
		SH2.getRow(6).getCell(11).setCellValue("계");
		SH2.getRow(6).getCell(12).setCellValue("영 접 자");
		SH2.getRow(6).getCell(29).setCellValue("계");

		SH2.createRow(7);
		for (int i = 12; i < 30; i++) {
			if (i == 12 || i == 13 || i == 17 || i == 21 || i == 25) {
				SH2.getRow(7).createCell(i).setCellStyle(csMenuBelowLeft);
			} else if (i == 16 || i == 20 || i == 24 || i == 28 || i == 29) {
				SH2.getRow(7).createCell(i).setCellStyle(csMenuBelowRight);
			} else {
				SH2.getRow(7).createCell(i).setCellStyle(csMenuBelowCenter);
			}
		}
		SH2.getRow(7).getCell(12).setCellValue("구분");
		SH2.getRow(7).getCell(13).setCellValue("전도자");
		SH2.getRow(7).getCell(14).setCellValue("영접자");
		SH2.getRow(7).getCell(15).setCellValue("사역자");
		SH2.getRow(7).getCell(16).setCellValue("나이/전화번호");
		SH2.getRow(7).getCell(17).setCellValue("전도자");
		SH2.getRow(7).getCell(18).setCellValue("영접자");
		SH2.getRow(7).getCell(19).setCellValue("사역자");
		SH2.getRow(7).getCell(20).setCellValue("나이/전화번호");
		SH2.getRow(7).getCell(21).setCellValue("전도자");
		SH2.getRow(7).getCell(22).setCellValue("영접자");
		SH2.getRow(7).getCell(23).setCellValue("사역자");
		SH2.getRow(7).getCell(24).setCellValue("나이/전화번호");
		SH2.getRow(7).getCell(25).setCellValue("전도자");
		SH2.getRow(7).getCell(26).setCellValue("영접자");
		SH2.getRow(7).getCell(27).setCellValue("사역자");
		SH2.getRow(7).getCell(28).setCellValue("나이/전화번호");

		for (int i = 8; i < nRow2; i++) {
			SH2.createRow(i);
		}

		int ccc = 7;
		int idx1 = 1;
		int idx2 = ccc;
		boolean first = true;
		int rowNum = 0;
		for (int i = 0; i < 10; i++) {
			first = true;
			rowNum = 0;
			for (int j = 0; j < list[i][0].length; j++) {
				if (first) {
					if (list[i][1][j].equals("1")) {
						SH2.getRow(idx2).createCell(idx1).setCellStyle(csATf);
					} else {
						SH2.getRow(idx2).createCell(idx1).setCellStyle(csBCf);
					}
				} else {
					if (list[i][1][j].equals("1")) {
						SH2.getRow(idx2).createCell(idx1).setCellStyle(csAT);
					} else {
						SH2.getRow(idx2).createCell(idx1).setCellStyle(csBC);
					}
				}
				SH2.getRow(idx2).getCell(idx1).setCellValue(list[i][0][j]);
				idx1++;
				if (idx1 == 11) {
					SH2.getRow(idx2).createCell(0).setCellStyle(csNormalLeft);
					SH2.getRow(idx2).createCell(11).setCellStyle(csNormalRight);
					idx1 = 1;
					idx2++;
					first = false;
					rowNum++;
				}
			}
			if (rowNum != 1) {
				if (i != 9) {
					SH2.addMergedRegion(new CellRangeAddress(idx2 - rowNum, idx2 - 1, 0, 0));
					SH2.addMergedRegion(new CellRangeAddress(idx2 - rowNum, idx2 - 1, 11, 11));
				} else {
					SH2.addMergedRegion(new CellRangeAddress(idx2 - rowNum, nRow2 - 3, 0, 0));
					SH2.addMergedRegion(new CellRangeAddress(idx2 - rowNum, nRow2 - 3, 11, 11));
				}
			}
			if (idx2 != nRow2 - 2) {
				for (int k = 0; k < nRow2 - 2 - idx2; k++) {
					SH2.getRow(idx2 + k).createCell(0).setCellStyle(csNormalLeft);
					SH2.getRow(idx2 + k).createCell(11).setCellStyle(csNormalRight);
				}
			}
			SH2.getRow(idx2 - rowNum).getCell(11).setCellValue(String.valueOf(listNum[i]));
			switch (i) {
			case 0:
				SH2.getRow(idx2 - rowNum).getCell(0).setCellValue("1.교역자");
				break;
			case 1:
				SH2.getRow(idx2 - rowNum).getCell(0).setCellValue("2.장로");
				break;
			case 2:
				SH2.getRow(idx2 - rowNum).getCell(0).setCellValue("3.안수집사");
				break;
			case 3:
				SH2.getRow(idx2 - rowNum).getCell(0).setCellValue("4.권사");
				break;
			case 4:
				SH2.getRow(idx2 - rowNum).getCell(0).setCellValue("5.서리집사(남)");
				break;
			case 5:
				SH2.getRow(idx2 - rowNum).getCell(0).setCellValue("6.서리집사(여)");
				break;
			case 6:
				SH2.getRow(idx2 - rowNum).getCell(0).setCellValue("7.권찰및성도");
				break;
			case 7:
				SH2.getRow(idx2 - rowNum).getCell(0).setCellValue("8.청년");
				break;
			case 8:
				SH2.getRow(idx2 - rowNum).getCell(0).setCellValue("9.대학");
				break;
			case 9:
				SH2.getRow(idx2 - rowNum).getCell(0).setCellValue("10.기타");
				break;
			}
		}

		SH2.addMergedRegion(new CellRangeAddress(nRow2 - 2, nRow2 - 1, 11, 11));
		SH2.addMergedRegion(new CellRangeAddress(nRow2 - 2, nRow2 - 1, 0, 0));
		SH2.addMergedRegion(new CellRangeAddress(nRow2 - 2, nRow2 - 2, 8, 10));
		SH2.addMergedRegion(new CellRangeAddress(nRow2 - 1, nRow2 - 1, 8, 10));
		for (int i = 0; i < 12; i++) {
			if (i == 0) {
				SH2.getRow(nRow2 - 2).createCell(i).setCellStyle(csMenuNormalLeft);
			} else if (i == 11) {
				SH2.getRow(nRow2 - 2).createCell(i).setCellStyle(csMenuNormalRight);
			} else {
				SH2.getRow(nRow2 - 2).createCell(i).setCellStyle(csMenuNormalCenter);
			}
		}
		SH2.getRow(nRow2 - 2).getCell(0).setCellValue("전체참석인원");
		SH2.getRow(nRow2 - 2).getCell(1).setCellValue("1조");
		SH2.getRow(nRow2 - 2).getCell(2).setCellValue("2조");
		SH2.getRow(nRow2 - 2).getCell(3).setCellValue("3조");
		SH2.getRow(nRow2 - 2).getCell(4).setCellValue("4조");
		SH2.getRow(nRow2 - 2).getCell(5).setCellValue("5조");
		SH2.getRow(nRow2 - 2).getCell(6).setCellValue("6조");
		SH2.getRow(nRow2 - 2).getCell(7).setCellValue("7조");
		SH2.getRow(nRow2 - 2).getCell(8).setCellValue("지교회(본교회포함)");
		for (int i = 0; i < 12; i++) {
			if (i == 0) {
				SH2.getRow(nRow2 - 1).createCell(i).setCellStyle(csMenuBelowLeft);
			} else if (i == 11) {
				SH2.getRow(nRow2 - 1).createCell(i).setCellStyle(csMenuBelowRight);
			} else {
				SH2.getRow(nRow2 - 1).createCell(i).setCellStyle(csMenuBelowCenter);
			}
		}
		int sum = 0;
		for (int i = 1; i <= 8; i++) {
			SH2.getRow(nRow2 - 1).getCell(i).setCellValue(String.valueOf(attNum[i - 1]));
			sum += attNum[i - 1];
		}
		SH2.getRow(nRow2 - 2).getCell(11).setCellValue(String.valueOf(sum));

		ccc = 8;
		idx1 = 0;
		idx2 = ccc;
		first = true;
		rowNum = 0;
		for (int i = 0; i < 19; i++) {
			first = true;
			rowNum = 0;
			if (!bList[i][0][0].equals("0")) {
				for (int j = 0; j < bList[i][0].length; j++) {
					SH2.getRow(idx2).createCell(idx1 * 4 + 13).setCellValue(bList[i][0][j]);
					SH2.getRow(idx2).createCell(idx1 * 4 + 14).setCellValue(bList[i][1][j]);
					SH2.getRow(idx2).createCell(idx1 * 4 + 15).setCellValue(bList[i][2][j]);
					if (bList[i][3][j].equals("")) {
						SH2.getRow(idx2).createCell(idx1 * 4 + 16).setCellValue(bList[i][4][j]);
					} else if (bList[i][4][j].equals("")) {
						SH2.getRow(idx2).createCell(idx1 * 4 + 16).setCellValue(bList[i][3][j]);
					} else {
						SH2.getRow(idx2).createCell(idx1 * 4 + 16).setCellValue(bList[i][3][j] + "/" + bList[i][4][j]);
					}
					if (first) {
						SH2.getRow(idx2).getCell(idx1 * 4 + 13).setCellStyle(csNormalAboveLeft);
						if (bList[i][5][j].equals("1")) {
							SH2.getRow(idx2).getCell(idx1 * 4 + 14).setCellStyle(csRemeetAbove);
						} else {
							SH2.getRow(idx2).getCell(idx1 * 4 + 14).setCellStyle(csNormalAboveCenter);
						}
						SH2.getRow(idx2).getCell(idx1 * 4 + 15).setCellStyle(csNormalAboveCenter);
						SH2.getRow(idx2).getCell(idx1 * 4 + 16).setCellStyle(csNormalAboveRight);
					} else {
						SH2.getRow(idx2).getCell(idx1 * 4 + 13).setCellStyle(csNormalLeft);
						if (bList[i][5][j].equals("1")) {
							SH2.getRow(idx2).getCell(idx1 * 4 + 14).setCellStyle(csRemeet);
						} else {
							SH2.getRow(idx2).getCell(idx1 * 4 + 14).setCellStyle(csNormalCenter);
						}
						SH2.getRow(idx2).getCell(idx1 * 4 + 15).setCellStyle(csNormalCenter);
						SH2.getRow(idx2).getCell(idx1 * 4 + 16).setCellStyle(csNormalRight);
					}
					idx1++;
					if (idx1 == 4) {
						SH2.getRow(idx2).createCell(12).setCellStyle(csSum);
						SH2.getRow(idx2).createCell(29).setCellStyle(csSum);
						idx1 = 0;
						idx2++;
						first = false;
						rowNum++;
					}
				}
				if (rowNum != 1) {
					SH2.addMergedRegion(new CellRangeAddress(idx2 - rowNum, idx2 - 1, 12, 12));
					SH2.addMergedRegion(new CellRangeAddress(idx2 - rowNum, idx2 - 1, 29, 29));
				}
				SH2.getRow(idx2 - rowNum).getCell(29).setCellValue(String.valueOf(bN[i]));
				switch (i) {
				case 0:
					SH2.getRow(idx2 - rowNum).getCell(12).setCellValue("1조");
					break;
				case 1:
					SH2.getRow(idx2 - rowNum).getCell(12).setCellValue("2조");
					break;
				case 2:
					SH2.getRow(idx2 - rowNum).getCell(12).setCellValue("3조");
					break;
				case 3:
					SH2.getRow(idx2 - rowNum).getCell(12).setCellValue("4조");
					break;
				case 4:
					SH2.getRow(idx2 - rowNum).getCell(12).setCellValue("5조");
					break;
				case 5:
					SH2.getRow(idx2 - rowNum).getCell(12).setCellValue("6조");
					break;
				case 6:
					SH2.getRow(idx2 - rowNum).getCell(12).setCellValue("7조");
					break;
				case 7:
					SH2.getRow(idx2 - rowNum).getCell(12).setCellValue("거창");
					break;
				case 8:
					SH2.getRow(idx2 - rowNum).getCell(12).setCellValue("상주");
					break;
				case 9:
					SH2.getRow(idx2 - rowNum).getCell(12).setCellValue("경주");
					break;
				case 10:
					SH2.getRow(idx2 - rowNum).getCell(12).setCellValue("성주");
					break;
				case 11:
					SH2.getRow(idx2 - rowNum).getCell(12).setCellValue("영주");
					break;
				case 12:
					SH2.getRow(idx2 - rowNum).getCell(12).setCellValue("북대구");
					break;
				case 13:
					SH2.getRow(idx2 - rowNum).getCell(12).setCellValue("포항");
					break;
				case 14:
					SH2.getRow(idx2 - rowNum).getCell(12).setCellValue("대전");
					break;
				case 15:
					SH2.getRow(idx2 - rowNum).getCell(12).setCellValue("부여");
					break;
				case 16:
					SH2.getRow(idx2 - rowNum).getCell(12).setCellValue("서대구");
					break;
				case 17:
					SH2.getRow(idx2 - rowNum).getCell(12).setCellValue("창원");
					break;
				case 18:
					SH2.getRow(idx2 - rowNum).getCell(12).setCellValue("북부산");
					break;
				}
			}
		}

		SH2.addMergedRegion(new CellRangeAddress(idx2, idx2, 12, 28));
		sum = 0;
		for (int i = 0; i < bN.length; i++) {
			sum += bN[i];
		}
		for (int i = 12; i <= 29; i++) {
			SH2.getRow(idx2).createCell(i).setCellStyle(csYellow);
			if (i == 12) {
				SH2.getRow(idx2).getCell(i).setCellValue("영 접 자 합 계");
			} else if (i == 29) {
				SH2.getRow(idx2).getCell(i).setCellValue(String.valueOf(sum));
			}
		}

		int needs = 0;
		for (int i = 0; i < 19; i++) {
			if (!wmList[i][0][0].equals("0")) {
				needs += wmList[i][0].length;
			}
		}
		needs = needs / 4;
		ccc = nRow2 - 3 - needs;

		for (int i = idx2 + 1; i < ccc; i++) {
			SH2.getRow(i).createCell(29).setCellStyle(csLine);
		}

		SH2.addMergedRegion(new CellRangeAddress(ccc, ccc, 12, 28));
		SH2.addMergedRegion(new CellRangeAddress(ccc, ccc + 1, 29, 29));
		for (int i = 12; i < 30; i++) {
			if (i == 12) {
				SH2.getRow(ccc).createCell(i).setCellStyle(csMenuAboveLeft);
			} else if (i == 28 || i == 29) {
				SH2.getRow(ccc).createCell(i).setCellStyle(csMenuAboveRight);
			} else {
				SH2.getRow(ccc).createCell(i).setCellStyle(csMenuAboveCenter);
			}
		}
		SH2.getRow(ccc).getCell(12).setCellValue("말 씀 운 동");
		SH2.getRow(ccc).getCell(29).setCellValue("계");
		ccc++;

		for (int i = 12; i < 30; i++) {
			if (i == 12 || i == 13 || i == 17 || i == 21 || i == 25) {
				SH2.getRow(ccc).createCell(i).setCellStyle(csMenuBelowLeft);
			} else if (i == 16 || i == 20 || i == 24 || i == 28 || i == 29) {
				SH2.getRow(ccc).createCell(i).setCellStyle(csMenuBelowRight);
			} else {
				SH2.getRow(ccc).createCell(i).setCellStyle(csMenuBelowCenter);
			}
		}
		SH2.addMergedRegion(new CellRangeAddress(ccc, ccc, 14, 15));
		SH2.getRow(ccc).getCell(12).setCellValue("구분");
		SH2.getRow(ccc).getCell(13).setCellValue("사역자");
		SH2.getRow(ccc).getCell(14).setCellValue("영접자/만난횟수");
		SH2.getRow(ccc).getCell(16).setCellValue("장소");
		SH2.addMergedRegion(new CellRangeAddress(ccc, ccc, 18, 19));
		SH2.getRow(ccc).getCell(17).setCellValue("사역자");
		SH2.getRow(ccc).getCell(18).setCellValue("영접자/만난횟수");
		SH2.getRow(ccc).getCell(20).setCellValue("장소");
		SH2.addMergedRegion(new CellRangeAddress(ccc, ccc, 22, 23));
		SH2.getRow(ccc).getCell(21).setCellValue("사역자");
		SH2.getRow(ccc).getCell(22).setCellValue("영접자/만난횟수");
		SH2.getRow(ccc).getCell(24).setCellValue("장소");
		SH2.addMergedRegion(new CellRangeAddress(ccc, ccc, 26, 27));
		SH2.getRow(ccc).getCell(25).setCellValue("사역자");
		SH2.getRow(ccc).getCell(26).setCellValue("영접자/만난횟수");
		SH2.getRow(ccc).getCell(28).setCellValue("장소");

		ccc++;
		idx1 = 0;
		idx2 = ccc;
		first = true;
		rowNum = 0;
		for (int i = 0; i < 19; i++) {
			first = true;
			rowNum = 0;
			if (!wmList[i][0][0].equals("0")) {
				for (int j = 0; j < wmList[i][0].length; j++) {
					SH2.addMergedRegion(new CellRangeAddress(idx2, idx2, idx1 * 4 + 14, idx1 * 4 + 15));
					SH2.getRow(idx2).createCell(idx1 * 4 + 13).setCellValue(wmList[i][0][j]);
					if (wmList[i][2][j].equals("")) {
						SH2.getRow(idx2).createCell(idx1 * 4 + 14).setCellValue(wmList[i][1][j]);
					} else {
						SH2.getRow(idx2).createCell(idx1 * 4 + 14)
								.setCellValue(wmList[i][1][j] + "/" + wmList[i][2][j]);
					}
					SH2.getRow(idx2).createCell(idx1 * 4 + 16).setCellValue(wmList[i][3][j]);
					if (first) {
						SH2.getRow(idx2).getCell(idx1 * 4 + 13).setCellStyle(csNormalAboveLeft);
						SH2.getRow(idx2).getCell(idx1 * 4 + 14).setCellStyle(csNormalAboveCenter);
						SH2.getRow(idx2).createCell(idx1 * 4 + 15).setCellStyle(csNormalAboveCenter);
						SH2.getRow(idx2).getCell(idx1 * 4 + 16).setCellStyle(csNormalAboveRight);
					} else {
						SH2.getRow(idx2).getCell(idx1 * 4 + 13).setCellStyle(csNormalLeft);
						SH2.getRow(idx2).getCell(idx1 * 4 + 14).setCellStyle(csNormalCenter);
						SH2.getRow(idx2).createCell(idx1 * 4 + 15).setCellStyle(csNormalCenter);
						SH2.getRow(idx2).getCell(idx1 * 4 + 16).setCellStyle(csNormalRight);
					}
					idx1++;
					if (idx1 == 4) {
						SH2.getRow(idx2).createCell(12).setCellStyle(csSum);
						SH2.getRow(idx2).createCell(29).setCellStyle(csSum);
						idx1 = 0;
						idx2++;
						first = false;
						rowNum++;
					}
				}
				if (rowNum != 1) {
					SH2.addMergedRegion(new CellRangeAddress(idx2 - rowNum, idx2 - 1, 12, 12));
					SH2.addMergedRegion(new CellRangeAddress(idx2 - rowNum, idx2 - 1, 29, 29));
				}
				SH2.getRow(idx2 - rowNum).getCell(29).setCellValue(String.valueOf(wmN[i]));
				switch (i) {
				case 0:
					SH2.getRow(idx2 - rowNum).getCell(12).setCellValue("1조");
					break;
				case 1:
					SH2.getRow(idx2 - rowNum).getCell(12).setCellValue("2조");
					break;
				case 2:
					SH2.getRow(idx2 - rowNum).getCell(12).setCellValue("3조");
					break;
				case 3:
					SH2.getRow(idx2 - rowNum).getCell(12).setCellValue("4조");
					break;
				case 4:
					SH2.getRow(idx2 - rowNum).getCell(12).setCellValue("5조");
					break;
				case 5:
					SH2.getRow(idx2 - rowNum).getCell(12).setCellValue("6조");
					break;
				case 6:
					SH2.getRow(idx2 - rowNum).getCell(12).setCellValue("7조");
					break;
				case 7:
					SH2.getRow(idx2 - rowNum).getCell(12).setCellValue("거창");
					break;
				case 8:
					SH2.getRow(idx2 - rowNum).getCell(12).setCellValue("상주");
					break;
				case 9:
					SH2.getRow(idx2 - rowNum).getCell(12).setCellValue("경주");
					break;
				case 10:
					SH2.getRow(idx2 - rowNum).getCell(12).setCellValue("성주");
					break;
				case 11:
					SH2.getRow(idx2 - rowNum).getCell(12).setCellValue("영주");
					break;
				case 12:
					SH2.getRow(idx2 - rowNum).getCell(12).setCellValue("북대구");
					break;
				case 13:
					SH2.getRow(idx2 - rowNum).getCell(12).setCellValue("포항");
					break;
				case 14:
					SH2.getRow(idx2 - rowNum).getCell(12).setCellValue("대전");
					break;
				case 15:
					SH2.getRow(idx2 - rowNum).getCell(12).setCellValue("부여");
					break;
				case 16:
					SH2.getRow(idx2 - rowNum).getCell(12).setCellValue("서대구");
					break;
				case 17:
					SH2.getRow(idx2 - rowNum).getCell(12).setCellValue("창원");
					break;
				case 18:
					SH2.getRow(idx2 - rowNum).getCell(12).setCellValue("북부산");
					break;
				}
			}
		}

		SH2.addMergedRegion(new CellRangeAddress(idx2, idx2, 12, 28));
		sum = 0;
		for (int i = 0; i < bN.length; i++) {
			sum += wmN[i];
		}
		for (int i = 12; i <= 29; i++) {
			SH2.getRow(idx2).createCell(i).setCellStyle(csYellow);
			if (i == 12) {
				SH2.getRow(idx2).getCell(i).setCellValue("말 씀 운 동 합 계");
			} else if (i == 29) {
				SH2.getRow(idx2).getCell(i).setCellValue(String.valueOf(sum));
			}
		}

	}

	public int getNumRow1() {
		return nRow1;
	}
}
