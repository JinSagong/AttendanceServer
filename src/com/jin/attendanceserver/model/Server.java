package com.jin.attendanceserver.model;

import java.io.BufferedReader;
import java.io.InputStreamReader;
import java.io.PrintWriter;
import java.net.Socket;

import org.apache.commons.math3.util.Pair;

public class Server extends Thread {

	BufferedReader inputReader;
	PrintWriter outputWriter;

	String i_id, i_key, i_type, i_num, i_word, i_addname, i_addgroup;
	String[] o_user, o_status, o_logs, o_fpattendance_etc, o_fpattendance_search, i_fpa_names, i_fpa_contents,
			i_fpa_etc, i_fpa_search, i_believer_name, i_wordmovement_name, o_consideration;
	String[][] o_list, o_attendance, i_attendance, o_fpattendance, o_searchresult, o_believer, o_wordmovement,
			i_believer, i_wordmovement;
	int[] i_fpa_checks;
	boolean o_dup;

	public Server(Socket socket, DatabaseManagement dbm) throws Exception {

		inputReader = new BufferedReader(new InputStreamReader(socket.getInputStream(), "UTF-8"));
		outputWriter = new PrintWriter(socket.getOutputStream());

		i_type = inputReader.readLine();

		switch (i_type) {

		// isAlive
		case "isAlive":
			// do nothing
			break;

		case "getOnOff":
			Pair<Boolean, Boolean> onoff = dbm.getOnOff();
			outputWriter.println(onoff.getFirst());
			outputWriter.flush();
			outputWriter.println(onoff.getSecond());
			outputWriter.flush();
			break;

		// LogIn
		case "LogIn":
			i_id = inputReader.readLine();
			o_user = new String[4];
			o_user = dbm.readId(i_id, null);
			if (o_user[0] != null) {
				outputWriter.println(2);
				outputWriter.flush();
				outputWriter.println(o_user[2]);
				outputWriter.flush();
				outputWriter.println(o_user[3]);
				outputWriter.flush();
				dbm.writeLogs(o_user[2], "LogIn", null);
			} else {
				outputWriter.println(1);
				outputWriter.flush();
				outputWriter.println("FALSE");
				outputWriter.flush();
			}
			break;

		// AutoLogIn
		case "AutoLogIn":
			i_key = inputReader.readLine();
			outputWriter.println("TRUE");
			outputWriter.flush();
			dbm.writeLogs(i_key, "AutoLogIn", null);
			break;

		// LogOut
		case "LogOut":
			i_key = inputReader.readLine();
			outputWriter.println("TRUE");
			outputWriter.flush();
			dbm.writeLogs(i_key, "LogOut", null);
			break;

		// getList
		case "getList":
			i_type = inputReader.readLine();
			o_list = dbm.readStatus(i_type);
			outputWriter.println(o_list[0].length);
			outputWriter.flush();
			for (int i = 0; i < o_list[0].length; i++) {
				outputWriter.println(o_list[1][i]);
				outputWriter.flush();
				outputWriter.println(o_list[2][i]);
				outputWriter.flush();
			}
			break;

		// getStatus
		case "getStatus":
			i_type = inputReader.readLine();
			o_status = dbm.readStatus(i_type)[2];
			outputWriter.println(o_status.length);
			outputWriter.flush();
			for (int i = 0; i < o_status.length; i++) {
				outputWriter.println(o_status[i]);
				outputWriter.flush();
			}
			break;

		// getLogs
		case "getLogs":
			int numlog = dbm.getNumLogs();
			if (numlog > 299) {
				numlog = 299;
			}
			o_logs = new String[numlog];
			o_logs = dbm.readLogs();
			outputWriter.println(o_logs.length);
			outputWriter.flush();
			for (int i = 0; i < o_logs.length; i++) {
				outputWriter.println(o_logs[i]);
				outputWriter.flush();
			}
			break;

		// getAttendance
		case "getAttendance":
			i_type = inputReader.readLine();
			o_attendance = dbm.readAttendance(i_type);
			outputWriter.println(o_attendance[0].length);
			outputWriter.flush();
			for (int i = 0; i < o_attendance[0].length; i++) {
				outputWriter.println(o_attendance[0][i]);
				outputWriter.flush();
				outputWriter.println(o_attendance[1][i]);
				outputWriter.flush();
				outputWriter.println(o_attendance[2][i]);
				outputWriter.flush();
			}
			break;

		// setAttendance
		case "setAttendance":
			i_key = inputReader.readLine();
			i_type = inputReader.readLine();
			i_num = inputReader.readLine();
			i_attendance = new String[3][Integer.valueOf(i_num)];
			for (int i = 0; i < i_attendance[0].length; i++) {
				i_attendance[0][i] = inputReader.readLine();
				i_attendance[1][i] = inputReader.readLine();
				i_attendance[2][i] = inputReader.readLine();
			}
			if (i_type.contains("#"))
				dbm.writeHomeAttendance(i_type, i_attendance);
			else
				dbm.writeAttendance(i_type, i_attendance);
			dbm.writeStatus(i_type);
			outputWriter.println("TRUE");
			outputWriter.flush();
			dbm.writeLogs(i_key, "Check", i_type);
			break;

		// getFpAttendance
		case "getFpAttendance":
			i_type = inputReader.readLine();
			o_fpattendance = dbm.readFpAttendance(i_type);
			o_fpattendance_etc = dbm.readFpEtcAttendance(i_type);
			o_fpattendance_search = dbm.readFpSearchAttendance(i_type);
			outputWriter.println(o_fpattendance[0].length);
			outputWriter.flush();
			for (int i = 0; i < o_fpattendance[0].length; i++) {
				outputWriter.println(o_fpattendance[0][i]);
				outputWriter.flush();
				outputWriter.println(o_fpattendance[1][i]);
				outputWriter.flush();
				outputWriter.println(o_fpattendance[2][i]);
				outputWriter.flush();
				outputWriter.println(o_fpattendance[3][i]);
				outputWriter.flush();
				outputWriter.println(o_fpattendance[4][i]);
				outputWriter.flush();
				outputWriter.println(o_fpattendance[5][i]);
				outputWriter.flush();
				outputWriter.println(o_fpattendance[6][i]);
				outputWriter.flush();
			}
			outputWriter.println(o_fpattendance_etc.length);
			outputWriter.flush();
			for (int i = 0; i < o_fpattendance_etc.length; i++) {
				outputWriter.println(o_fpattendance_etc[i]);
				outputWriter.flush();
			}
			outputWriter.println(o_fpattendance_search.length);
			outputWriter.flush();
			for (int i = 0; i < o_fpattendance_search.length; i++) {
				outputWriter.println(o_fpattendance_search[i]);
				outputWriter.flush();
			}
			break;

		// getSearchResult
		case "getSearchResult":
			i_type = inputReader.readLine();
			o_searchresult = dbm.readSearchResult(i_type);
			outputWriter.println(o_searchresult[0].length);
			outputWriter.flush();
			for (int i = 0; i < o_searchresult[0].length; i++) {
				outputWriter.println(o_searchresult[0][i]);
				outputWriter.flush();
				outputWriter.println(o_searchresult[1][i]);
				outputWriter.flush();
				outputWriter.println(o_searchresult[2][i]);
				outputWriter.flush();
				outputWriter.println(o_searchresult[3][i]);
				outputWriter.flush();
			}
			break;

		// setFpAttendance
		case "setFpAttendance":
			i_key = inputReader.readLine();
			i_type = inputReader.readLine();
			i_num = inputReader.readLine();
			i_fpa_names = new String[Integer.valueOf(i_num)];
			i_fpa_contents = new String[Integer.valueOf(i_num)];
			i_fpa_checks = new int[Integer.valueOf(i_num)];
			for (int i = 0; i < i_fpa_names.length; i++) {
				i_fpa_names[i] = inputReader.readLine();
				i_fpa_contents[i] = inputReader.readLine();
				i_fpa_checks[i] = Integer.valueOf(inputReader.readLine());
			}
			i_num = inputReader.readLine();
			i_fpa_etc = new String[Integer.valueOf(i_num)];
			for (int i = 0; i < i_fpa_etc.length; i++) {
				i_fpa_etc[i] = inputReader.readLine();
			}
			i_num = inputReader.readLine();
			i_fpa_search = new String[Integer.valueOf(i_num)];
			for (int i = 0; i < i_fpa_search.length; i++) {
				i_fpa_search[i] = inputReader.readLine();
			}
			dbm.writeFpAttendance(i_type, i_fpa_names, i_fpa_contents, i_fpa_checks, i_fpa_etc, i_fpa_search);
			dbm.writeStatus(i_type);
			outputWriter.println("TRUE");
			outputWriter.flush();
			dbm.writeLogs(i_key, "FpCheck", i_type);
			break;

		// getFpFruits
		case "getFpFruits":
			i_type = inputReader.readLine();
			o_believer = dbm.readFpBeliever(i_type);
			o_wordmovement = dbm.readFpWordMovement(i_type);
			outputWriter.println(o_believer[0].length);
			outputWriter.flush();
			for (int i = 0; i < o_believer[0].length; i++) {
				outputWriter.println(o_believer[0][i]);
				outputWriter.flush();
				outputWriter.println(o_believer[1][i]);
				outputWriter.flush();
				outputWriter.println(o_believer[2][i]);
				outputWriter.flush();
				outputWriter.println(o_believer[3][i]);
				outputWriter.flush();
				outputWriter.println(o_believer[4][i]);
				outputWriter.flush();
				outputWriter.println(o_believer[5][i]);
				outputWriter.flush();
			}
			outputWriter.println(o_wordmovement[0].length);
			outputWriter.flush();
			for (int i = 0; i < o_wordmovement[0].length; i++) {
				outputWriter.println(o_wordmovement[0][i]);
				outputWriter.flush();
				outputWriter.println(o_wordmovement[1][i]);
				outputWriter.flush();
				outputWriter.println(o_wordmovement[2][i]);
				outputWriter.flush();
				outputWriter.println(o_wordmovement[3][i]);
				outputWriter.flush();
			}
			break;

		// getSearchResultForFruits
		case "getSearchResultForFruits":
			i_type = inputReader.readLine();
			i_word = inputReader.readLine();
			o_searchresult = dbm.readSearchResultForFruits(i_type, i_word);
			outputWriter.println(o_searchresult[0].length);
			outputWriter.flush();
			for (int i = 0; i < o_searchresult[0].length; i++) {
				outputWriter.println(o_searchresult[0][i]);
				outputWriter.flush();
				outputWriter.println(o_searchresult[1][i]);
				outputWriter.flush();
			}
			break;

		// setFpFruits
		case "setFpFruits":
			i_key = inputReader.readLine();
			i_type = inputReader.readLine();
			i_num = inputReader.readLine();
			i_believer = new String[6][Integer.valueOf(i_num)];
			for (int i = 0; i < i_believer[0].length; i++) {
				for (int j = 0; j < 6; j++) {
					i_believer[j][i] = inputReader.readLine();
				}
			}
			i_num = inputReader.readLine();
			i_wordmovement = new String[4][Integer.valueOf(i_num)];
			for (int i = 0; i < i_wordmovement[0].length; i++) {
				for (int j = 0; j < 4; j++) {
					i_wordmovement[j][i] = inputReader.readLine();
				}
			}
			dbm.writeFpFruits(i_type, i_believer, i_wordmovement);
			dbm.writeStatus(i_type);
			outputWriter.println("TRUE");
			outputWriter.flush();
			dbm.writeLogs(i_key, "FpCheck", i_type);
			break;

		// setStatus
		case "setStatus":
			i_type = inputReader.readLine();
			dbm.writeFpStatus(i_type);
			outputWriter.println("TRUE");
			outputWriter.flush();
			break;

		// getConsideration
		case "getConsideration":
			i_type = inputReader.readLine();
			i_num = inputReader.readLine();
			i_believer_name = new String[Integer.valueOf(i_num)];
			for (int i = 0; i < i_believer_name.length; i++) {
				i_believer_name[i] = inputReader.readLine();
			}
			i_num = inputReader.readLine();
			i_wordmovement_name = new String[Integer.valueOf(i_num)];
			for (int i = 0; i < i_wordmovement_name.length; i++) {
				i_wordmovement_name[i] = inputReader.readLine();
			}
			o_consideration = dbm.getConsideration(i_type, i_believer_name, i_wordmovement_name);
			outputWriter.println(o_consideration[0]);
			outputWriter.flush();
			outputWriter.println(o_consideration[1]);
			outputWriter.flush();
			outputWriter.println(o_consideration[2]);
			outputWriter.flush();
			break;

		// getFpDuplication
		case "getFpDuplication":
			i_word = inputReader.readLine();
			o_dup = dbm.getFpDuplication(i_word);
			if (o_dup) {
				outputWriter.println("TRUE");
				outputWriter.flush();
			} else {
				outputWriter.println("FALSE");
				outputWriter.flush();
			}
			break;

		// addFpAttendance
		case "addFpAttendance":
			i_key = inputReader.readLine();
			i_type = inputReader.readLine();
			i_addname = inputReader.readLine();
			i_addgroup = inputReader.readLine();
			dbm.addFpAttendance(i_type, i_addname, i_addgroup);
			outputWriter.println("TRUE");
			outputWriter.flush();
			dbm.writeLogs(i_key, "FpAdd", i_type);
			break;

		// getDuplication
		case "getDuplication":
			i_word = inputReader.readLine();
			o_dup = dbm.getDuplication(i_word);
			if (o_dup) {
				outputWriter.println("TRUE");
				outputWriter.flush();
			} else {
				outputWriter.println("FALSE");
				outputWriter.flush();
			}
			break;

		// addAttendance
		case "addAttendance":
			i_key = inputReader.readLine();
			i_type = inputReader.readLine();
			i_addname = inputReader.readLine();
			i_addgroup = inputReader.readLine();
			dbm.addAttendance(i_type, i_addname, i_addgroup);
			outputWriter.println("TRUE");
			outputWriter.flush();
			dbm.writeLogs(i_key, "Add", i_type);
			break;
		}

		inputReader.close();
		outputWriter.close();
		socket.close();
	}
}
