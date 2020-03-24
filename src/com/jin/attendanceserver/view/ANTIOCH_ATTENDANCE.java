package com.jin.attendanceserver.view;

import java.awt.*;
import java.awt.event.*;
import java.io.IOException;
import java.net.ServerSocket;
import java.net.Socket;
import java.text.SimpleDateFormat;
import java.util.Date;

import javax.swing.*;

import com.jin.attendanceserver.model.DatabaseManagement;
import com.jin.attendanceserver.model.Server;
import com.jin.attendanceserver.util.DirectorySharedPreference;
import com.jin.attendanceserver.util.RunTimeMeasurement;
import com.jin.attendanceserver.util.RunningSharedPreference;

public class ANTIOCH_ATTENDANCE {

	public static void main(String[] args) throws IOException {
		// TODO Auto-generated method stub

		new ServerFrame();
	}
}

class ServerFrame extends JFrame implements ActionListener {
	private static final long serialVersionUID = 1L;

	final private int WIDTH = 800;
	final private int HEIGHT = 600;
	final private int MARGIN = 20;
	final private int BUTTON_WIDTH = 150;
	final private int BUTTON_HEIGHT = 60;
	final private int LABEL_HEIGHT = 20;

	// Define functions
	JPanel ServerPanel;
	Font default_font, title_font;
	JLabel Title, DirectoryPath, Status, Time;
	JButton ServerOn, ServerOff, Finder;
	TextArea Logs;
	SimpleDateFormat formatter;
	RunTimeMeasurement rtm;
	Thread th_server, th_timer, th_saver;
	String time, run_time;
	long startTime, currentTime;
	boolean powered, stopped;

	ServerSocket s_socket;
	Socket c_socket;
	final private int PORT = 8888;
	private String DIRECTORY_PATH;
	DatabaseManagement DB_Manager;
	DirectorySharedPreference dsp;
	RunningSharedPreference rsp;

	ServerFrame() {
		super();

		rsp = new RunningSharedPreference();
		boolean flag = rsp.getRunning();

		if (flag) {
			JOptionPane.showMessageDialog(this, "프로그램이 이미 가동중입니다.");
			dispose();

		} else {
			rsp.setRunning(true);

			initLayout();
			initFunction();

			th_server = new Thread() {
				@Override
				public void run() {
					while (powered) {
						synchronized (this) {
							try {
								wait();
							} catch (InterruptedException e) {
								// TODO Auto-generated catch block
							}
						}

						while (!stopped) {
							try {
								c_socket = s_socket.accept();
								new Server(c_socket, DB_Manager);
							} catch (Exception e) {
								// TODO Auto-generated catch block
							}
						}
					}
				}

			};

			th_timer = new Thread() {
				@Override
				public void run() {
					while (powered) {
						synchronized (this) {
							try {
								wait();
							} catch (InterruptedException e) {
								// TODO Auto-generated catch block
							}
						}

						while (!stopped) {
							try {
								Thread.sleep(100);
							} catch (InterruptedException e) {
								// TODO Auto-generated catch block
							}
							currentTime = System.currentTimeMillis();
							Time.setText("Run Time:  " + rtm.getRunTime(currentTime - startTime));
						}
					}

				}
			};

			th_saver = new Thread() {
				@Override
				public void run() {
					while (powered) {
						boolean saved = true;
						synchronized (this) {
							try {
								wait();
							} catch (InterruptedException e) {
								// TODO Auto-generated catch block
							}
						}

						while (!stopped) {
							try {
								Thread.sleep(100);
							} catch (InterruptedException e) {
								// TODO Auto-generated catch block
							}
							// every 30s
							if (!saved && (currentTime - startTime) % 30000 < 1000) {
								saved = true;
								Status.setText("Status of server... Server On (saving)");
								DB_Manager.save();
							} else if (saved && (currentTime - startTime) % 30000 > 3000
									&& (currentTime - startTime) % 30000 < 4000) {
								saved = false;
								Status.setText("Status of server... Server On");
							}
						}
					}

				}
			};

			th_server.start();
			th_timer.start();
			th_saver.start();
		}
	}

	public void initLayout() {

		// Set fonts
		default_font = new Font("", Font.BOLD, 16);
		title_font = new Font("궁서", Font.BOLD, 30);

		// Set main panel
		ServerPanel = new JPanel();
		ServerPanel.setPreferredSize(new Dimension(WIDTH, HEIGHT));
		ServerPanel.setBackground(Color.WHITE);

		// Set title label
		Title = new JLabel("구 우 미 이 서 어 버 어");
		Title.setFont(title_font);
		Title.setHorizontalAlignment(JLabel.CENTER);

		// Set buttons
		ServerOn = new JButton("Server On");
		ServerOff = new JButton("Server Off");
		ServerOn.setFont(default_font);
		ServerOff.setFont(default_font);
		ServerOn.setBackground(Color.WHITE);
		ServerOff.setBackground(Color.WHITE);
		ServerOn.setFocusable(false);
		ServerOff.setFocusable(false);
		ServerOff.setEnabled(false);

		ServerOn.addActionListener(this);
		ServerOff.addActionListener(this);

		// Set finder button
		Finder = new JButton("...");
		Finder.setFont(default_font);
		Finder.setBackground(Color.WHITE);
		Finder.setFocusable(false);
		Finder.addActionListener(this);

		// Set directory path label
		DirectoryPath = new JLabel();
		DirectoryPath.setFont(default_font);
		DirectoryPath.setForeground(Color.GRAY);

		// Set status label
		Status = new JLabel("Status of server... Server Off");
		Status.setFont(default_font);

		// Set time label
		Time = new JLabel();
		Time.setFont(default_font);
		run_time = "Run Time:  00:00:00";
		Time.setText(run_time);
		Time.setHorizontalAlignment(JLabel.RIGHT);

		// Set logs area
		Logs = new TextArea();
		Logs.setFont(default_font);
		Logs.setEditable(false);

		// Set layout
		ServerPanel.setLayout(null);
		Title.setBounds(0, MARGIN, WIDTH, HEIGHT / 12);
		ServerOn.setBounds(WIDTH / 2 - MARGIN - BUTTON_WIDTH, HEIGHT / 6, BUTTON_WIDTH, BUTTON_HEIGHT);
		ServerOff.setBounds(WIDTH / 2 + MARGIN, HEIGHT / 6, BUTTON_WIDTH, BUTTON_HEIGHT);
		Finder.setBounds(MARGIN, HEIGHT / 2 - MARGIN * 2 - LABEL_HEIGHT * 2, LABEL_HEIGHT, LABEL_HEIGHT);
		DirectoryPath.setBounds(MARGIN * 3 / 2 + LABEL_HEIGHT, HEIGHT / 2 - MARGIN * 2 - LABEL_HEIGHT * 2,
				WIDTH - LABEL_HEIGHT - MARGIN * 5 / 2, LABEL_HEIGHT);
		Status.setBounds(MARGIN, HEIGHT / 2 - MARGIN * 2 - LABEL_HEIGHT, WIDTH / 3 * 2 - MARGIN, LABEL_HEIGHT);
		Time.setBounds(WIDTH / 3 * 2, HEIGHT / 2 - MARGIN * 2 - LABEL_HEIGHT, WIDTH / 3 - MARGIN, LABEL_HEIGHT);
		Logs.setBounds(MARGIN, HEIGHT / 2 - MARGIN, WIDTH - MARGIN * 2, HEIGHT / 2 - MARGIN * 2);

		// Add items
		ServerPanel.add(Title);
		ServerPanel.add(ServerOn);
		ServerPanel.add(ServerOff);
		ServerPanel.add(Finder);
		ServerPanel.add(DirectoryPath);
		ServerPanel.add(Status);
		ServerPanel.add(Time);
		ServerPanel.add(Logs);
		add(ServerPanel);

		// Set main frame
		setBounds(10, 10, WIDTH, HEIGHT);
		setTitle("ANTIOCH ATTENDANCE SERVER");
		setResizable(false);
		setVisible(true);
		setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
		this.addWindowListener(new WindowAdapter() {
			@Override
			public void windowClosing(WindowEvent e) {
				rsp.setRunning(false);
			}
		});
	}

	public void initFunction() {
		dsp = new DirectorySharedPreference();

		DIRECTORY_PATH = dsp.getDirectory();
		DirectoryPath.setText("Directory Path: " + DIRECTORY_PATH);
		dsp.setDirectory(DIRECTORY_PATH);

		formatter = new SimpleDateFormat("yyyy-MM-dd HH:mm:ss.SSS");
		rtm = new RunTimeMeasurement();

		powered = true;
		stopped = true;
	}

	@Override
	public void actionPerformed(ActionEvent e) {
		// TODO Auto-generated method stub

		if (e.getSource() == ServerOn) {
			ServerOn();

		} else if (e.getSource() == ServerOff) {
			ServerOff();

		} else if (e.getSource() == Finder) {
			if (stopped) {
				try {
					UIManager.setLookAndFeel(UIManager.getSystemLookAndFeelClassName());
					JFileChooser fc = new JFileChooser();
					fc.setFileSelectionMode(JFileChooser.DIRECTORIES_ONLY);
					fc.setAcceptAllFileFilterUsed(false);
					if (fc.showOpenDialog(null) == JFileChooser.APPROVE_OPTION) {
						DIRECTORY_PATH = fc.getSelectedFile().toString();
						DirectoryPath.setText("Directory Path:  " + DIRECTORY_PATH);
						dsp.setDirectory(DIRECTORY_PATH);
					}
					UIManager.setLookAndFeel(UIManager.getCrossPlatformLookAndFeelClassName());
				} catch (ClassNotFoundException | InstantiationException | IllegalAccessException
						| UnsupportedLookAndFeelException e1) {
					// TODO Auto-generated catch block
				}
			}
		}
	}

	public void ServerOn() {
		ServerOn.setEnabled(false);
		ServerOff.setEnabled(true);
		setDefaultCloseOperation(JFrame.DO_NOTHING_ON_CLOSE);

		DB_Manager = new DatabaseManagement(DIRECTORY_PATH, Logs, formatter);
		DB_Manager.init();

		time = formatter.format(new Date());
		startTime = System.currentTimeMillis();
		stopped = false;

		try {
			s_socket = new ServerSocket(PORT);
		} catch (IOException e) {
			// TODO Auto-generated catch block
		}

		synchronized (th_server) {
			th_server.notify();
		}
		synchronized (th_timer) {
			th_timer.notify();
		}
		synchronized (th_saver) {
			th_saver.notify();
		}

		Status.setText("Status of server... Server On");
		DB_Manager.writeLogs(null, "TurnOn", null);
	}

	public void ServerOff() {
		ServerOn.setEnabled(true);
		ServerOff.setEnabled(false);
		setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);

		time = formatter.format(new Date());
		stopped = true;
		try {
			s_socket.close();
		} catch (IOException e) {
			// TODO Auto-generated catch block
		}
		try {
			Thread.sleep(100);
		} catch (InterruptedException e) {
			// TODO Auto-generated catch block
		}

		DB_Manager.saveArchive();
		DB_Manager.writeLogs(null, "TurnOff", rtm.getRunTime(currentTime - startTime));
		Time.setText(run_time);
		Status.setText("Status of server... Server Off");
		DB_Manager.save();
		DB_Manager.exit();
	}
}