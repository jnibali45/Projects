import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.ResultSet;
import java.sql.SQLException;
import java.sql.Statement;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Collections;
import java.text.NumberFormat;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.util.Date;
import java.sql.Timestamp;

import java.awt.event.*;
import java.awt.*;
import java.awt.Color;

import javax.swing.*;

public class Stat_Tracker extends JFrame implements ActionListener {

	private static String monthSelected = "", password = "p4word";
	private static int place = 0;

	public static void mainMenu() {
		JFrame frame = new JFrame("Stats Program");
    	frame.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
    	frame.setSize(500, 200);
    	
    	JPanel panel = new JPanel();
    	panel.setBounds(10, 10, 475, 150);
    	panel.setBackground(Color.WHITE);
    	panel.setLayout(null);
    	
    	JButton button = new JButton("Select");
    	button.setBounds(frame.getWidth()/2 - 65, frame.getHeight()/2, 130, 30);
    	String months[] = {"January", "February", "March", "April", "May", "June", "July", 
    			"August", "September", "October", "November", "December"};
    	
    	JComboBox c = new JComboBox(months);
    	c.setBounds(frame.getWidth()/2 - 65, frame.getHeight()/2 - 40, 130, 30);
    	JLabel label = new JLabel("Select the desired month: ");
    	label.setBounds(frame.getWidth()/2 - 75, frame.getHeight()/2 - 80, 400, 30);
    	
    	button.addActionListener(new ActionListener() {
    		public void actionPerformed(ActionEvent e) {
    			String selected = "";
    			selected = (String) c.getSelectedItem();
    			monthSelected = selected;
    			System.out.println(monthSelected);
    			frame.setVisible(false);
    			try {
					finalMenu();
				} catch (IOException e1) {
					// TODO Auto-generated catch block
					e1.printStackTrace();
				}    			
    		}
    	});  	
    	
    	frame.add(panel);
    	frame.setLayout(null);
    	panel.add(c);
    	panel.add(button);
    	panel.add(label);
    	frame.setVisible(true);
    	frame.setResizable(false);
    	frame.setIconImage(createImage("/resources/cpark.png").getImage());
    	Dimension dim = Toolkit.getDefaultToolkit().getScreenSize();
    	frame.setLocation(dim.width/2-frame.getSize().width/2, dim.height/2-frame.getSize().height/2);
	}
	
	public static void passwordMenu() {
		JFrame frame = new JFrame("Stats Program");
    	frame.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
    	frame.setSize(500, 200);
    	
    	JPanel panel = new JPanel();
    	panel.setBounds(10, 10, 475, 150);
    	panel.setBackground(Color.WHITE);
    	panel.setLayout(null);
    	
    	JButton button = new JButton("Enter");
    	button.setBounds(frame.getWidth()/2 - 65, frame.getHeight()/2, 130, 30);
    	
    	JLabel label = new JLabel("Enter password to run stats: ");
    	label.setBounds(frame.getWidth()/2 - 75, frame.getHeight()/2 - 100, 400, 30);
    	
    	JLabel incorrect = new JLabel("Incorrect Password. Try Again.");
    	incorrect.setBounds(frame.getWidth()/2 - 80, frame.getHeight()/2 - 80, 400, 30);
    	incorrect.setForeground(Color.RED);
    	incorrect.setVisible(false);
    	
    	TextField text = new TextField(20);
    	text.setBounds(frame.getWidth()/2 - 65, frame.getHeight()/2 - 40, 130, 20);
    	
    	button.addActionListener(new ActionListener() {
    		public void actionPerformed(ActionEvent e) {
    			if(!(text.getText().equals(password))) {
    				incorrect.setVisible(true);
    				System.out.println(place);
    			} else {
    				frame.setVisible(false);
    				mainMenu();
    			}
    		}
    	}); 
    	
    	panel.add(incorrect);
    	frame.add(panel);
    	frame.setLayout(null);
    	panel.add(text);
    	panel.add(button);
    	panel.add(label);
    	frame.setVisible(true);
    	frame.setResizable(false);
    	Dimension dim = Toolkit.getDefaultToolkit().getScreenSize();
    	frame.setLocation(dim.width/2-frame.getSize().width/2, dim.height/2-frame.getSize().height/2);
    	
	}
	
	public static void finalMenu() throws IOException {
		JFrame frame = new JFrame("Stats Program");
    	frame.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
    	frame.setSize(500, 200);
    	
    	JPanel panel = new JPanel();
    	panel.setBounds(10, 10, 475, 150);
    	panel.setBackground(Color.WHITE);
    	panel.setLayout(null);
    	
    	JLabel label = new JLabel("Stats Run Successfully!");
    	label.setBounds(frame.getWidth()/2 - 75, frame.getHeight()/2 - 60, 400, 30);
    	
    	JLabel label2 = new JLabel("Sheet created at C:\\Users\\Jacob\\Desktop\\Fire\\Monthly Sheets\\" +
    	monthSelected);
    	label2.setBounds(frame.getWidth()/2 - 200, frame.getHeight()/2 - 40, 400, 30);
    	
    	frame.add(panel);
    	panel.add(label);
    	panel.add(label2);
    	frame.setLayout(null);
    	frame.setVisible(true);
    	frame.setResizable(false);
    	Dimension dim = Toolkit.getDefaultToolkit().getScreenSize();
    	frame.setLocation(dim.width/2-frame.getSize().width/2, dim.height/2-frame.getSize().height/2);
    	
    	runStats();
	}
	
	public static void runStats() throws IOException {
		// variables for connection and statement
		Connection connection = null;
		Statement statement = null;

		// variables for member, fire and ambo lists
		ResultSet members = null, fire_responses = null, ambo_responses = null;

		// set variables to true for them to print, false to hide print
		boolean monthly = true, specific = true, shamePrint = true;

		// Load/register JDBC driver class
		try {
			Class.forName("net.ucanaccess.jdbc.UcanaccessDriver");

		} catch (ClassNotFoundException cnfex) {
			System.out.println("Problem in loading or " + "registering MS Access JDBC driver");
			cnfex.printStackTrace();
		}

		// Open database connection
		try {
			// string connections for access file location
			String msAccDB = "C:/Users/Jacob/Desktop/Fire/2013_IRS.accdb";
			String dbURL = "jdbc:ucanaccess://" + msAccDB;

			// connecting to actual database
			connection = DriverManager.getConnection(dbURL);
			statement = connection.createStatement();

			// parsing members list
			members = statement.executeQuery("SELECT * FROM [CPVFD MEMBERS]");
			fire_responses = statement.executeQuery("SELECT * FROM [RESPONSES FIRE]");
			ambo_responses = statement.executeQuery("SELECT * FROM [RESPONSES AMBO]");

			System.out.println("================");
			System.out.println("= ANNUAL STATS =");
			System.out.println("================\n");
			System.out.println("Last Name\tFirst Name\tID\tFire Calls\tENG Driver\tTRK Driver\tENG Off.\tTRK Off."
					+ "\tCT Driver\tCT Officer\tStat\tEMS Calls\tAmbo Dr.\tTotal Stats");
			System.out.println("=========\t==========\t==\t==========\t==========\t==========\t========\t========"
					+ "\t=========\t==========\t====\t=========\t========\t===========");

			// arrays for member list (last, first name & ids)
			ArrayList<String> last_names = new ArrayList<String>();
			ArrayList<String> first_names = new ArrayList<String>();
			ArrayList<Integer> ids = new ArrayList<Integer>();
			ArrayList<Integer> stats = new ArrayList<Integer>();
			ArrayList<Integer> ems = new ArrayList<Integer>();
			ArrayList<Integer> fire = new ArrayList<Integer>();
			ArrayList<Integer> monthly_stats = new ArrayList<Integer>();
			ArrayList<Integer> monthly_ems = new ArrayList<Integer>();
			ArrayList<Integer> monthly_fire = new ArrayList<Integer>();
			ArrayList<Integer> eng_off = new ArrayList<Integer>();
			ArrayList<Integer> eng_dr = new ArrayList<Integer>();
			ArrayList<Integer> trk_off = new ArrayList<Integer>();
			ArrayList<Integer> trk_dr = new ArrayList<Integer>();
			ArrayList<Integer> ct_dr = new ArrayList<Integer>();
			ArrayList<Integer> ct_off = new ArrayList<Integer>();
			ArrayList<Integer> stat = new ArrayList<Integer>();
			ArrayList<Integer> ems_dr = new ArrayList<Integer>();
			ArrayList<Integer> monthly_trk_bk = new ArrayList<Integer>();
			ArrayList<Integer> monthly_eng_bk = new ArrayList<Integer>();
			ArrayList<Integer> monthly_ct_bk = new ArrayList<Integer>();
			ArrayList<Integer> monthly_am_bk = new ArrayList<Integer>();
			ArrayList<Integer> monthly_am_off = new ArrayList<Integer>();
			ArrayList<Integer> monthly_ct_dr = new ArrayList<Integer>();
			ArrayList<Integer> monthly_ct_off = new ArrayList<Integer>();
			ArrayList<Integer> monthly_stat = new ArrayList<Integer>();
			ArrayList<Integer> monthly_ems_dr = new ArrayList<Integer>();
			ArrayList<Integer> monthly_eng_off = new ArrayList<Integer>();
			ArrayList<Integer> monthly_eng_dr = new ArrayList<Integer>();
			ArrayList<Integer> monthly_trk_off = new ArrayList<Integer>();
			ArrayList<Integer> monthly_trk_dr = new ArrayList<Integer>();
			ArrayList<Integer> shame = new ArrayList<Integer>();
			ArrayList<Integer> monthly_shame = new ArrayList<Integer>();
			ArrayList<Integer> am_127_dr = new ArrayList<Integer>();

			NumberFormat defaultFormat = NumberFormat.getPercentInstance();
			defaultFormat.setMinimumFractionDigits(2);

			// loops through CPVFD members and adds data to array lists
			while (members.next()) {
				last_names.add(members.getString(1));
				first_names.add(members.getString(2));
				ids.add(members.getInt(3));
				stats.add(0);
				ems.add(0);
				fire.add(0);
				ct_dr.add(0);
				ct_off.add(0);
				stat.add(0);
				eng_dr.add(0);
				eng_off.add(0);
				ems_dr.add(0);
				trk_dr.add(0);
				trk_off.add(0);
				monthly_trk_bk.add(0);
				monthly_eng_bk.add(0);
				monthly_ct_bk.add(0);
				monthly_am_bk.add(0);
				monthly_am_off.add(0);
				monthly_ct_dr.add(0);
				monthly_ct_off.add(0);
				monthly_stat.add(0);
				monthly_ems_dr.add(0);
				monthly_eng_dr.add(0);
				monthly_eng_off.add(0);
				monthly_trk_dr.add(0);
				monthly_trk_off.add(0);
				monthly_stats.add(0);
				monthly_ems.add(0);
				monthly_fire.add(0);
				shame.add(0);
				monthly_shame.add(0);
				am_127_dr.add(0);
			}

			// ArrayList<String> dates = new ArrayList<String>();
			int fire_calls = 0, ems_calls = 0, transports = 0, monthly_fire_calls = 0, monthly_ems_calls = 0,
					working_fires = 0, box_street = 0, monthly_working_fires = 0, monthly_box_street = 0;
			ArrayList<String> fire_units = new ArrayList<String>(Arrays.asList("E121", "E122", "TK12", "Cart12",
					"Reserve Truck", "F12", "C12", "C12A", "C12B", "Stat", null));
			ArrayList<Integer> fire_units_count = new ArrayList<Integer>(Arrays.asList(0, 0, 0, 0, 0, 0, 0, 0, 0, 0));
			ArrayList<String> ems_units = new ArrayList<String>(
					Arrays.asList("A127", "A128", "A129", "A127 (M31203)", "PA812"));
			ArrayList<Integer> ems_units_count = new ArrayList<Integer>(Arrays.asList(0, 0, 0, 0, 0));
			ArrayList<Integer> ems_units_transports = new ArrayList<Integer>(Arrays.asList(0, 0, 0, 0, 0));
			ArrayList<Integer> fire_units_bs = new ArrayList<Integer>(Arrays.asList(0, 0, 0, 0, 0, 0));
			ArrayList<Integer> fire_units_work = new ArrayList<Integer>(Arrays.asList(0, 0, 0, 0, 0, 0));
			ArrayList<Integer> unit_ids = new ArrayList<Integer>();
			ArrayList<Integer> eng_off_ids = new ArrayList<Integer>();
			ArrayList<Integer> eng_dr_ids = new ArrayList<Integer>();
			ArrayList<Integer> trk_off_ids = new ArrayList<Integer>();
			ArrayList<Integer> trk_dr_ids = new ArrayList<Integer>();
			ArrayList<Integer> ems_dr_ids = new ArrayList<Integer>();
			ArrayList<Integer> shame_ids = new ArrayList<Integer>();
			ArrayList<Integer> ct_dr_ids = new ArrayList<Integer>();
			ArrayList<Integer> ct_off_ids = new ArrayList<Integer>();
			ArrayList<Integer> stat_ids = new ArrayList<Integer>();
			ArrayList<Integer> eng_bk_ids = new ArrayList<Integer>();
			ArrayList<Integer> trk_bk_ids = new ArrayList<Integer>();
			ArrayList<Integer> am_bk_ids = new ArrayList<Integer>();
			ArrayList<Integer> ct_bk_ids = new ArrayList<Integer>();
			ArrayList<Integer> am_off_ids = new ArrayList<Integer>();
			ArrayList<Integer> am_127_dr_ids = new ArrayList<Integer>();

			String temp = "";
			ArrayList<String> months = new ArrayList<String>(Arrays.asList("January", "February", "March", "April", 
					"May", "June", "July", "August", "September", "October", "November", "December"));
			while (fire_responses.next()) {
				temp = fire_responses.getString(2);

				unit_ids = new ArrayList<Integer>();
				eng_off_ids = new ArrayList<Integer>();
				eng_dr_ids = new ArrayList<Integer>();
				trk_off_ids = new ArrayList<Integer>();
				trk_dr_ids = new ArrayList<Integer>();
				shame_ids = new ArrayList<Integer>();
				ct_dr_ids = new ArrayList<Integer>();
				ct_off_ids = new ArrayList<Integer>();
				stat_ids = new ArrayList<Integer>();

				// Only adds calls for the year of 2019
				if (temp.substring(0, 4).equals("2019")) {
					fire_calls++;

					if (fire_responses.getBoolean("Box or Street Assignment") == true) {
						box_street++;
					}

					if (fire_responses.getBoolean("Working Fire") == true) {
						working_fires++;
					}

					// loops through the units, makes sure is in list of valid unit names
					// then adds an id for all 9 seats of the unit, if member is present in
					// seat then adds their unit id to unit_ids, if no member present then
					// will add 0
					for (int i = 1; i < 6; i++) {
						String uni = fire_responses.getString(11 * i);
						if (fire_units.contains(uni)) {
							// increments number of calls for specific apparatus
							if (uni != null) {
								fire_units_count.set(fire_units.indexOf(uni),
										fire_units_count.get(fire_units.indexOf(uni)) + 1);
							}
							for (int j = ((11 * i) + 1); j < ((11 * i) + 10); j++) {
								unit_ids.add(fire_responses.getInt(j));

								if (uni != null && (uni.equals("E121") || uni.equals("E122")) && j % 12 == 0) {
									eng_off_ids.add(fire_responses.getInt(j));
								} else if (uni != null && (uni.equals("TK12") || uni.equals("Reserve Truck"))
										&& j % 12 == 0) {
									trk_off_ids.add(fire_responses.getInt(j));
								} else if (uni != null && (uni.equals("E121") || uni.equals("E122")) && j % 13 == 0) {
									eng_dr_ids.add(fire_responses.getInt(j));
								} else if (uni != null && (uni.equals("TK12") || uni.equals("Reserve Truck"))
										&& j % 13 == 0) {
									trk_dr_ids.add(fire_responses.getInt(j));
								} else if (uni != null && (uni.equals("Cart12")) && j % 12 == 0) {
									ct_off_ids.add(fire_responses.getInt(j));
								} else if (uni != null && (uni.equals("Cart12")) && j % 13 == 0) {
									ct_dr_ids.add(fire_responses.getInt(j));
								} else if (uni != null && (uni.equals("Stat"))) {
									stat_ids.add(fire_responses.getInt(j));
								}
							}
						}
					}

					// loops through the temporary unit ids array list made for each call, checks
					// to make sure the id at index is an actual id (aka not 0), then checks the
					// member ids array list for that id number. After finding id number, uses index
					// of that id in ids array list to increment the stats number at same array
					// index
					// in stats array list by 1
					for (int i = 0; i < unit_ids.size(); i++) {
						int id = unit_ids.get(i);

						if (id != 0) {
							for (int j = 0; j < ids.size(); j++) {
								if (id == ids.get(j)) {
									stats.set(j, stats.get(j) + 1);
									fire.set(j, fire.get(j) + 1);
								}
							}
						}
					}

					// Adds truck officer stats
					for (int i = 0; i < trk_off_ids.size(); i++) {
						int id = trk_off_ids.get(i);

						if (id != 0) {
							for (int j = 0; j < ids.size(); j++) {
								if (id == ids.get(j)) {
									trk_off.set(j, trk_off.get(j) + 1);
								}
							}
						}
					}

					// Adds engine officer stats
					for (int i = 0; i < eng_off_ids.size(); i++) {
						int id = eng_off_ids.get(i);

						if (id != 0) {
							for (int j = 0; j < ids.size(); j++) {
								if (id == ids.get(j)) {
									eng_off.set(j, eng_off.get(j) + 1);
								}
							}
						}
					}

					// Adds truck driver stats
					for (int i = 0; i < trk_dr_ids.size(); i++) {
						int id = trk_dr_ids.get(i);

						if (id != 0) {
							for (int j = 0; j < ids.size(); j++) {
								if (id == ids.get(j)) {
									trk_dr.set(j, trk_dr.get(j) + 1);
								}
							}
						}
					}

					// Adds engine driverr stats
					for (int i = 0; i < eng_dr_ids.size(); i++) {
						int id = eng_dr_ids.get(i);

						if (id != 0 && id != 25370) {
							for (int j = 0; j < ids.size(); j++) {
								if (id == ids.get(j)) {
									eng_dr.set(j, eng_dr.get(j) + 1);
								}
							}
						}
					}

					// Adds cart officer stats
					for (int i = 0; i < ct_off_ids.size(); i++) {
						int id = ct_off_ids.get(i);

						if (id != 0) {
							for (int j = 0; j < ids.size(); j++) {
								if (id == ids.get(j)) {
									ct_off.set(j, ct_off.get(j) + 1);
								}
							}
						}
					}

					// Adds cart driver stats
					for (int i = 0; i < ct_dr_ids.size(); i++) {
						int id = ct_dr_ids.get(i);

						if (id != 0) {
							for (int j = 0; j < ids.size(); j++) {
								if (id == ids.get(j)) {
									ct_dr.set(j, ct_dr.get(j) + 1);
								}
							}
						}
					}

					// Adds stat stats
					for (int i = 0; i < stat_ids.size(); i++) {
						int id = stat_ids.get(i);

						if (id != 0) {
							for (int j = 0; j < ids.size(); j++) {
								if (id == ids.get(j)) {
									stat.set(j, stat.get(j) + 1);
								}
							}
						}
					}

				}
			}

			while (ambo_responses.next()) {
				temp = ambo_responses.getString(2);
				// Only adds calls for the year of 2019
				unit_ids = new ArrayList<Integer>();
				ems_dr_ids = new ArrayList<Integer>();
				am_127_dr_ids = new ArrayList<Integer>();

				if (temp != null && temp.substring(0, 4).equals("2019")) {
					ems_calls++;

					// checks ids for possible 4 seats on ambulance
					unit_ids.add(ambo_responses.getInt("Unit 1 Off ID"));
					unit_ids.add(ambo_responses.getInt("Unit 1 Dr ID"));
					ems_dr_ids.add(ambo_responses.getInt("Unit 1 Dr ID"));
					unit_ids.add(ambo_responses.getInt("Unit 1 FF1 ID"));
					unit_ids.add(ambo_responses.getInt("Unit 1 FF2 ID"));

					// increments call for specific apparatus
					String uni = ambo_responses.getString("Unit 1");
					if (uni != null) {
						ems_units_count.set(ems_units.indexOf(uni), ems_units_count.get(ems_units.indexOf(uni)) + 1);

						if (uni.equals("A127")) {
							am_127_dr_ids.add(ambo_responses.getInt("Unit 1 Dr ID"));
						}
					}

					// checks if transport and updates transport numbers + unit transports if needed
					String transport = ambo_responses.getString("Disposition");
					if (transport != null && transport.equals("Transport")) {
						transports++;
						ems_units_transports.set(ems_units.indexOf(uni),
								ems_units_transports.get(ems_units.indexOf(uni)) + 1);
					}
					// gives stat based on unit ids list
					for (int i = 0; i < unit_ids.size(); i++) {
						int id = unit_ids.get(i);
						if (id != 0) {
							for (int j = 0; j < ids.size(); j++) {
								if (id == ids.get(j)) {
									stats.set(j, stats.get(j) + 1);
									ems.set(j, ems.get(j) + 1);
								}
							}
						}
					}

					// Adds stat for ambo driver
					for (int i = 0; i < ems_dr_ids.size(); i++) {
						int id = ems_dr_ids.get(i);
						if (id != 0) {
							for (int j = 0; j < ids.size(); j++) {
								if (id == ids.get(j)) {
									ems_dr.set(j, ems_dr.get(j) + 1);
								}
							}
						}
					}

					for (int i = 0; i < am_127_dr_ids.size(); i++) {
						int id = am_127_dr_ids.get(i);
						if (id != 0) {
							for (int j = 0; j < ids.size(); j++) {
								if (id == ids.get(j)) {
									am_127_dr.set(j, am_127_dr.get(j) + 1);
								}
							}
						}
					}

				}
			}

			// Prints out the total annual stats for each member
			for (int i = 0; i < last_names.size(); i++) {
				if (stats.get(i) != 0) {
					if (last_names.get(i).length() <= 7 && first_names.get(i).length() <= 7) {
						System.out.println(last_names.get(i) + "\t" + first_names.get(i) + "\t" + ids.get(i) + " \t"
						/*
						 * + fire.get(i) + "\t\t" + eng_dr.get(i) + "\t\t" + trk_dr.get(i) + "\t\t" +
						 * eng_off.get(i) + "\t\t" + trk_off.get(i) + "\t\t" + ct_dr.get(i) + "\t\t" +
						 * ct_off.get(i) + "\t\t" + stat.get(i) + "\t\t" + ems.get(i) + "\t\t"
						 */
								+ ems_dr.get(i) + "\t " + am_127_dr.get(i) + "\t\t" + stats.get(i));
					} else if (last_names.get(i).length() <= 7 && first_names.get(i).length() > 7) {
						System.out.println(last_names.get(i) + "\t" + first_names.get(i) + "\t" + ids.get(i) + " \t"
						/*
						 * + fire.get(i) + "\t\t" + eng_dr.get(i) + "\t\t" + trk_dr.get(i) + "\t\t" +
						 * eng_off.get(i) + "\t\t" + trk_off.get(i) + "\t\t" + ct_dr.get(i) + "\t\t" +
						 * ct_off.get(i) + "\t\t" + stat.get(i) + "\t\t" + ems.get(i) + "\t\t"
						 */
								+ ems_dr.get(i) + "\t " + am_127_dr.get(i) + "\t\t" + stats.get(i));
					} else if (last_names.get(i).length() > 7 && first_names.get(i).length() <= 7) {
						System.out.println(last_names.get(i) + "\t" + first_names.get(i) + "\t" + ids.get(i) + " \t"
						/*
						 * + fire.get(i) + "\t\t" + eng_dr.get(i) + "\t\t" + trk_dr.get(i) + "\t\t" +
						 * eng_off.get(i) + "\t\t" + trk_off.get(i) + "\t\t" + ct_dr.get(i) + "\t\t" +
						 * ct_off.get(i) + "\t\t" + stat.get(i) + "\t\t" + ems.get(i) + "\t\t"
						 */
								+ ems_dr.get(i) + "\t " + am_127_dr.get(i) + "\t\t" + stats.get(i));
					} else {
						System.out.println(last_names.get(i) + "\t" + first_names.get(i) + "\t" + ids.get(i) + " \t"
						/*
						 * + fire.get(i) + "\t\t" + eng_dr.get(i) + "\t\t" + trk_dr.get(i) + "\t\t" +
						 * eng_off.get(i) + "\t\t" + trk_off.get(i) + "\t\t" + ct_dr.get(i) + "\t\t" +
						 * ct_off.get(i) + "\t\t" + stat.get(i) + "\t\t" + ems.get(i) + "\t\t"
						 */
								+ ems_dr.get(i) + "\t " + am_127_dr.get(i) + "\t\t" + stats.get(i));
					}
				}
			}

			if (specific) {
				System.out.println("\n==================");
				System.out.println("= SPECIFIC STATS =");
				System.out.println("==================\n");

				System.out.println("Total Fire Calls: " + fire_calls);
				System.out.println("Total Box/Street Alarms: " + box_street);
				System.out.println("Total Working Fires: " + working_fires);
				System.out.println("Total EMS Calls: " + ems_calls);
				System.out.println("Total Transports: " + transports);
				System.out.println("Total Station Calls: " + (ems_calls + fire_calls) + "\n");

				System.out.println("CALLS BY APPARATUS: \n");
				System.out.println("Apparatus\tCalls Run\tTransports (if valid)\tPercentage Transports");
				System.out.println("=========\t=========\t=====================\t=====================");

				for (int i = 0; i < fire_units_count.size(); i++) {
					if (fire_units.get(i).equals("Reserve Truck")) {
						System.out.println(fire_units.get(i) + "\t" + fire_units_count.get(i) + "\t\tNA\t\t\tNA");
					} else {
						System.out.println(fire_units.get(i) + "\t\t" + fire_units_count.get(i) + "\t\tNA\t\t\tNA");
					}
				}

				for (int i = 0; i < ems_units.size(); i++) {
					if (ems_units_count.get(i) != 0 && ems_units.get(i).equals("A127 (M31203)")) {
						double percentage = ((double) (ems_units_transports.get(i).intValue())
								/ (double) (ems_units_count.get(i)));
						System.out.println(ems_units.get(i) + "\t" + ems_units_count.get(i) + "\t\t"
								+ ems_units_transports.get(i) + "\t\t\t" + defaultFormat.format(percentage));
					} else if (ems_units_count.get(i) != 0) {
						double percentage = ((double) (ems_units_transports.get(i).intValue())
								/ (double) (ems_units_count.get(i)));
						System.out.println(ems_units.get(i) + "\t\t" + ems_units_count.get(i) + "\t\t"
								+ ems_units_transports.get(i) + "\t\t\t" + defaultFormat.format(percentage));
					} else {
						System.out.println(ems_units.get(i) + "\t\t" + ems_units_count.get(i) + "\t\t"
								+ ems_units_transports.get(i) + "\t\t\t" + "%NaN");
					}
				}
			}

			if (monthly) {
				int month = 0;
				
				for(int i = 0; i < months.size(); i++) {
					if(monthSelected.equals(months.get(i))) {
						month = (i + 1);
					}
				}				
				
				String check = "";

				fire_responses = statement.executeQuery("SELECT * FROM [RESPONSES FIRE]");
				ambo_responses = statement.executeQuery("SELECT * FROM [RESPONSES AMBO]");

				if (month < 10) {
					check = "2019-0";
					check += month;
				} else {
					check = "2019-";
					check += month;
				}

				while (fire_responses.next()) {
					temp = fire_responses.getString(2);
					unit_ids = new ArrayList<Integer>();
					eng_off_ids = new ArrayList<Integer>();
					eng_dr_ids = new ArrayList<Integer>();
					trk_off_ids = new ArrayList<Integer>();
					trk_dr_ids = new ArrayList<Integer>();
					ct_dr_ids = new ArrayList<Integer>();
					ct_off_ids = new ArrayList<Integer>();
					stat_ids = new ArrayList<Integer>();
					eng_bk_ids = new ArrayList<Integer>();
					trk_bk_ids = new ArrayList<Integer>();
					ct_bk_ids = new ArrayList<Integer>();
					shame_ids = new ArrayList<Integer>();

					if (temp != null && temp.substring(0, 7).equals(check)) {
						monthly_fire_calls++;

						if (fire_responses.getBoolean("Box or Street Assignment") == true) {
							monthly_box_street++;
						}

						if (fire_responses.getBoolean("Working Fire") == true) {
							monthly_working_fires++;
						}

						// loops through the units, makes sure is in list of valid unit names
						// then adds an id for all 9 seats of the unit, if member is present in
						// seat then adds their unit id to unit_ids, if no member present then
						// will add 0
						for (int i = 1; i < 6; i++) {
							String uni = fire_responses.getString(11 * i);
							if (fire_units.contains(uni)) {
								for (int j = ((11 * i) + 1); j < ((11 * i) + 10); j++) {
									unit_ids.add(fire_responses.getInt(j));

									if (uni != null && (uni.equals("E121") || uni.equals("E122")) && j % 12 == 0) {
										eng_off_ids.add(fire_responses.getInt(j));
									} else if (uni != null && (uni.equals("TK12") || uni.equals("Reserve Truck"))
											&& j % 12 == 0) {
										trk_off_ids.add(fire_responses.getInt(j));
									} else if (uni != null && (uni.equals("E121") || uni.equals("E122"))
											&& j % 13 == 0) {
										eng_dr_ids.add(fire_responses.getInt(j));
									} else if (uni != null && (uni.equals("TK12") || uni.equals("Reserve Truck"))
											&& j % 13 == 0) {
										trk_dr_ids.add(fire_responses.getInt(j));
									} else if (uni != null && (uni.equals("Cart12")) && j % 12 == 0) {
										ct_off_ids.add(fire_responses.getInt(j));
									} else if (uni != null && (uni.equals("Cart12")) && j % 13 == 0) {
										ct_dr_ids.add(fire_responses.getInt(j));
									} else if (uni != null && (uni.equals("Cart12"))
											&& ((j % 13 != 0) && (j % 12 != 0))) {
										ct_bk_ids.add(fire_responses.getInt(j));
									} else if (uni != null && (uni.equals("E121") || uni.equals("E122"))
											&& ((j % 13 != 0) && (j % 12 != 0))) {
										eng_bk_ids.add(fire_responses.getInt(j));
									} else if (uni != null && (uni.equals("TK12") || uni.equals("Reserve Truck"))
											&& ((j % 13 != 0) && (j % 12 != 0))) {
										trk_bk_ids.add(fire_responses.getInt(j));
									} else if (uni != null && (uni.equals("Stat"))) {
										stat_ids.add(fire_responses.getInt(j));
									}
								}
							} else if (uni != null && uni.equals("Slept Through")) {
								for (int j = ((11 * i) + 1); j < ((11 * i) + 10); j++) {
									shame_ids.add(fire_responses.getInt(j));
								}
							}
						}

						// loops through the temporary unit ids array list made for each call, checks
						// to make sure the id at index is an actual id (aka not 0), then checks the
						// member ids array list for that id number. After finding id number, uses index
						// of that id in ids array list to increment the stats number at same array
						// index
						// in stats array list by 1
						for (int i = 0; i < unit_ids.size(); i++) {
							int id = unit_ids.get(i);

							if (id != 0) {
								for (int j = 0; j < ids.size(); j++) {
									if (id == ids.get(j)) {
										monthly_stats.set(j, monthly_stats.get(j) + 1);
										monthly_fire.set(j, monthly_fire.get(j) + 1);
									}
								}
							}
						}

						// Adds shame stat
						for (int i = 0; i < shame_ids.size(); i++) {
							int id = shame_ids.get(i);

							if (id != 0 && id != 25370) {
								for (int j = 0; j < ids.size(); j++) {
									if (id == ids.get(j)) {
										shame.set(j, shame.get(j) + 1);
									}
								}
							}
						}

						// Adds monthly eng officer stats
						for (int i = 0; i < eng_off_ids.size(); i++) {
							int id = eng_off_ids.get(i);

							if (id != 0) {
								for (int j = 0; j < ids.size(); j++) {
									if (id == ids.get(j)) {
										monthly_eng_off.set(j, monthly_eng_off.get(j) + 1);
									}
								}
							}
						}

						// Adds monthly eng driver stats
						for (int i = 0; i < eng_dr_ids.size(); i++) {
							int id = eng_dr_ids.get(i);

							if (id != 0) {
								for (int j = 0; j < ids.size(); j++) {
									if (id == ids.get(j)) {
										monthly_eng_dr.set(j, monthly_eng_dr.get(j) + 1);
									}
								}
							}
						}

						// Adds monthly trk officer stats
						for (int i = 0; i < trk_off_ids.size(); i++) {
							int id = trk_off_ids.get(i);

							if (id != 0) {
								for (int j = 0; j < ids.size(); j++) {
									if (id == ids.get(j)) {
										monthly_trk_off.set(j, monthly_trk_off.get(j) + 1);
									}
								}
							}
						}

						// Adds monthly trk driver stats
						for (int i = 0; i < trk_dr_ids.size(); i++) {
							int id = trk_dr_ids.get(i);

							if (id != 0) {
								for (int j = 0; j < ids.size(); j++) {
									if (id == ids.get(j)) {
										monthly_trk_dr.set(j, monthly_trk_dr.get(j) + 1);
									}
								}
							}
						}

						// Adds cart officer stats
						for (int i = 0; i < ct_off_ids.size(); i++) {
							int id = ct_off_ids.get(i);

							if (id != 0) {
								for (int j = 0; j < ids.size(); j++) {
									if (id == ids.get(j)) {
										monthly_ct_off.set(j, monthly_ct_off.get(j) + 1);
									}
								}
							}
						}

						// Adds cart driver stats
						for (int i = 0; i < ct_dr_ids.size(); i++) {
							int id = ct_dr_ids.get(i);

							if (id != 0) {
								for (int j = 0; j < ids.size(); j++) {
									if (id == ids.get(j)) {
										monthly_ct_dr.set(j, monthly_ct_dr.get(j) + 1);
									}
								}
							}
						}

						// Adds stat stats
						for (int i = 0; i < stat_ids.size(); i++) {
							int id = stat_ids.get(i);

							if (id != 0) {
								for (int j = 0; j < ids.size(); j++) {
									if (id == ids.get(j)) {
										monthly_stat.set(j, monthly_stat.get(j) + 1);
									}
								}
							}
						}

						// Adds monthly eng back stats
						for (int i = 0; i < eng_bk_ids.size(); i++) {
							int id = eng_bk_ids.get(i);

							if (id != 0) {
								for (int j = 0; j < ids.size(); j++) {
									if (id == ids.get(j)) {
										monthly_eng_bk.set(j, monthly_eng_bk.get(j) + 1);
									}
								}
							}
						}

						// Adds monthly truck back stats
						for (int i = 0; i < trk_bk_ids.size(); i++) {
							int id = trk_bk_ids.get(i);

							if (id != 0) {
								for (int j = 0; j < ids.size(); j++) {
									if (id == ids.get(j)) {
										monthly_trk_bk.set(j, monthly_trk_bk.get(j) + 1);
									}
								}
							}
						}

						// Adds monthly cart back stats
						for (int i = 0; i < ct_bk_ids.size(); i++) {
							int id = ct_bk_ids.get(i);

							if (id != 0) {
								for (int j = 0; j < ids.size(); j++) {
									if (id == ids.get(j)) {
										monthly_ct_bk.set(j, monthly_ct_bk.get(j) + 1);
									}
								}
							}
						}
					}
				}

				while (ambo_responses.next()) {
					temp = ambo_responses.getString(2);
					// Only adds calls for the year of 2019
					unit_ids = new ArrayList<Integer>();
					ems_dr_ids = new ArrayList<Integer>();
					am_bk_ids = new ArrayList<Integer>();
					am_off_ids = new ArrayList<Integer>();

					if (temp != null && temp.substring(0, 7).equals(check)) {
						monthly_ems_calls++;

						// checks ids for possible 4 seats on ambulance

						unit_ids.add(ambo_responses.getInt("Unit 1 Off ID"));
						unit_ids.add(ambo_responses.getInt("Unit 1 Dr ID"));
						unit_ids.add(ambo_responses.getInt("Unit 1 FF1 ID"));
						unit_ids.add(ambo_responses.getInt("Unit 1 FF2 ID"));

						ems_dr_ids.add(ambo_responses.getInt("Unit 1 Dr ID"));
						am_off_ids.add(ambo_responses.getInt("Unit 1 Off ID"));
						am_bk_ids.add(ambo_responses.getInt("Unit 1 FF1 ID"));
						am_bk_ids.add(ambo_responses.getInt("Unit 1 FF2 ID"));

						// gives stat based on unit ids list
						for (int i = 0; i < unit_ids.size(); i++) {
							int id = unit_ids.get(i);
							if (id != 0) {
								for (int j = 0; j < ids.size(); j++) {
									if (id == ids.get(j)) {
										monthly_stats.set(j, monthly_stats.get(j) + 1);
										monthly_ems.set(j, monthly_ems.get(j) + 1);
									}
								}
							}
						}

						for (int i = 0; i < ems_dr_ids.size(); i++) {
							int id = ems_dr_ids.get(i);
							if (id != 0) {
								for (int j = 0; j < ids.size(); j++) {
									if (id == ids.get(j)) {
										monthly_ems_dr.set(j, monthly_ems_dr.get(j) + 1);
									}
								}
							}
						}

						for (int i = 0; i < am_bk_ids.size(); i++) {
							int id = am_bk_ids.get(i);
							if (id != 0) {
								for (int j = 0; j < ids.size(); j++) {
									if (id == ids.get(j)) {
										monthly_am_bk.set(j, monthly_am_bk.get(j) + 1);
									}
								}
							}
						}

						for (int i = 0; i < am_off_ids.size(); i++) {
							int id = am_off_ids.get(i);
							if (id != 0) {
								for (int j = 0; j < ids.size(); j++) {
									if (id == ids.get(j)) {
										monthly_am_off.set(j, monthly_am_off.get(j) + 1);
									}
								}
							}
						}

					}
				}

				// Rankings for the month
				ArrayList<Integer> active_month = new ArrayList<Integer>();
				for (int i = 0; i < monthly_stats.size(); i++) {
					if (monthly_stats.get(i) != 0) {
						active_month.add(monthly_stats.get(i));
					}
				}

				float month_rank[] = new float[active_month.size()];

				for (int i = 0; i < active_month.size(); i++) {
					int r = 1, s = 1;

					for (int j = 0; j < active_month.size(); j++) {
						if (j != i && active_month.get(j) > active_month.get(i)) {
							r += 1;
						}

						if (j != i && active_month.get(j) == active_month.get(i)) {
							s += 1;
						}
					}

					month_rank[i] = r + (float) (s - 1) / (float) 2;
				}

				// Rankings for the year
				float year_rank_all[] = new float[stats.size()];

				for (int i = 0; i < stats.size(); i++) {
					int r = 1, s = 1;

					for (int j = 0; j < stats.size(); j++) {
						if (j != i && stats.get(j) > stats.get(i)) {
							r += 1;
						}

						if (j != i && stats.get(j) == stats.get(i)) {
							s += 1;
						}
					}

					year_rank_all[i] = r + (float) (s - 1) / (float) 2;
				}

				System.out.println("\n=========================");
				System.out.println("= MONTHLY STATS: " + month + "/2019 =");
				System.out.println("=========================\n");
				System.out.println("Monthly Fire Calls: " + monthly_fire_calls);
				System.out.println("Monthly Box/Street: " + monthly_box_street);
				System.out.println("Monthly Working Fires: " + monthly_working_fires);
				System.out.println("Monthly EMS Calls: " + monthly_ems_calls);
				System.out.println("Total Monthly Calls: " + (monthly_fire_calls + monthly_ems_calls));
				System.out
						.println("Percent Annual Calls: "
								+ defaultFormat.format(
										(double) (monthly_fire_calls + monthly_ems_calls) / (fire_calls + ems_calls))
								+ "\n");
				System.out.println("Last Name\tFirst Name\tID\tFire Calls\tENG Driver\tTRK Driver\tENG Off.\tTRK Off."
						+ "\tCT Driver\tCT Officer\tStat\tEMS Calls\tAmbo Dr.\tTotal Mo. Stat\tMonth Percent\tMO. Rank");
				System.out.println("=========\t==========\t==\t==========\t==========\t==========\t========\t========"
						+ "\t=========\t==========\t====\t=========\t========\t==============\t=============\n");
				int id_num = 0;
				int idx = 0;
				for (int i = 0; i < last_names.size(); i++) {

					if (monthly_stats.get(i) != 0) {
						id_num = ids.get(i);
						System.out.println(last_names.get(i) + "\t" + first_names.get(i) + "\t" + ids.get(i) + "\t"
								+ monthly_trk_dr.get(i) + "\t" + monthly_trk_off.get(i) + "\t" + monthly_trk_bk.get(i)
								+ "\t" + monthly_eng_dr.get(i) + "\t" + monthly_eng_off.get(i) + "\t"
								+ monthly_eng_bk.get(i) + "\t" + monthly_ct_dr.get(i) + "\t" + monthly_ct_off.get(i)
								+ "\t" + monthly_ct_bk.get(i) + "\t" + monthly_ems_dr.get(i) + "\t"
								+ monthly_am_off.get(i) + "\t" + monthly_am_bk.get(i) + "\t" + monthly_stat.get(i)
								+ "\t" + monthly_fire.get(i) + "\t" + monthly_ems.get(i) + "\t" + shame.get(i) + "\t"
								+ monthly_stats.get(i) + "\t" + stats.get(i) + "\t"
								+ defaultFormat.format(
										(double) monthly_stats.get(i) / (monthly_fire_calls + monthly_ems_calls))
								+ "\t" + month_rank[idx] + "\t" + year_rank_all[i]);
						idx++;
					}
				}

				// Creates workbook
				Workbook workbook = new XSSFWorkbook();

				// Creation Helper to assist with header formatting
				CreationHelper createHelper = workbook.getCreationHelper();
				Sheet sheet = workbook.createSheet("Employee");

				Font headerFont = workbook.createFont();
				headerFont.setBold(true);

				CellStyle headerCellStyle = workbook.createCellStyle();
				headerCellStyle.setFont(headerFont);

				// Create a Header Row *****TRY CHANGE TO 1
				Row headerRow = sheet.createRow(0);

				// Array List of header titles
				ArrayList<String> header = new ArrayList<String>(Arrays.asList("Last Name", "First Name", "ID", "TKD",
						"TKO", "TKB", "END", "ENO", "ENB", "CTD", "CTO", "CTB", "AMD", "AMO", "AMB", "STAT", "FIRE",
						"EMS", "SHAME", "MON", "YR", "MO %", "MO RANK", "YR RANK"));

				// Create cells for header ******TRY CHANGE TO 1 AND ICRMNT REST
				for (int i = 0; i < header.size(); i++) {
					Cell cell = headerRow.createCell(i);
					cell.setCellValue(header.get(i));
					cell.setCellStyle(headerCellStyle);
				}

				// Inputs data to spreadsheet
				idx = 0;
				for (int i = 0; i < last_names.size(); i++) {
					if (monthly_stats.get(i) != 0) {
						Row row = sheet.createRow(idx + 1);

						row.createCell(0).setCellValue(last_names.get(i));
						row.createCell(1).setCellValue(first_names.get(i));
						row.createCell(2).setCellValue(ids.get(i));
						row.createCell(3).setCellValue(monthly_trk_dr.get(i));
						row.createCell(4).setCellValue(monthly_trk_off.get(i));
						row.createCell(5).setCellValue(monthly_trk_bk.get(i));
						row.createCell(6).setCellValue(monthly_eng_dr.get(i));
						row.createCell(7).setCellValue(monthly_eng_off.get(i));
						row.createCell(8).setCellValue(monthly_eng_bk.get(i));
						row.createCell(9).setCellValue(monthly_ct_dr.get(i));
						row.createCell(10).setCellValue(monthly_ct_off.get(i));
						row.createCell(11).setCellValue(monthly_ct_bk.get(i));
						row.createCell(12).setCellValue(monthly_ems_dr.get(i));
						row.createCell(13).setCellValue(monthly_am_off.get(i));
						row.createCell(14).setCellValue(monthly_am_bk.get(i));
						row.createCell(15).setCellValue(monthly_stat.get(i));
						row.createCell(16).setCellValue(monthly_fire.get(i));
						row.createCell(17).setCellValue(monthly_ems.get(i));
						row.createCell(18).setCellValue(shame.get(i));
						row.createCell(19).setCellValue(monthly_stats.get(i));
						row.createCell(20).setCellValue(stats.get(i));
						row.createCell(21).setCellValue(defaultFormat
								.format((double) monthly_stats.get(i) / (monthly_fire_calls + monthly_ems_calls)));
						row.createCell(22).setCellValue((double) month_rank[idx]);
						row.createCell(23).setCellValue((double) year_rank_all[i]);

						idx++;
					}
				}

				// Resize all columns to fit the content size
				for (int i = 0; i < header.size(); i++) {
					sheet.autoSizeColumn(i);
				}

				Date date = new Date();
				long time = date.getTime();
				Timestamp ts = new Timestamp(time);
				String tempExt = ts.toString();
				String fileExt = "";

				for (int i = 0; i < tempExt.length(); i++) {
					if (tempExt.charAt(i) != ':') {
						fileExt += tempExt.charAt(i);
					}
				}

				// Write the output to a file
				
				FileOutputStream fileOut = new FileOutputStream("C:/Users/Jacob/Desktop/Fire/Monthly Sheets/" + 
				months.get(month - 1) + "/STATS " + months.get(month - 1) + " (" + fileExt + ").xlsx");
				workbook.write(fileOut);
				fileOut.close();

				// Closing the workbook
				workbook.close();
			}

			if (shamePrint) {

				System.out.println("\n=======================");
				System.out.println("= SLEPT THROUGH STATS =");
				System.out.println("=======================\n");
				System.out.println("Last Name\tFirst Name\tID\tCalls Slept Through");
				System.out.println("=========\t==========\t==\t===================");

				for (int i = 0; i < last_names.size(); i++) {
					if (shame.get(i) != 0) {
						if (last_names.get(i).length() <= 7 && first_names.get(i).length() <= 7) {
							System.out.println(last_names.get(i) + "\t\t" + first_names.get(i) + "\t\t" + ids.get(i)
									+ "\t\t" + shame.get(i));
						} else if (last_names.get(i).length() <= 7 && first_names.get(i).length() > 7) {
							System.out.println(last_names.get(i) + "\t\t" + first_names.get(i) + "\t" + ids.get(i)
									+ "\t\t" + shame.get(i));
						} else if (last_names.get(i).length() > 7 && first_names.get(i).length() <= 7) {
							System.out.println(last_names.get(i) + "\t" + first_names.get(i) + "\t\t" + ids.get(i)
									+ "\t\t" + shame.get(i));
						} else {
							System.out.println(last_names.get(i) + "\t" + first_names.get(i) + "\t" + ids.get(i)
									+ "\t\t" + shame.get(i));
						}

					}
				}
			}

		} catch (SQLException sqlex) {
			sqlex.printStackTrace();
		} finally {
			try {
				if (null != connection) {
					members.close();
					fire_responses.close();
					ambo_responses.close();

					connection.close();
				}
			} catch (SQLException sqlex) {
				sqlex.printStackTrace();
			}
		}
	}

	public static void main(String[] args) {
		passwordMenu();
	}

	@Override
	public void actionPerformed(ActionEvent e) {
		// TODO Auto-generated method stub
		
	}

}
