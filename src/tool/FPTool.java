package tool;

import java.awt.*;
import java.awt.Dialog.ModalExclusionType;
import javax.swing.*;
import javax.swing.filechooser.FileNameExtensionFilter;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.awt.event.ActionListener;
import java.awt.event.KeyEvent;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.math.BigDecimal;
import java.net.URL;
import java.util.*;
import java.util.List;
import java.awt.event.ActionEvent;
import java.awt.event.ItemEvent;
import java.awt.event.ItemListener;

public class FPTool extends JFrame {

//-------------------------------------------------------VARIABLE-----------------------------------------------------------------------------
	private JPanel contentPane;
	private JLabel ufc, background, vaf, result;
	JTable table;
	JTextField textfp, textvaf, textufc;
	JTextField textloc[] = new JTextField[11];
	ButtonGroup group2[] = new ButtonGroup[15];
	CheckboxGroup cbg = new CheckboxGroup();
	JButton btnCal = new JButton("Calculate");
	JButton btnReset = new JButton("Reset");
	JButton btnSave = new JButton("Save Result");
	JButton btnBrowse = new JButton("Browse...");
	JButton btnDecre[][] = new JButton[6][4];
	JButton btnIncre[][] = new JButton[6][4];
	JRadioButton radLow, radAve, radHi;
	int factor[][] = new int[6][4];
	int number[][] = new int[6][4];
	int resultUFC[][] = new int[6][4];
	boolean check = true;
	JRadioButton selectVAF[][] = new JRadioButton[15][7];
	Checkbox selectAll[] = new Checkbox[7];
	int resultF[] = new int[15];
	final TextField input[][] = new TextField[6][4];
	int sloc[] = new int[11];
	String excelFilePath = null;
	int countstep = 0;

	double ufcV = 0, vafV = 0, resultV = 0;
	int sum = 0, hang = 1, b = 0, cot = 1, c = 0;
	int step = 0;
	int dem = 0;
	int demn = 0;

	public static final int COLUMN_INDEX_ELEMENT = 0;
	public static final int COLUMN_INDEX_L = 1;
	public static final int COLUMN_INDEX_A = 2;
	public static final int COLUMN_INDEX_H = 3;
	public static final int COLUMN_INDEX_S = 4;

	private static CellStyle cellStyleFormatNumber = null;
	Workbook workbook;
	List<FunctionPoint> books;

	WriteExcel w = new WriteExcel();

	public FPTool() throws IOException {
		URL url = FPTool.class.getResource("/input.txt");
		InputStream input = url.openStream();
		Scanner sc = new Scanner(input);
		for (int i = 1; i <= 5; i++)
			for (int j = 1; j <= 3; j++)
				factor[i][j] = sc.nextInt();
		sc.close();
		URL urlloc = FPTool.class.getResource("/sloc.txt");
		InputStream inputloc = urlloc.openStream();
		sc = new Scanner(inputloc);
		for (int i = 1; i <= 10; i++)
			sloc[i] = sc.nextInt();
		sc.close();
		prepareGUI();

	}

	// Launch the application.
	public static void main(String[] args) throws IOException {
		FPTool frame = new FPTool();
		frame.setExtendedState(JFrame.MAXIMIZED_BOTH); // set JFrame full screen
		frame.setVisible(true);

	}

//-------------------------------------------------------GUI-----------------------------------------------------------------------------

	public void prepareGUI() {

		setModalExclusionType(ModalExclusionType.APPLICATION_EXCLUDE);
		setTitle("FunctionPointTool_462NIS_Group4");
		setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
		setBounds(0, 0, 1366, 768);
		contentPane = new JPanel();
		setContentPane(contentPane);
		contentPane.setLayout(null);
		contentPane.setBackground(SystemColor.white);

		// Set background
		background = new JLabel();
		contentPane.add(background);
		background.setSize(1376, 742);

		URL urlI;
		Icon table;
		urlI = FPTool.class.getResource("/1.png");
		table = new ImageIcon(urlI);
		ufc = new JLabel(table);
		ufc.setBounds(0, 0, 540, 349);
		ufc.setIcon(table);
		ufc.setVisible(true);

		urlI = FPTool.class.getResource("/2.png");
		table = new ImageIcon(urlI);
		vaf = new JLabel(table);
		vaf.setBounds(520, 11, 810, 702);
		vaf.setIcon(table);
		vaf.setVisible(true);

		urlI = FPTool.class.getResource("/3.png");
		table = new ImageIcon(urlI);
		result = new JLabel(table);
		result.setBounds(13, 429, 505, 265);
		result.setIcon(table);
		background.add(result);
		result.setVisible(true);

		btnCal.setBounds(120, 370, 100, 40);
		btnCal.setBackground(SystemColor.white);
		btnCal.setFont(new Font("Times New Roman", Font.PLAIN, 16));
		background.add(btnCal);

		btnReset.setFont(new Font("Times New Roman", Font.PLAIN, 16));
		btnReset.setBackground(SystemColor.white);
		btnReset.setBounds(20, 370, 75, 40);
		background.add(btnReset);

		btnSave.setFont(new Font("Times New Roman", Font.PLAIN, 16));
		btnSave.setBackground(SystemColor.white);
		btnSave.setBounds(245, 370, 120, 40);
		background.add(btnSave);
		btnSave.setEnabled(false);

		btnBrowse.setFont(new Font("Times New Roman", Font.PLAIN, 16));
		btnBrowse.setBackground(SystemColor.white);
		btnBrowse.setBounds(390, 370, 120, 40);
		background.add(btnBrowse);

		URL urlI1 = FPTool.class.getResource("/triangle.png");
		ImageIcon icon = new ImageIcon(urlI1);
		URL urlI2 = FPTool.class.getResource("/triangle1.png");
		ImageIcon icon1 = new ImageIcon(urlI2);
		b = 0;
		c = 0;
		for (int i = 1; i <= 5; i++) {
			for (int j = 1; j <= 3; j++) {

				btnDecre[i][j] = new JButton(icon);
				btnDecre[i][j].setBackground(SystemColor.white);
				btnDecre[i][j].setBounds(310 + c, 105 + b, 10, 10);
				btnDecre[i][j].setBorderPainted(false);
				background.add(btnDecre[i][j]);

				btnIncre[i][j] = new JButton(icon1);
				btnIncre[i][j].setBackground(SystemColor.white);
				btnIncre[i][j].setBounds(310 + c, 87 + b, 10, 10);
				btnIncre[i][j].setBorderPainted(false);
				background.add(btnIncre[i][j]);
				c += 90;
			}
			b += 54;
			c = 0;
		}

		textufc = new JTextField();
		textufc.setBounds(142, 461, 122, 26);
		textufc.setBackground(SystemColor.white);
		textufc.setEditable(false);
		textufc.setHorizontalAlignment(SwingConstants.CENTER);
		textufc.setFont(new Font("Times New Roman", Font.PLAIN, 18));
		background.add(textufc);

		textvaf = new JTextField();
		textvaf.setBounds(393, 461, 122, 26);
		textvaf.setBackground(SystemColor.white);
		textvaf.setEditable(false);
		textvaf.setHorizontalAlignment(SwingConstants.CENTER);
		textvaf.setFont(new Font("Times New Roman", Font.PLAIN, 18));
		background.add(textvaf);

		textfp = new JTextField();
		textfp.setBounds(129, 61, 373, 26);
		textfp.setBackground(SystemColor.white);
		textfp.setEditable(false);
		textfp.setHorizontalAlignment(SwingConstants.CENTER);
		textfp.setFont(new Font("Times New Roman", Font.PLAIN, 18));
		b = 0;
		c = 0;
		for (int i = 1; i <= 10; i++) {
			textloc[i] = new JTextField();
			textloc[i].setBounds(142 + b, 549 + c, 122, 26);
			textloc[i].setBackground(SystemColor.white);
			textloc[i].setEditable(false);
			textloc[i].setHorizontalAlignment(SwingConstants.CENTER);
			textloc[i].setFont(new Font("Times New Roman", Font.PLAIN, 18));
			if (i % 2 == 0) {
				b = 0;
				c += 29;
			} else
				b += 251;
		}

		// UFC select
		b = 0;
		c = 0;
		for (int i = 1; i <= 5; i++) {
			for (int j = 1; j <= 3; j++) {
				input[i][j] = new TextField(4);
				input[i][j].setBackground(SystemColor.white);
				input[i][j].setVisible(true);
				input[i][j].setBounds(260 + c, 90 + b, 50, 24);
				input[i][j].setFont(new Font("Times New Roman", Font.PLAIN, 16));
				input[i][j].setText("0");
				background.add(input[i][j]);
				c += 90;

			}
			b += 55;
			c = 0;
		}

		background.add(ufc);

		// VAF select
		c = 0;
		for (int i = 1; i <= 6; i++) {
			if (i == 1)
				selectAll[i] = new Checkbox("", cbg, true);
			else
				selectAll[i] = new Checkbox("", cbg, false);
			selectAll[i].setBounds(998 + c, 46, 20, 15);
			selectAll[i].setBackground(Color.decode("#f52549"));
			background.add(selectAll[i]);
			c += 59;
		}
		b = 0;
		c = 0;
		hang = 1;
		cot = 1;
		while (hang <= 14) {
			while (cot <= 6) {
				selectVAF[hang][cot] = new JRadioButton("", true);
				selectVAF[hang][cot].setMnemonic(KeyEvent.VK_C);
				selectVAF[hang][cot].setBounds(996 + c, 80 + b, 20, 15);
				selectVAF[hang][cot].setBackground(SystemColor.white);

				if (cot == 6) {
					group2[hang] = new ButtonGroup();
					for (int i = 1; i <= 6; i++) {
						group2[hang].add(selectVAF[hang][i]);
						background.add(selectVAF[hang][i]);
						selectVAF[hang][i].setVisible(true);
					}
				}
				c += 59;
				cot++;
			}
			b += 45;
			hang++;
			cot = 1;
			c = 0;
		}
		background.add(vaf);

		selectAll[1].addItemListener(new ItemListener() {
			public void itemStateChanged(ItemEvent e) {
				if (e.getStateChange() == 1)
					for (int i = 1; i <= 14; i++) {
						selectVAF[i][1].setSelected(true);
					}
			}
		});
		selectAll[2].addItemListener(new ItemListener() {
			public void itemStateChanged(ItemEvent e) {
				if (e.getStateChange() == 1)
					for (int i = 1; i <= 14; i++) {
						selectVAF[i][2].setSelected(true);
					}
			}
		});
		selectAll[3].addItemListener(new ItemListener() {
			public void itemStateChanged(ItemEvent e) {
				if (e.getStateChange() == 1)
					for (int i = 1; i <= 14; i++) {
						selectVAF[i][3].setSelected(true);
					}
			}
		});
		selectAll[4].addItemListener(new ItemListener() {
			public void itemStateChanged(ItemEvent e) {
				if (e.getStateChange() == 1)
					for (int i = 1; i <= 14; i++) {
						selectVAF[i][4].setSelected(true);
					}
			}
		});
		selectAll[5].addItemListener(new ItemListener() {
			public void itemStateChanged(ItemEvent e) {
				if (e.getStateChange() == 1)
					for (int i = 1; i <= 14; i++) {
						selectVAF[i][5].setSelected(true);
					}
			}
		});
		selectAll[6].addItemListener(new ItemListener() {
			public void itemStateChanged(ItemEvent e) {
				if (e.getStateChange() == 1)
					for (int i = 1; i <= 14; i++) {
						selectVAF[i][6].setSelected(true);
					}
			}
		});

//-------------------------------------------------------BUTTON-----------------------------------------------------------------------------

		btnCal.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent arg0) {
				for (int i = 1; i <= 5; i++)
					for (int j = 1; j <= 3; j++) {
						if (input[i][j].getText().isEmpty()) {
							number[i][j] = 0;
							input[i][j].setText("0");
						} else
							try {
								number[i][j] = Integer.parseInt(input[i][j].getText());
								if (number[i][j] < 0) {
									check = false;
								}
							} catch (Exception e) {
								check = false;
							}
					}
				double dou = 0.0;
				if (check == false) {
					JOptionPane.showMessageDialog(background,
							"Please enter valid number, select complexity weighting factor and scale!");
					textfp.setText(null);
				} else {
					for (int i = 1; i <= 5; i++) {
						for (int j = 1; j <= 3; j++)
							if (number[i][j] != 0) {
								resultUFC[i][j] = number[i][j] * factor[i][j];
								ufcV += resultUFC[i][j];
							}
					}

					for (int i = 1; i <= 14; i++)
						for (int j = 1; j <= 6; j++)
							if (selectVAF[i][j].isSelected()) {
								resultF[i] = j - 1;
								sum += j - 1;
							}
					vafV = 0.65 + 0.01 * sum;
					vafV = (double) Math.round(vafV * 1000) / 1000;
					resultV = ufcV * vafV;
					textufc.setText(ufcV + "");
					textvaf.setText(vafV + "");
					dou = (double) Math.round(resultV * 1000) / 1000;
					textfp.setText(String.valueOf(dou));
					btnSave.setEnabled(true);
				}
				result.add(textfp);
				textfp.setVisible(true);
				for (int i = 1; i <= 10; i++) {
					textloc[i].setText((int) Math.ceil((double) dou * sloc[i]) + "");
					background.add(textloc[i]);
				}
				sum = 0;
				ufcV = 0;
				demn = 0;
				dem = 0;
			}

		});

		btnReset.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent e) {
				for (int i = 1; i <= 5; i++)
					for (int j = 1; j <= 3; j++) {
						input[i][j].setText("0");
						number[i][j] = 0;
						resultUFC[i][j] = 0;
						input[i][j].setEditable(true);
					}

				sum = 0;
				ufcV = 0;
				textfp.setText(null);
				textufc.setText(null);
				textvaf.setText(null);
				for (int i = 1; i <= 10; i++)
					textloc[i].setText(null);
				for (int i = 1; i <= 14; i++)
					selectVAF[i][1].setSelected(true);
				selectAll[1].setState(true);
				check = true;
				dem = 0;
				demn = 0;
				btnSave.setEnabled(false);
			}
		});

		btnSave.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent arg0) {
				try {
					UIManager.setLookAndFeel(UIManager.getSystemLookAndFeelClassName());
				} catch (Exception e) {
				}
				SwingUtilities.invokeLater(new Runnable() {
					public void run() {
						try {
							showSaveFileDialog();
							//btnSave.setEnabled(false);
							books = getFunctionPoint();
							step++;
							writeExcel(books, excelFilePath);
							excelFilePath = null;
						} catch (IOException e) {
							JOptionPane.showMessageDialog(background, "File is in other process. Please close it!");
						}
					}
				});
			}
		});

		btnBrowse.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent arg0) {
				try {
					UIManager.setLookAndFeel(UIManager.getSystemLookAndFeelClassName());
				} catch (Exception e) {
				}
				SwingUtilities.invokeLater(new Runnable() {
					public void run() {
						try {
							excelFilePath = showOpenFileDialog();
							books = readExcel(excelFilePath);
							excelFilePath = null;
						} catch (IOException e) {
							System.out.println("Cancel!");
						}
					}
				});
			}
		});

		btnDecre[1][1].addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent e) {
				try {
					number[1][1] = Integer.parseInt(input[1][1].getText());
					if (number[1][1] < 0) {
						check = false;
					}
				} catch (Exception a) {
					check = false;
				}

				if (check == false) {
					JOptionPane.showMessageDialog(background, "Please enter valid number");
					check = true;
				}

				else {
					number[1][1]--;
					if (number[1][1] < 0)
						number[1][1] = 0;
					input[1][1].setText(number[1][1] + "");
				}
			}
		});
		btnIncre[1][1].addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent e) {
				try {
					number[1][1] = Integer.parseInt(input[1][1].getText());
					if (number[1][1] < 0) {
						check = false;
					}
				} catch (Exception a) {
					check = false;
				}

				if (check == false) {
					JOptionPane.showMessageDialog(background, "Please enter valid number");
					check = true;
				}

				else {
					number[1][1]++;
					input[1][1].setText(number[1][1] + "");
				}
			}
		});
		btnDecre[1][2].addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent e) {
				try {
					number[1][2] = Integer.parseInt(input[1][2].getText());
					if (number[1][2] < 0) {
						check = false;
					}
				} catch (Exception a) {
					check = false;
				}

				if (check == false) {
					JOptionPane.showMessageDialog(background, "Please enter valid number");
					check = true;
				}

				else {
					number[1][2]--;
					if (number[1][2] < 0)
						number[1][2] = 0;
					input[1][2].setText(number[1][2] + "");
				}
			}
		});
		btnIncre[1][2].addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent e) {
				try {
					number[1][2] = Integer.parseInt(input[1][2].getText());
					if (number[1][2] < 0) {
						check = false;
					}
				} catch (Exception a) {
					check = false;
				}

				if (check == false) {
					JOptionPane.showMessageDialog(background, "Please enter valid number");
					check = true;
				}

				else {
					number[1][2]++;
					input[1][2].setText(number[1][2] + "");
				}
			}
		});
		btnDecre[1][3].addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent e) {
				try {
					number[1][3] = Integer.parseInt(input[1][3].getText());
					if (number[1][3] < 0) {
						check = false;
					}
				} catch (Exception a) {
					check = false;
				}

				if (check == false) {
					JOptionPane.showMessageDialog(background, "Please enter valid number");
					check = true;
				}

				else {
					number[1][3]--;
					if (number[1][3] < 0)
						number[1][3] = 0;
					input[1][3].setText(number[1][3] + "");
				}
			}
		});
		btnIncre[1][3].addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent e) {
				try {
					number[1][3] = Integer.parseInt(input[1][3].getText());
					if (number[1][3] < 0) {
						check = false;
					}
				} catch (Exception a) {
					check = false;
				}

				if (check == false) {
					JOptionPane.showMessageDialog(background, "Please enter valid number");
					check = true;
				}

				else {
					number[1][3]++;
					input[1][3].setText(number[1][3] + "");
				}
			}
		});

		btnDecre[2][1].addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent e) {
				try {
					number[2][1] = Integer.parseInt(input[2][1].getText());
					if (number[2][1] < 0) {
						check = false;
					}
				} catch (Exception a) {
					check = false;
				}

				if (check == false) {
					JOptionPane.showMessageDialog(background, "Please enter valid number");
					check = true;
				}

				else {
					number[2][1]--;
					if (number[2][1] < 0)
						number[2][1] = 0;
					input[2][1].setText(number[2][1] + "");
				}
			}
		});
		btnIncre[2][1].addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent e) {
				try {
					number[2][1] = Integer.parseInt(input[2][1].getText());
					if (number[2][1] < 0) {
						check = false;
					}
				} catch (Exception a) {
					check = false;
				}

				if (check == false) {
					JOptionPane.showMessageDialog(background, "Please enter valid number");
					check = true;
				}

				else {
					number[2][1]++;
					input[2][1].setText(number[2][1] + "");
				}
			}
		});
		btnDecre[2][2].addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent e) {
				try {
					number[2][2] = Integer.parseInt(input[2][2].getText());
					if (number[2][2] < 0) {
						check = false;
					}
				} catch (Exception a) {
					check = false;
				}

				if (check == false) {
					JOptionPane.showMessageDialog(background, "Please enter valid number");
					check = true;
				}

				else {
					number[2][2]--;
					if (number[2][2] < 0)
						number[2][2] = 0;
					input[2][2].setText(number[2][2] + "");
				}
			}
		});
		btnIncre[2][2].addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent e) {
				try {
					number[2][2] = Integer.parseInt(input[2][2].getText());
					if (number[2][2] < 0) {
						check = false;
					}
				} catch (Exception a) {
					check = false;
				}

				if (check == false) {
					JOptionPane.showMessageDialog(background, "Please enter valid number");
					check = true;
				}

				else {
					number[2][2]++;
					input[2][2].setText(number[2][2] + "");
				}
			}
		});
		btnDecre[2][3].addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent e) {
				try {
					number[2][3] = Integer.parseInt(input[2][3].getText());
					if (number[2][3] < 0) {
						check = false;
					}
				} catch (Exception a) {
					check = false;
				}

				if (check == false) {
					JOptionPane.showMessageDialog(background, "Please enter valid number");
					check = true;
				}

				else {
					number[2][3]--;
					if (number[2][3] < 0)
						number[2][3] = 0;
					input[2][3].setText(number[2][3] + "");
				}
			}
		});
		btnIncre[2][3].addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent e) {
				try {
					number[2][3] = Integer.parseInt(input[2][3].getText());
					if (number[2][3] < 0) {
						check = false;
					}
				} catch (Exception a) {
					check = false;
				}

				if (check == false) {
					JOptionPane.showMessageDialog(background, "Please enter valid number");
					check = true;
				}

				else {
					number[2][3]++;
					input[2][3].setText(number[2][3] + "");
				}
			}
		});

		btnDecre[3][1].addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent e) {
				try {
					number[3][1] = Integer.parseInt(input[3][1].getText());
					if (number[3][1] < 0) {
						check = false;
					}
				} catch (Exception a) {
					check = false;
				}

				if (check == false) {
					JOptionPane.showMessageDialog(background, "Please enter valid number");
					check = true;
				}

				else {
					number[3][1]--;
					if (number[3][1] < 0)
						number[3][1] = 0;
					input[3][1].setText(number[3][1] + "");
				}
			}
		});
		btnIncre[3][1].addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent e) {
				try {
					number[3][1] = Integer.parseInt(input[3][1].getText());
					if (number[3][1] < 0) {
						check = false;
					}
				} catch (Exception a) {
					check = false;
				}

				if (check == false) {
					JOptionPane.showMessageDialog(background, "Please enter valid number");
					check = true;
				}

				else {
					number[3][1]++;
					input[3][1].setText(number[3][1] + "");
				}
			}
		});
		btnDecre[3][2].addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent e) {
				try {
					number[3][2] = Integer.parseInt(input[3][2].getText());
					if (number[3][2] < 0) {
						check = false;
					}
				} catch (Exception a) {
					check = false;
				}

				if (check == false) {
					JOptionPane.showMessageDialog(background, "Please enter valid number");
					check = true;
				}

				else {
					number[3][2]--;
					if (number[3][2] < 0)
						number[3][2] = 0;
					input[3][2].setText(number[3][2] + "");
				}
			}
		});
		btnIncre[3][2].addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent e) {
				try {
					number[3][2] = Integer.parseInt(input[3][2].getText());
					if (number[3][2] < 0) {
						check = false;
					}
				} catch (Exception a) {
					check = false;
				}

				if (check == false) {
					JOptionPane.showMessageDialog(background, "Please enter valid number");
					check = true;
				}

				else {
					number[3][2]++;
					input[3][2].setText(number[3][2] + "");
				}
			}
		});
		btnDecre[3][3].addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent e) {
				try {
					number[3][3] = Integer.parseInt(input[3][3].getText());
					if (number[3][3] < 0) {
						check = false;
					}
				} catch (Exception a) {
					check = false;
				}

				if (check == false) {
					JOptionPane.showMessageDialog(background, "Please enter valid number");
					check = true;
				}

				else {
					number[3][3]--;
					if (number[3][3] < 0)
						number[3][3] = 0;
					input[3][3].setText(number[3][3] + "");
				}
			}
		});
		btnIncre[3][3].addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent e) {
				try {
					number[3][3] = Integer.parseInt(input[3][3].getText());
					if (number[3][3] < 0) {
						check = false;
					}
				} catch (Exception a) {
					check = false;
				}

				if (check == false) {
					JOptionPane.showMessageDialog(background, "Please enter valid number");
					check = true;
				}

				else {
					number[3][3]++;
					input[3][3].setText(number[3][3] + "");
				}
			}
		});

		btnDecre[4][1].addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent e) {
				try {
					number[4][1] = Integer.parseInt(input[4][1].getText());
					if (number[4][1] < 0) {
						check = false;
					}
				} catch (Exception a) {
					check = false;
				}

				if (check == false) {
					JOptionPane.showMessageDialog(background, "Please enter valid number");
					check = true;
				}

				else {
					number[4][1]--;
					if (number[4][1] < 0)
						number[4][1] = 0;
					input[4][1].setText(number[4][1] + "");
				}
			}
		});
		btnIncre[4][1].addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent e) {
				try {
					number[4][1] = Integer.parseInt(input[4][1].getText());
					if (number[4][1] < 0) {
						check = false;
					}
				} catch (Exception a) {
					check = false;
				}

				if (check == false) {
					JOptionPane.showMessageDialog(background, "Please enter valid number");
					check = true;
				}

				else {
					number[4][1]++;
					input[4][1].setText(number[4][1] + "");
				}
			}
		});
		btnDecre[4][2].addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent e) {
				try {
					number[4][2] = Integer.parseInt(input[4][2].getText());
					if (number[4][2] < 0) {
						check = false;
					}
				} catch (Exception a) {
					check = false;
				}

				if (check == false) {
					JOptionPane.showMessageDialog(background, "Please enter valid number");
					check = true;
				}

				else {
					number[4][2]--;
					if (number[4][2] < 0)
						number[4][2] = 0;
					input[4][2].setText(number[4][2] + "");
				}
			}
		});
		btnIncre[4][2].addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent e) {
				try {
					number[4][2] = Integer.parseInt(input[4][2].getText());
					if (number[4][2] < 0) {
						check = false;
					}
				} catch (Exception a) {
					check = false;
				}

				if (check == false) {
					JOptionPane.showMessageDialog(background, "Please enter valid number");
					check = true;
				}

				else {
					number[4][2]++;
					input[4][2].setText(number[4][2] + "");
				}
			}
		});
		btnDecre[4][3].addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent e) {
				try {
					number[4][3] = Integer.parseInt(input[4][3].getText());
					if (number[4][3] < 0) {
						check = false;
					}
				} catch (Exception a) {
					check = false;
				}

				if (check == false) {
					JOptionPane.showMessageDialog(background, "Please enter valid number");
					check = true;
				}

				else {
					number[4][3]--;
					if (number[4][3] < 0)
						number[4][3] = 0;
					input[4][3].setText(number[4][3] + "");
				}
			}
		});
		btnIncre[4][3].addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent e) {
				try {
					number[4][3] = Integer.parseInt(input[4][3].getText());
					if (number[4][3] < 0) {
						check = false;
					}
				} catch (Exception a) {
					check = false;
				}

				if (check == false) {
					JOptionPane.showMessageDialog(background, "Please enter valid number");
					check = true;
				}

				else {
					number[4][3]++;
					input[4][3].setText(number[4][3] + "");
				}
			}
		});

		btnDecre[5][1].addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent e) {
				try {
					number[5][1] = Integer.parseInt(input[5][1].getText());
					if (number[5][1] < 0) {
						check = false;
					}
				} catch (Exception a) {
					check = false;
				}

				if (check == false) {
					JOptionPane.showMessageDialog(background, "Please enter valid number");
					check = true;
				}

				else {
					number[5][1]--;
					if (number[5][1] < 0)
						number[5][1] = 0;
					input[5][1].setText(number[5][1] + "");
				}
			}
		});
		btnIncre[5][1].addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent e) {
				try {
					number[5][1] = Integer.parseInt(input[5][1].getText());
					if (number[5][1] < 0) {
						check = false;
					}
				} catch (Exception a) {
					check = false;
				}

				if (check == false) {
					JOptionPane.showMessageDialog(background, "Please enter valid number");
					check = true;
				}

				else {
					number[5][1]++;
					input[5][1].setText(number[5][1] + "");
				}
			}
		});
		btnDecre[5][2].addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent e) {
				try {
					number[5][2] = Integer.parseInt(input[5][2].getText());
					if (number[5][2] < 0) {
						check = false;
					}
				} catch (Exception a) {
					check = false;
				}

				if (check == false) {
					JOptionPane.showMessageDialog(background, "Please enter valid number");
					check = true;
				}

				else {
					number[5][2]--;
					if (number[5][2] < 0)
						number[5][2] = 0;
					input[5][2].setText(number[5][2] + "");
				}
			}
		});
		btnIncre[5][2].addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent e) {
				try {
					number[5][2] = Integer.parseInt(input[5][2].getText());
					if (number[5][2] < 0) {
						check = false;
					}
				} catch (Exception a) {
					check = false;
				}

				if (check == false) {
					JOptionPane.showMessageDialog(background, "Please enter valid number");
					check = true;
				}

				else {
					number[5][2]++;
					input[5][2].setText(number[5][2] + "");
				}
			}
		});
		btnDecre[5][3].addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent e) {
				try {
					number[5][3] = Integer.parseInt(input[5][3].getText());
					if (number[5][3] < 0) {
						check = false;
					}
				} catch (Exception a) {
					check = false;
				}

				if (check == false) {
					JOptionPane.showMessageDialog(background, "Please enter valid number");
					check = true;
				}

				else {
					number[5][3]--;
					if (number[5][3] < 0)
						number[5][3] = 0;
					input[5][3].setText(number[5][3] + "");
				}
			}
		});
		btnIncre[5][3].addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent e) {
				try {
					number[5][3] = Integer.parseInt(input[5][3].getText());
					if (number[5][3] < 0) {
						check = false;
					}
				} catch (Exception a) {
					check = false;
				}

				if (check == false) {
					JOptionPane.showMessageDialog(background, "Please enter valid number");
					check = true;
				}

				else {
					number[5][3]++;
					input[5][3].setText(number[5][3] + "");
				}
			}
		});
	}

//-------------------------------------------------------METHOD-----------------------------------------------------------------------------
	public String showOpenFileDialog() {
		JFileChooser fileChooser = new JFileChooser();
		fileChooser.setCurrentDirectory(new File(System.getProperty("user.home")));
		fileChooser.addChoosableFileFilter(new FileNameExtensionFilter("MS Office Documents", "xlsx", "xls"));
		fileChooser.setAcceptAllFileFilterUsed(false);

		int result = fileChooser.showOpenDialog(this);
		if (result == JFileChooser.APPROVE_OPTION) {
			File selectedFile = fileChooser.getSelectedFile();
			return selectedFile.getAbsolutePath();
		}
		return null;
	}

	public void showSaveFileDialog() {
		JFileChooser fileChooser = new JFileChooser() {
			@Override
			public void approveSelection() {
				File selectedFile = this.getSelectedFile();
				if (selectedFile.exists() && getDialogType() == SAVE_DIALOG) {
					int result;
					if (dem == 0) {
						result = JOptionPane.showConfirmDialog(this, "The file exists, overwrite?", "Existing file",
								JOptionPane.YES_NO_CANCEL_OPTION);
					} else {
						demn++;
						result = JOptionPane.showConfirmDialog(this, "Your result was successfully saved",
								"Existing file", JOptionPane.CLOSED_OPTION);
					}
					dem++;
					switch (result) {
					case JOptionPane.CANCEL_OPTION:
						cancelSelection();
						return;
					case JOptionPane.YES_OPTION:
						btnSave.setEnabled(false);
						super.approveSelection();
						return;
					default:
						return;
					}
				}
				super.approveSelection();
			}
		};
		fileChooser.setCurrentDirectory(new File(System.getProperty("user.home")));
		fileChooser.addChoosableFileFilter(new FileNameExtensionFilter("MS Office Documents", "xlsx", "xls"));
		fileChooser.setAcceptAllFileFilterUsed(false);
		int result = fileChooser.showSaveDialog(this);
		if (result == JFileChooser.APPROVE_OPTION) {
			fileChooser.approveSelection();
			File fileToSave = fileChooser.getSelectedFile();
			excelFilePath = fileToSave.getAbsolutePath();
		}
	}

//------------Write data to Excel
	public void writeExcel(List<FunctionPoint> books, String excelFilePath) throws IOException {

		XSSFWorkbook workbook = new XSSFWorkbook();
		// Create a blank sheet
		XSSFSheet sheet;
		if (workbook.getNumberOfSheets() == 0)
			sheet = workbook.createSheet("FunctionPoint");
		else
			sheet = workbook.getSheetAt(0);

		int rowIndex = 0;
		// Write header
		writeHeader(sheet, rowIndex);
		// Write data
		rowIndex++;
		for (FunctionPoint book : books) {
			// Create row
			Row row = sheet.createRow(rowIndex);
			// Write data on row
			writeBook(book, row);
			rowIndex++;
		}

		// Write footer
		writeFooter(sheet, rowIndex);
		writeVAF(sheet);
		// Auto resize column width
		int numberOfColumn = sheet.getRow(0).getPhysicalNumberOfCells();
		autosizeColumn(sheet, numberOfColumn);
		// Create file excel
		try {
			// Write the workbook in file system
			FileOutputStream out = new FileOutputStream(new File(excelFilePath));
			workbook.write(out);
			out.close();
			if (demn == 0) {
				btnSave.setEnabled(false);
				JOptionPane.showConfirmDialog(this, "Your result was successfully saved", "Existing file",
						JOptionPane.CLOSED_OPTION);	
				
			}
		} catch (Exception e) {
			System.out.println("Cancel!");

		}
		workbook.close();
	}

	// Create dummy data
	private List<FunctionPoint> getFunctionPoint() {
		List<FunctionPoint> listBook = new ArrayList<>();
		FunctionPoint book;
		for (int i = 1; i <= 5; i++) {
			if (i == 1) {
				book = new FunctionPoint("External Inputs (EI)", resultUFC[1][1], resultUFC[1][2], resultUFC[1][3]);
				listBook.add(book);
			}
			if (i == 2) {
				book = new FunctionPoint("External Outputs (EO)", resultUFC[2][1], resultUFC[2][2], resultUFC[2][3]);
				listBook.add(book);
			}
			if (i == 3) {
				book = new FunctionPoint("External Inquiries (EQ)", resultUFC[3][1], resultUFC[3][2], resultUFC[3][3]);
				listBook.add(book);
			}
			if (i == 4) {
				book = new FunctionPoint("External Interface Files (EIF)", resultUFC[4][1], resultUFC[4][2],
						resultUFC[4][3]);
				listBook.add(book);
			}
			if (i == 5) {
				book = new FunctionPoint("Internal Logical Files (ILF)", resultUFC[5][1], resultUFC[5][2],
						resultUFC[5][3]);
				listBook.add(book);
			}
		}
		return listBook;
	}

	// Write header with format
	private void writeHeader(Sheet sheet, int rowIndex) {
		// create CellStyle
		CellStyle cellStyle = w.createStyleForHeader(sheet);

		// Create row
		Row row = sheet.createRow(rowIndex);

		// Create cells
		Cell cell = row.createCell(COLUMN_INDEX_ELEMENT);
		cell.setCellStyle(cellStyle);
		cell.setCellValue("Elements");

		cell = row.createCell(COLUMN_INDEX_L);
		cell.setCellStyle(cellStyle);
		cell.setCellValue("L");

		cell = row.createCell(COLUMN_INDEX_A);
		cell.setCellStyle(cellStyle);
		cell.setCellValue("A");

		cell = row.createCell(COLUMN_INDEX_H);
		cell.setCellStyle(cellStyle);
		cell.setCellValue("H");

		cell = row.createCell(COLUMN_INDEX_S);
		cell.setCellStyle(cellStyle);
		cell.setCellValue("Sum");
	}

	// Write data
	private static void writeBook(FunctionPoint book, Row row) {
		Cell cell = row.createCell(COLUMN_INDEX_ELEMENT);
		cell.setCellValue(book.getEle());

		cell = row.createCell(COLUMN_INDEX_L);
		cell.setCellValue(book.getL());

		cell = row.createCell(COLUMN_INDEX_A);
		cell.setCellValue(book.getA());
		cell.setCellStyle(cellStyleFormatNumber);

		cell = row.createCell(COLUMN_INDEX_H);
		cell.setCellValue(book.getH());

		// Create cell formula
		// totalMoney = price * quantity
		cell = row.createCell(COLUMN_INDEX_S, CellType.FORMULA);
		cell.setCellStyle(cellStyleFormatNumber);
		int currentRow = row.getRowNum() + 1;
		String columnL = CellReference.convertNumToColString(COLUMN_INDEX_L);
		String columnA = CellReference.convertNumToColString(COLUMN_INDEX_A);
		String columnH = CellReference.convertNumToColString(COLUMN_INDEX_H);
		cell.setCellFormula(columnA + currentRow + "+" + columnL + currentRow + "+" + columnH + currentRow);
	}

	// Write footer
	private void writeFooter(Sheet sheet, int rowIndex) {
		// Create row
		Row row = sheet.createRow(rowIndex);
		Cell cellE = row.createCell(COLUMN_INDEX_ELEMENT);
		cellE.setCellValue("UAF: ");
		Cell cell = row.createCell(COLUMN_INDEX_S, CellType.FORMULA);
		cell.setCellFormula("SUM(E2:E6)");
		int d = 1;
		Cell cellF, cellV;
		for (int i = 1; i <= 5; i++) {
			cellF = row.createCell(5 + d);
			cellV = row.createCell(5 + d + 1);
			switch (5 + d) {
			case 6:
				cellF.setCellValue("1GL: ");
				cellV.setCellValue(Integer.parseInt(textloc[1].getText()));
				d += 2;
				break;
			case 8:
				cellF.setCellValue("Pascal: ");
				cellV.setCellValue(Integer.parseInt(textloc[2].getText()));
				d += 2;
				break;
			case 10:
				cellF.setCellValue("2GL: ");
				cellV.setCellValue(Integer.parseInt(textloc[3].getText()));
				d += 2;
				break;
			case 12:
				cellF.setCellValue("C++: ");
				cellV.setCellValue(Integer.parseInt(textloc[4].getText()));
				d += 2;
				break;
			case 14:
				cellF.setCellValue("3GL: ");
				cellV.setCellValue(Integer.parseInt(textloc[5].getText()));
				d = 1;
				break;
			}

		}

		Row row1 = sheet.createRow(rowIndex + 1);
		Cell cellVa = row1.createCell(COLUMN_INDEX_ELEMENT);
		cellVa.setCellValue("VAF: ");
		Cell cell1 = row1.createCell(COLUMN_INDEX_S, CellType.FORMULA);
		cell1.setCellFormula("0.65+0.01*SUM(G2:G6,I2:I6,K2:K5)");

		for (int i = 1; i <= 5; i++) {
			cellF = row1.createCell(5 + d);
			cellV = row1.createCell(5 + d + 1);
			switch (5 + d) {
			case 6:
				cellF.setCellValue("Java 2: ");
				cellV.setCellValue(Integer.parseInt(textloc[6].getText()));
				d += 2;
				break;
			case 8:
				cellF.setCellValue("4GL: ");
				cellV.setCellValue(Integer.parseInt(textloc[7].getText()));
				d += 2;
				break;
			case 10:
				cellF.setCellValue("Excel: ");
				cellV.setCellValue(Integer.parseInt(textloc[8].getText()));
				d += 2;
				break;
			case 12:
				cellF.setCellValue("Assembler: ");
				cellV.setCellValue(Integer.parseInt(textloc[9].getText()));
				d += 2;
				break;
			case 14:
				cellF.setCellValue("SQL: ");
				cellV.setCellValue(Integer.parseInt(textloc[10].getText()));
				d = 1;
				break;
			}

		}

		Row row2 = sheet.createRow(rowIndex + 2);
		Cell cellR = row2.createCell(COLUMN_INDEX_ELEMENT);
		cellR.setCellValue("FP: ");
		Cell cell2 = row2.createCell(COLUMN_INDEX_S, CellType.FORMULA);
		cell2.setCellFormula("E7*E8");

	}

	private void writeVAF(Sheet sheet) {
		// Create row
		for (int i = 1; i <= 14; i++)
			for (int j = 1; j <= 6; j++) {
				if (i <= 5) {
					Row row = sheet.getRow(i);
					Cell cell = row.createCell(5);
					cell.setCellValue("F" + i);
					Cell cell1 = row.createCell(6);
					cell1.setCellValue(resultF[i]);
				} else if (i <= 10) {
					Row row = sheet.getRow(i - 5);
					Cell cell = row.createCell(7);
					cell.setCellValue("F" + i);
					Cell cell1 = row.createCell(8);
					cell1.setCellValue(resultF[i]);
				} else {
					Row row = sheet.getRow(i - 10);
					Cell cell = row.createCell(9);
					cell.setCellValue("F" + i);
					Cell cell1 = row.createCell(10);
					cell1.setCellValue(resultF[i]);
				}
			}

	}

	// Auto resize column width
	private static void autosizeColumn(Sheet sheet, int lastColumn) {
		for (int columnIndex = 0; columnIndex < lastColumn; columnIndex++) {
			sheet.autoSizeColumn(columnIndex);
		}
	}

	// ------------Read data from Excel
	public List<FunctionPoint> readExcel(String excelFilePath) throws IOException {
		List<FunctionPoint> listBooks = new ArrayList<>();

		// Get file
		try {
			InputStream inputStream = new FileInputStream(new File(excelFilePath));

			// Get workbook
			Workbook workbook = getWorkbook(inputStream, excelFilePath);

			// Get sheet
			Sheet sheet = workbook.getSheet("FunctionPoint");

			// Get all rows
			Iterator<Row> iterator = sheet.iterator();
			boolean oldcheck = true, stop = false;

			while (iterator.hasNext()) {
				Row nextRow = iterator.next();
				if (nextRow.getRowNum() == 0) {
					// Ignore header
					continue;
				}
				// Get all cells
				Iterator<Cell> cellIterator = nextRow.cellIterator();

				// Read cells and set value for book object
				FunctionPoint book = new FunctionPoint();
				boolean newcheck = true;
				while (cellIterator.hasNext()) {
					// Read cell

					Cell cell = cellIterator.next();
					Object cellValue = getCellValue(cell);
					if (cellValue == null || cellValue.toString().isEmpty()) {
						continue;
					}
					// Set value for book object
					int columnIndex = cell.getColumnIndex();
					newcheck = true;
					int tmp;
					switch (columnIndex) {
					case COLUMN_INDEX_L:
						try {
							tmp = new BigDecimal((double) cellValue).intValue();
							if (tmp < 0)
								newcheck = false;
						} catch (Exception e) {
							newcheck = false;
						}
						break;
					case COLUMN_INDEX_H:
						try {
							tmp = new BigDecimal((double) cellValue).intValue();
							if (tmp < 0)
								newcheck = false;
						} catch (Exception e) {
							newcheck = false;
						}
						break;
					case COLUMN_INDEX_A:
						try {
							tmp = new BigDecimal((double) cellValue).intValue();
							if (tmp < 0)
								newcheck = false;
						} catch (Exception e) {
							newcheck = false;
						}
						break;
					case 6:
						try {
							tmp = new BigDecimal((double) cellValue).intValue();
							if (tmp < 0 || tmp > 5)
								newcheck = false;
						} catch (Exception e) {
							newcheck = false;
						}
						break;
					case 8:
						try {
							tmp = new BigDecimal((double) cellValue).intValue();
							if (tmp < 0 || tmp > 5)
								newcheck = false;
						} catch (Exception e) {
							newcheck = false;
						}
						break;
					case 10:
						try {
							tmp = new BigDecimal((double) cellValue).intValue();
							if (tmp < 0 || tmp > 5)
								newcheck = false;
						} catch (Exception e) {
							newcheck = false;
						}
						break;
					default:
						break;
					}

					if (newcheck == false) {
						JOptionPane.showMessageDialog(background, "Invalid input, please check your file");
						break;
					} else {
						switch (columnIndex) {
						case COLUMN_INDEX_L:
							book.setL(new BigDecimal((double) cellValue).intValue());
							break;
						case COLUMN_INDEX_H:
							book.setH(new BigDecimal((double) cellValue).intValue());
							break;
						case COLUMN_INDEX_A:
							book.setA(new BigDecimal((double) cellValue).intValue());
							break;
						case 6:
							book.setF1(new BigDecimal((double) cellValue).intValue());
							break;
						case 8:
							book.setF2(new BigDecimal((double) cellValue).intValue());
							break;
						case 10:
							book.setF3(new BigDecimal((double) cellValue).intValue());
							break;
						default:
							break;
						}
					}
					if (cell.getColumnIndex() == 8 && cell.getRowIndex() == 5) {
						stop = true;
						break;
					}

				}
				if (newcheck == false) {
					oldcheck = false;
					break;
				}

				listBooks.add(book);
				if (stop == true)
					break;
			}

			if (oldcheck == false)
				return null;

			for (int i = 1; i <= 5; i++) {
				number[i][1] = listBooks.get(i - 1).getL();
				number[i][2] = listBooks.get(i - 1).getA();
				number[i][3] = listBooks.get(i - 1).getH();
			}
			for (int i = 1; i <= 5; i++) {
				if (i < 5) {
					resultF[i] = listBooks.get(i - 1).getF1();
					resultF[i + 5] = listBooks.get(i - 1).getF2();
					resultF[i + 10] = listBooks.get(i - 1).getF3();
					continue;
				}
				if (i == 5) {
					resultF[i] = listBooks.get(i - 1).getF1();
					resultF[i + 5] = listBooks.get(i - 1).getF2();
					continue;
				}
			}

			for (int i = 1; i <= 5; i++)
				for (int j = 1; j <= 3; j++) {
					input[i][j].setText(number[i][j] / factor[i][j] + "");
				}
			for (int i = 1; i <= 14; i++)
				selectVAF[i][resultF[i] + 1].setSelected(true);

			workbook.close();
			inputStream.close();
		} catch (Exception e) {
			System.out.println("Cancel!");
		}
		return listBooks;
	}

	// Get Workbook
	private Workbook getWorkbook(InputStream inputStream, String excelFilePath) throws IOException {
		Workbook workbook = null;
		if (excelFilePath.endsWith("xlsx")) {
			workbook = new XSSFWorkbook(inputStream);
		} else if (excelFilePath.endsWith("xls")) {
			workbook = new HSSFWorkbook(inputStream);
		} else {
			JOptionPane.showConfirmDialog(this, "The specified file is not Excel file", "Existing file",
					JOptionPane.CLOSED_OPTION);
		}

		return workbook;
	}

	// Get cell value
	private static Object getCellValue(Cell cell) {
		CellType cellType = cell.getCellType();
		Object cellValue = null;
		switch (cellType) {
		case BOOLEAN:
			cellValue = cell.getBooleanCellValue();
			break;
		case FORMULA:
			Workbook workbook = cell.getSheet().getWorkbook();
			FormulaEvaluator evaluator = workbook.getCreationHelper().createFormulaEvaluator();
			cellValue = evaluator.evaluate(cell).getNumberValue();
			break;
		case NUMERIC:
			cellValue = cell.getNumericCellValue();
			break;
		case STRING:
			cellValue = cell.getStringCellValue();
			break;
		case _NONE:
		case BLANK:
		case ERROR:
			break;
		default:
			break;
		}
		return cellValue;
	}
}
