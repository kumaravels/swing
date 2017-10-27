import java.awt.BorderLayout;
import java.awt.Dimension;
import java.awt.FlowLayout;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.awt.event.WindowAdapter;
import java.awt.event.WindowEvent;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.HashMap;
import java.util.Iterator;
import java.util.Map;
import java.util.Vector;

import javax.swing.DefaultCellEditor;
import javax.swing.JButton;
import javax.swing.JComboBox;
import javax.swing.JFileChooser;
import javax.swing.JFrame;
import javax.swing.JLabel;
import javax.swing.JOptionPane;
import javax.swing.JPanel;
import javax.swing.JProgressBar;
import javax.swing.JScrollPane;
import javax.swing.JTable;
import javax.swing.table.AbstractTableModel;
import javax.swing.table.DefaultTableModel;
import javax.swing.table.TableColumnModel;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Demo {

	private JFrame mainFrame;
	private JLabel statusLabel;
	private static String sourceFileName;
	private JTable tblMap;
	private JTable tblCategory;
	private JProgressBar pbStatus;
	// private JLabel progressStatusLabel;
	private HashMap<String, String> hashMapCategory;
	private static final int CREATED_COL_INDEX = 3;
	private static final int RESOLVED_COL_INDEX = 4;
	private static final int CLOSED_COL_INDEX = 5;
	private static final int REASSIGNMENT_COUNT = 11;
	private static final int EFFORT_HRS = 15; 
	private static final int RD_MON = 17;
	private static final int CR_MON = 18;
	private static final int MON = 19;
	private static final int DAY = 20;
	private static final int TIME = 21;
	private static final int MTTR_DURATION_DAYS = 22;
	private static final int RD_MTTR = 23;	
	private static final int PRIMARY_CATEGORY_COL_INDEX = 28;
	private static final int SECONDARY_CATEGORY_COL_INDEX = 27;
	private static final int REASON_CODE_COL_INDEX = 26;

	public Demo() throws Exception {
		prepareGUI();
	}

	public static void main(String[] args) throws Exception {
		Demo swingControlDemo = new Demo();
		swingControlDemo.showFileChooserDemo(sourceFileName);
	}

	private void prepareGUI() throws IOException {
		mainFrame = new JFrame("ASM Ticket Analysis");
		mainFrame.setSize(750, 600);
		hashMapCategory = new HashMap<String, String>();

		mainFrame.addWindowListener(new WindowAdapter() {
			public void windowClosing(WindowEvent windowEvent) {
				System.exit(0);
			}
		});

		JPanel headerPanel = new JPanel();
		JLabel headerLabel = new JLabel("", JLabel.CENTER);

		statusLabel = new JLabel("", JLabel.CENTER);
		statusLabel.setSize(50, 100);

		JPanel controlPanel = new JPanel();
		controlPanel.add(statusLabel);
		controlPanel.setLayout(new FlowLayout());

		headerPanel.add(headerLabel);
		headerPanel.add(controlPanel);

		// Generate ASM Template Excel
		createASMTemplateExcel();

		DefaultTableModel dm = new DefaultTableModel(0, 0);
		String headerColumns[] = new String[] { "Key Words", "Primary", "Secondary", "Reason" };
		dm.setColumnIdentifiers(headerColumns);
		tblCategory = new JTable();
		tblCategory.setModel(dm);
		loadCategoryData(dm);
		populateSearchKeys();
		Dimension preferredSize = new Dimension(700, 100);
		JScrollPane jscrollCategory = new JScrollPane(tblCategory);
		jscrollCategory.setPreferredSize(preferredSize);
		headerPanel.add(jscrollCategory);

		headerPanel.setLayout(new FlowLayout());
		headerPanel.setSize(100, 100);

		headerLabel.setText("Choose file:");
		final JFileChooser fileDialog = new JFileChooser();
		JButton showFileDialogButton = new JButton("Browse Ticket Dump File");

		headerPanel.setBorder(javax.swing.BorderFactory.createLineBorder(new java.awt.Color(0, 0, 0)));

		// mainFrame.add(headerPanel);
		mainFrame.add(headerPanel, BorderLayout.NORTH);

		showFileDialogButton.addActionListener(new ActionListener() {
			@SuppressWarnings("rawtypes")
			@Override
			public void actionPerformed(ActionEvent e) {
				int returnVal = fileDialog.showOpenDialog(mainFrame);

				if (returnVal == JFileChooser.APPROVE_OPTION) {
					java.io.File file = fileDialog.getSelectedFile();					
					sourceFileName = file.getAbsolutePath();
					statusLabel.setText(sourceFileName);

					// refreshScreen();
					// ////////////////////
					try {
						tblMap = new JTable(new ComboBoxTableModel(sourceFileName));
					} catch (IOException e1) {
						JOptionPane.showMessageDialog(mainFrame, e1.getMessage());
					}

					// Create the combo box editor
					@SuppressWarnings("unchecked")
					JComboBox comboBox = new JComboBox(ComboBoxTableModel.getValidStates());
					comboBox.setEditable(true);
					DefaultCellEditor editor = new DefaultCellEditor(comboBox);

					// Assign the editor to the second column
					TableColumnModel tcm = tblMap.getColumnModel();
					tcm.getColumn(1).setCellEditor(editor);

					// Set column widths
					tcm.getColumn(0).setPreferredWidth(150);
					tcm.getColumn(1).setPreferredWidth(150);

					// Set row height
					tblMap.setRowHeight(20);

					JPanel tablePanel = new JPanel();
					tablePanel.add(new JScrollPane(tblMap));
					// tablePanel.add();
					tablePanel.setSize(300, 400);

					tablePanel.setBorder(javax.swing.BorderFactory.createLineBorder(new java.awt.Color(0, 0, 0)));

					JButton btnViewTicketDumpExcel = new JButton();
					btnViewTicketDumpExcel.setText("View Ticket Dump Excel");
					btnViewTicketDumpExcel.addActionListener(new ActionListener() {
						@Override
						public void actionPerformed(ActionEvent e) {
							try {
								showPopupExcelData(sourceFileName);
							} catch (IOException e1) {
								JOptionPane.showMessageDialog(mainFrame, e1.getMessage());
							}
						}
					});

					JButton btnGenerate = new JButton();
					btnGenerate.setText("Migrate Data from Ticket Dump to ASM Template");
					btnGenerate.setSize(50, 50);

					btnGenerate.addActionListener(new ActionListener() {
						@Override
						public void actionPerformed(ActionEvent e) {
							// ////////////////////
							try {
								generateData();
							} catch (IOException e1) {
								JOptionPane.showMessageDialog(mainFrame, e1.getMessage());
							}
							// ////////////////////
						}
					});

					JButton btnViewAsmExcel = new JButton();
					btnViewAsmExcel.setText("View ASM Template");
					btnViewAsmExcel.addActionListener(new ActionListener() {
						@Override
						public void actionPerformed(ActionEvent e) {
							try {
								showPopupExcelData("ASM.xlsx");
							} catch (IOException e1) {
								JOptionPane.showMessageDialog(mainFrame, e1.getMessage());
							}
						}
					});

					pbStatus = new JProgressBar();

					JPanel bottomPanel = new JPanel();
					bottomPanel.add(btnViewTicketDumpExcel);
					bottomPanel.add(btnGenerate);
					bottomPanel.add(btnViewAsmExcel);
					bottomPanel.add(pbStatus);

					bottomPanel.setBorder(javax.swing.BorderFactory.createLineBorder(new java.awt.Color(0, 0, 0)));

					/*
					 * allPanel = new JPanel(); allPanel.add(headerPanel); allPanel.add(tablePanel);
					 * allPanel.add(bottomPanel);
					 */

					// mainFrame.add(allPanel);
					// mainFrame.add(headerPanel, BorderLayout.NORTH);
					mainFrame.add(tablePanel, BorderLayout.CENTER);
					mainFrame.add(bottomPanel, BorderLayout.SOUTH);
					mainFrame.setVisible(true);

					// ///////////////////
				} else {
					// statusLabel.setText("Open command cancelled by user." );
				}
			}
		});
		controlPanel.add(showFileDialogButton);
	}

	private void showFileChooserDemo(String fileName) throws IOException {
		mainFrame.setVisible(true);
	}

	private void loadCategoryData(DefaultTableModel dm) {
		Vector<Object> dataRow = new Vector<Object>();
		// Password
		dataRow.add("password/reset/login");

		dataRow.add("Access");
		dataRow.add("Login /Password Issues");
		dataRow.add("Password reset");
		dm.addRow(dataRow);
		// Access
		dataRow = new Vector<Object>();
		dataRow.add("access/role/account/permission");

		dataRow.add("Access");
		dataRow.add("Giving Permission to different role/application");
		dataRow.add("Access Related Issues");
		dm.addRow(dataRow);
		// Authorization
		dataRow = new Vector<Object>();
		dataRow.add("authorization");

		dataRow.add("Access");
		dataRow.add("Authorization error");
		dataRow.add("Access Related Issues");
		dm.addRow(dataRow);
		// User Creation
		dataRow = new Vector<Object>();
		dataRow.add("user creation");

		dataRow.add("Access");
		dataRow.add("User Creation/Modification");
		dataRow.add("Access Related Issues");
		dm.addRow(dataRow);
		// Db
		dataRow = new Vector<Object>();
		dataRow.add("datafix/dbscript/script/sql/query/oracle");

		dataRow.add("Database");
		dataRow.add("Static data changes");
		dataRow.add("Datafix");
		dm.addRow(dataRow);
		// Restart
		dataRow = new Vector<Object>();
		dataRow.add("restart/reboot");

		dataRow.add("Generic IT");
		dataRow.add("Restart");
		dataRow.add("Restart Server");
		dm.addRow(dataRow);
		// File Operations
		dataRow = new Vector<Object>();
		dataRow.add("file operations");

		dataRow.add("Directory / File Operations");
		dataRow.add("File operations");
		dataRow.add("File operations");
		dm.addRow(dataRow);
		// Batch issue
		dataRow = new Vector<Object>();
		dataRow.add("batch/job/batch failure/job failure/job not completed/job aborted");

		dataRow.add("Batch issue");
		dataRow.add("Batch failure (Job Status, Job Failure, Job not completed, Job Aborted)");
		dataRow.add("Batch issue");
		dm.addRow(dataRow);

	}

	private void populateSearchKeys() {
		int rowIndexCount = tblCategory.getRowCount();

		for (int rowIndex = 0; rowIndex <= rowIndexCount - 1; rowIndex++) {
			String keyString = (String) tblCategory.getValueAt(rowIndex, 0);
			String[] keys = keyString.split("/");
			String primaryString = (String) tblCategory.getValueAt(rowIndex, 1);
			String secondaryString = (String) tblCategory.getValueAt(rowIndex, 2);
			String reasonCodeString = (String) tblCategory.getValueAt(rowIndex, 3);
			for (String strKey : keys) {
				hashMapCategory.put(strKey, primaryString + ";" + secondaryString + ";" + reasonCodeString);
			}

		}

	}

	private void generateData() throws IOException {
		try {
			pbStatus.setValue(0);
			pbStatus.setVisible(true);

			FileInputStream fis = new FileInputStream(sourceFileName);
			FileInputStream fisDestination = new FileInputStream("ASM.xlsx");
			XSSFWorkbook workbook = new XSSFWorkbook(fis);
			XSSFWorkbook workbookDestination = new XSSFWorkbook(fisDestination);
			XSSFSheet spreadsheet = workbook.getSheetAt(0);
			XSSFSheet spreadsheetDest = workbookDestination.getSheetAt(0);

			for (int rowIndex = 0; rowIndex < ComboBoxTableModel.colCount; rowIndex++) {
				String destinationColIndex = tblMap.getValueAt(rowIndex, 1).toString().trim();
				if (!destinationColIndex.isEmpty()) {
					String colIndex = tblMap.getValueAt(rowIndex, 0).toString().trim();
					if (!colIndex.isEmpty()) {
						colIndex = colIndex.split(":")[0];

						int columnIndex = Integer.parseInt(colIndex);

						destinationColIndex = destinationColIndex.split(":")[0];
						int destinationColumnIndex = Integer.parseInt(destinationColIndex);
						pbStatus.setMaximum(spreadsheet.getPhysicalNumberOfRows());
						Iterator<Row> rowIterator = spreadsheet.iterator();
						XSSFRow row = (XSSFRow) rowIterator.next();
						while (rowIterator.hasNext()) {
							row = (XSSFRow) rowIterator.next();
							Iterator<Cell> cellIterator = row.cellIterator();

							while (cellIterator.hasNext()) {
								Cell sourcecell = cellIterator.next();

								if (sourcecell.getColumnIndex() == columnIndex) {

									setSourceCellDataToDestination(spreadsheetDest, destinationColumnIndex, sourcecell);
								}
								performSearchCategory(spreadsheetDest, destinationColumnIndex, sourcecell);
							}
						}

					}
				}
				pbStatus.setValue(rowIndex);
			}			
			FileOutputStream out = new FileOutputStream("ASM.xlsx");
			workbookDestination.write(out);
			out.close();
			fis.close();
			fisDestination.close();
			applyCellFormat();
			pbStatus.setMaximum(pbStatus.getValue());
			JOptionPane.showMessageDialog(mainFrame, "Data Migrated Successfully into ASM Template!!!");

		} catch (Exception ex) {
			JOptionPane.showMessageDialog(mainFrame, ex.getMessage());
		}
	}

	private void applyCellFormat() throws IOException {
		FileInputStream fisDestination = new FileInputStream("ASM.xlsx");
		
		XSSFWorkbook workbookDestination = new XSSFWorkbook(fisDestination);
		XSSFSheet spreadsheet = workbookDestination.getSheetAt(0);
		String asmDateFormat = "MM/DD/YYYY HH:MM:SS";
		String monthFormat = "MMM-YY";
		String dateFormat = "DDD";
		String hourFormat = "HH";
		int rowCount = spreadsheet.getPhysicalNumberOfRows();
		int colCount = spreadsheet.getRow(0).getPhysicalNumberOfCells();
		for (int rowIndex = 0; rowIndex < rowCount; rowIndex++) {
			for (int colIndex = 0; colIndex < colCount; colIndex++) {
				XSSFCell currentCell = spreadsheet.getRow(rowIndex).getCell(colIndex);
				if (currentCell== null ) {
					currentCell = spreadsheet.getRow(rowIndex).createCell(colIndex);
				}
				if ( rowIndex == 0){
					CellStyle style = workbookDestination.createCellStyle();
					style.setFillBackgroundColor(IndexedColors.BLUE_GREY.getIndex());
                    style.setFillPattern(CellStyle.ALT_BARS);
                    style.setBorderBottom(CellStyle.BORDER_THIN);
                    style.setBottomBorderColor(IndexedColors.BLACK.getIndex());
                    style.setBorderLeft(CellStyle.BORDER_THIN);
                    style.setLeftBorderColor(IndexedColors.BLACK.getIndex());
                    style.setBorderRight(CellStyle.BORDER_THIN);
                    style.setRightBorderColor(IndexedColors.BLACK.getIndex());
                    style.setBorderTop(CellStyle.BORDER_THIN);
                    style.setTopBorderColor(IndexedColors.BLACK.getIndex());
                    style.setFillForegroundColor(IndexedColors.BLACK.getIndex());                                        
                    style.setFillPattern(CellStyle.ALT_BARS);
                    
                    XSSFFont font= workbookDestination.createFont();
                    font.setFontHeightInPoints((short)10);
                    font.setFontName("Arial");
                    font.setColor(IndexedColors.WHITE.getIndex());
                    font.setBold(true);
                    font.setItalic(false);
                    style.setFont(font);
                    
					currentCell.setCellStyle(style);					
				}else{
						switch (colIndex) {
						case CREATED_COL_INDEX: {						
							CreationHelper creationHelper = workbookDestination.getCreationHelper();
							CellStyle cellStyle = workbookDestination.createCellStyle();
							cellStyle.setDataFormat(creationHelper.createDataFormat().getFormat(asmDateFormat));
							currentCell.setCellStyle(cellStyle);
							break;
						}
						case RESOLVED_COL_INDEX: {				
							CreationHelper creationHelper = workbookDestination.getCreationHelper();
							CellStyle cellStyle = workbookDestination.createCellStyle();
							cellStyle.setDataFormat(creationHelper.createDataFormat().getFormat(asmDateFormat));
							currentCell.setCellStyle(cellStyle);
							break;	
						}
						case CLOSED_COL_INDEX: {						
							CreationHelper creationHelper = workbookDestination.getCreationHelper();
							CellStyle cellStyle = workbookDestination.createCellStyle();
							cellStyle.setDataFormat(creationHelper.createDataFormat().getFormat(asmDateFormat));
							currentCell.setCellStyle(cellStyle);
							break;	
						}
						case REASSIGNMENT_COUNT: {						
							currentCell.setCellType(Cell.CELL_TYPE_NUMERIC);
							break;	
						}				
						case EFFORT_HRS:{						
							currentCell.setCellType(Cell.CELL_TYPE_FORMULA);
							currentCell.setCellFormula("E"+rowIndex+1+"-"+"D"+rowIndex+1);
							break;						
						}
						case RD_MON:{
							currentCell.setCellType(Cell.CELL_TYPE_FORMULA);						
							currentCell.setCellFormula("E"+rowIndex+1);
							CreationHelper creationHelper = workbookDestination.getCreationHelper();
							CellStyle cellStyle = workbookDestination.createCellStyle();
							cellStyle.setDataFormat(creationHelper.createDataFormat().getFormat(monthFormat));
							currentCell.setCellStyle(cellStyle);
							break;
						}					
						case CR_MON:{						
							currentCell.setCellType(Cell.CELL_TYPE_FORMULA);						
							currentCell.setCellFormula("D"+rowIndex+1);
							CreationHelper creationHelper = workbookDestination.getCreationHelper();
							CellStyle cellStyle = workbookDestination.createCellStyle();
							cellStyle.setDataFormat(creationHelper.createDataFormat().getFormat(monthFormat));
							currentCell.setCellStyle(cellStyle);
							break;
						}
						case MON:{
							currentCell.setCellType(Cell.CELL_TYPE_FORMULA);						
							currentCell.setCellFormula("F"+rowIndex+1);
							CreationHelper creationHelper = workbookDestination.getCreationHelper();
							CellStyle cellStyle = workbookDestination.createCellStyle();
							cellStyle.setDataFormat(creationHelper.createDataFormat().getFormat(monthFormat));
							currentCell.setCellStyle(cellStyle);						
							break;
						}
						case DAY:{
							currentCell.setCellType(Cell.CELL_TYPE_FORMULA);						
							currentCell.setCellFormula("D"+rowIndex+1);
							CreationHelper creationHelper = workbookDestination.getCreationHelper();
							CellStyle cellStyle = workbookDestination.createCellStyle();
							cellStyle.setDataFormat(creationHelper.createDataFormat().getFormat(dateFormat));
							currentCell.setCellStyle(cellStyle);
							break;
						}
						case TIME:{
							currentCell.setCellType(Cell.CELL_TYPE_FORMULA);						
							currentCell.setCellFormula("D"+rowIndex+1);
							CreationHelper creationHelper = workbookDestination.getCreationHelper();
							CellStyle cellStyle = workbookDestination.createCellStyle();
							cellStyle.setDataFormat(creationHelper.createDataFormat().getFormat(hourFormat));
							currentCell.setCellStyle(cellStyle);
							break;
						}
						case MTTR_DURATION_DAYS:{					
							currentCell.setCellType(Cell.CELL_TYPE_FORMULA);
							currentCell.setCellFormula("F"+rowIndex+1+"-D"+rowIndex+1);
							break;
						}
						case RD_MTTR:{						
							currentCell.setCellType(Cell.CELL_TYPE_FORMULA);
							currentCell.setCellFormula("E"+rowIndex+1+"-D"+rowIndex+1);
							break;
						}					
					}
				}
				
				
			}
		}
		
		FileOutputStream out = new FileOutputStream("ASM.xlsx");
		workbookDestination.write(out);
		out.close();		
		fisDestination.close();
	}

	private void setSourceCellDataToDestination(XSSFSheet spreadsheetDest, int destinationColumnIndex,
			Cell sourceCell) {
		XSSFRow rowDestination = spreadsheetDest.getRow(sourceCell.getRowIndex());
		Cell columnDestination = null;
		if (rowDestination != null) {
			columnDestination = rowDestination.getCell(destinationColumnIndex);
			if (columnDestination == null) {
				columnDestination = rowDestination.createCell(destinationColumnIndex);
			}
		} else {
			rowDestination = spreadsheetDest.createRow(sourceCell.getRowIndex());
			columnDestination = rowDestination.createCell(destinationColumnIndex);
		}
		switch (sourceCell.getCellType()) {
		case Cell.CELL_TYPE_STRING:
			columnDestination.setCellValue(sourceCell.getStringCellValue());
			break;
		case Cell.CELL_TYPE_NUMERIC:
			columnDestination.setCellValue(sourceCell.getNumericCellValue());
			break;
		}
	}

	private void performSearchCategory(XSSFSheet spreadsheetDest, int destinationColumnIndex, Cell sourceCell) {
		if (sourceCell.getCellType() == Cell.CELL_TYPE_STRING) {
			for (Map.Entry<String, String> entry : hashMapCategory.entrySet()) {
				if (sourceCell.getStringCellValue().contains(entry.getKey())) {
					XSSFRow rowDestination = spreadsheetDest.getRow(sourceCell.getRowIndex());
					Cell columnDestination = null;
					if (rowDestination != null) {
						columnDestination = rowDestination.getCell(REASON_CODE_COL_INDEX);
						if (columnDestination == null) {
							columnDestination = rowDestination.createCell(REASON_CODE_COL_INDEX);
						}
					} else {
						rowDestination = spreadsheetDest.createRow(sourceCell.getRowIndex());
						columnDestination = rowDestination.createCell(REASON_CODE_COL_INDEX);
					}
					String[] categoryValues = entry.getValue().split(";");
					columnDestination.setCellValue(categoryValues[2]);
					columnDestination = rowDestination.getCell(SECONDARY_CATEGORY_COL_INDEX);
					if (columnDestination == null) {
						columnDestination = rowDestination.createCell(SECONDARY_CATEGORY_COL_INDEX);
					}
					columnDestination.setCellValue(categoryValues[1]);
					columnDestination = rowDestination.getCell(PRIMARY_CATEGORY_COL_INDEX);
					if (columnDestination == null) {
						columnDestination = rowDestination.createCell(PRIMARY_CATEGORY_COL_INDEX);
					}
					columnDestination.setCellValue(categoryValues[0]);
				}
			}
		}
	}
	
	private void showPopupExcelData(String excelFileName) throws IOException {
		final JFrame popupFrame = new JFrame("View Excel File : " + excelFileName);
		popupFrame.setSize(750, 600);

		popupFrame.addWindowListener(new WindowAdapter() {
			public void windowClosing(WindowEvent windowEvent) {
				// popupFrame.setVisible(false);
				popupFrame.dispose();
			}
		});

		DefaultTableModel dmPopup = new DefaultTableModel(0, 0);
		int excelColumnCount = 0;

		FileInputStream fis = new FileInputStream(excelFileName);

		XSSFWorkbook workbook = new XSSFWorkbook(fis);
		XSSFSheet spreadsheet = workbook.getSheetAt(0);

		excelColumnCount = spreadsheet.getRow(0).getPhysicalNumberOfCells();

		String headerColumns[] = new String[excelColumnCount];

		for (int colIndex = 0; colIndex <= excelColumnCount - 1; colIndex++) {
			XSSFCell currentCell = spreadsheet.getRow(0).getCell(colIndex);
			headerColumns[colIndex] = currentCell.toString();
		}

		dmPopup.setColumnIdentifiers(headerColumns);
		JTable tblPopup = new JTable();
		tblPopup.setModel(dmPopup);

		for (int colIndex = 0; colIndex <= excelColumnCount - 1; colIndex++) {
			tblPopup.getColumnModel().getColumn(colIndex).setWidth(100);
		}

		Dimension preferredSize = new Dimension(700, 600);
		JScrollPane jscrollCategory = new JScrollPane(tblPopup);
		jscrollCategory.setPreferredSize(preferredSize);
		JPanel panel = new JPanel();
		panel.add(jscrollCategory);

		panel.setLayout(new FlowLayout());

		Iterator<Row> rowIterator = spreadsheet.iterator();
		XSSFRow row = (XSSFRow) rowIterator.next();
		Vector<Object> dataRow = null;
		while (rowIterator.hasNext()) {
			row = (XSSFRow) rowIterator.next();
			Iterator<Cell> cellIterator = row.cellIterator();
			dataRow = new Vector<Object>();
			while (cellIterator.hasNext()) {
				Cell sourcecell = cellIterator.next();
				switch (sourcecell.getCellType()) {
				case Cell.CELL_TYPE_STRING:

					dataRow.add(sourcecell.getStringCellValue());

					break;
				case Cell.CELL_TYPE_NUMERIC:
					dataRow.add(sourcecell.getNumericCellValue());
					break;
				}
			}
			dmPopup.addRow(dataRow);
		}
		popupFrame.add(panel);
		popupFrame.setVisible(true);
	}

	private void createASMTemplateExcel() throws IOException {
		File file = new File("ASM.xlsx");
		if (file.exists()) {
			file.delete();
		}

		XSSFWorkbook workbook = new XSSFWorkbook();
		XSSFSheet spreadsheet = workbook.createSheet("formula");
		XSSFRow row = spreadsheet.createRow(0);
		XSSFCell cell = row.createCell(0);
		cell.setCellValue("Incident");
		cell = row.createCell(1);
		cell.setCellValue("Type");
		cell = row.createCell(2);
		cell.setCellValue("Priority");
		cell = row.createCell(3);
		cell.setCellValue("Created");
		// cell.setCellType(Cell.CELL_TYPE_NUMERIC);
		cell = row.createCell(4);
		cell.setCellValue("Resolved");
		// cell.setCellType(Cell.CELL_TYPE_NUMERIC);
		cell = row.createCell(5);
		cell.setCellValue("Closed");
		// cell.setCellType(Cell.CELL_TYPE_NUMERIC);
		cell = row.createCell(6);
		cell.setCellValue("Status");
		cell = row.createCell(7);
		cell.setCellValue("Assigned To");
		cell = row.createCell(8);
		cell.setCellValue("Assignment Group");
		cell = row.createCell(9);
		cell.setCellValue("Tower");
		cell = row.createCell(10);
		cell.setCellValue("Severity");
		cell = row.createCell(11);
		cell.setCellValue("Reassignment count");
		// cell.setCellType(Cell.CELL_TYPE_NUMERIC);
		cell = row.createCell(12);
		cell.setCellValue("Short Description");
		cell = row.createCell(13);
		cell.setCellValue("Description");
		cell = row.createCell(14);
		cell.setCellValue("Causing CI");
		cell = row.createCell(15);
		cell.setCellValue("Effort (Hrs)");
		cell = row.createCell(16);
		cell.setCellValue("KeDB referred");
		cell = row.createCell(17);
		cell.setCellValue("Rd_Mon");
		cell = row.createCell(18);
		cell.setCellValue("CR_Mon");
		cell = row.createCell(19);
		cell.setCellValue("MON");
		cell = row.createCell(20);
		cell.setCellValue("DAY");
		cell = row.createCell(21);
		cell.setCellValue("TIME");
		cell = row.createCell(22);
		cell.setCellValue("MTTR (Duration - Days)");
		cell = row.createCell(23);
		cell.setCellValue("Rd_MTTR");
		cell = row.createCell(24);
		cell.setCellValue("Product Type");
		cell = row.createCell(25);
		cell.setCellValue("Technology");
		cell = row.createCell(26);
		cell.setCellValue("Reason Code");
		cell = row.createCell(27);
		cell.setCellValue("Secondary Category");
		cell = row.createCell(28);
		cell.setCellValue("Primary Category");
		cell = row.createCell(29);
		cell.setCellValue("3R Analysis");
		cell = row.createCell(30);
		cell.setCellValue("L1.5 Scope");		

		FileOutputStream out = new FileOutputStream(new File("ASM.xlsx"));
		workbook.write(out);
		out.close();
		System.out.println("ASM.xlsx Created successfully");
	}
}

@SuppressWarnings("serial")
class ComboBoxTableModel extends AbstractTableModel {

	protected static int colCount;
	protected Object[][] data;
	protected static final String[] validStates = { " ", "0:Incident", "1:Type", "2:Priority", "3:Created",
			"4:Resolved", "5:Closed", "6:Status", "7:Assigned To", "8:Assignment Group", "9:Tower", "10:Severity",
			"11:Reassignment count", "12:Short Description", "13:Description", "14:Causing CI", "15:Effort (Hrs)",
			"16:KeDB referred", "17:Rd_Mon", "18:CR_Mon", "19:MON", "20:DAY", "21:TIME", "22:MTTR (Duration - Days)",
			"23:Rd_MTTR", "24:Product Type", "25:Technology", "26:Reason Code", "27:Secondary Category",
			"28:Primary Category", "29:3R Analysis", "30:L1.5 Scope" };

	public ComboBoxTableModel(String fileName) throws IOException {
		initilizeData(fileName);
	}

	public void initilizeData(String fileName) throws IOException {
		FileInputStream fis = new FileInputStream(new File(fileName));
		XSSFWorkbook workbook = new XSSFWorkbook(fis);
		XSSFSheet spreadsheet = workbook.getSheetAt(0);
		Iterator<Row> rowIterator = spreadsheet.iterator();
		XSSFRow row = null;
		colCount = spreadsheet.getRow(0).getPhysicalNumberOfCells();
		data = new Object[colCount][validStates.length];
		row = (XSSFRow) rowIterator.next();
		Iterator<Cell> cellIterator = row.cellIterator();
		int rowIndex = 0;
		while (cellIterator.hasNext()) {

			Cell cell = cellIterator.next();
			data[rowIndex][0] = cell.getColumnIndex() + ":" + cell.getStringCellValue();
			 
			data[rowIndex][1] = getMatchedField(cell.getStringCellValue());//validStates[0];
			rowIndex++;
		}
	}
	
	private String getMatchedField(String sourceFieldValue){
		String matchedField = validStates[0];
		for(String strASMField : validStates ){
			if (strASMField.endsWith(sourceFieldValue)){
				matchedField = strASMField;
			}
		}
		return matchedField;
	}

	// Implementation of TableModel interface
	public int getRowCount() {
		return data.length;
	}

	public int getColumnCount() {
		return COLUMN_COUNT;
	}

	public Object getValueAt(int row, int column) {
		return data[row][column];
	}

	@SuppressWarnings({ "unchecked", "rawtypes" })
	public Class getColumnClass(int column) {
		return (data[0][column]).getClass();
	}

	public String getColumnName(int column) {
		return columnNames[column];
	}

	public boolean isCellEditable(int row, int column) {
		return column == 1;
	}

	public void setValueAt(Object value, int row, int column) {
		if (isValidValue(value)) {
			data[row][column] = value;
			fireTableRowsUpdated(row, row);
		}
	}

	// Extra public methods
	public static String[] getValidStates() {
		return validStates;
	}

	// Protected methods
	protected boolean isValidValue(Object value) {
		if (value instanceof String) {
			String sValue = (String) value;

			for (int i = 0; i < validStates.length; i++) {
				if (sValue.equals(validStates[i])) {
					return true;
				}
			}
		}

		return false;
	}

	protected static final int COLUMN_COUNT = 2;

	protected static final String[] columnNames = { "Source Dump", "Existing ASM Template" };

}

