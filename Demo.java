import java.awt.BorderLayout;
import java.awt.Dimension;
import java.awt.FlowLayout;
import java.awt.GridLayout;
import java.awt.ScrollPane;
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
import javax.swing.JPanel;
import javax.swing.JScrollPane;
import javax.swing.JTable;
import javax.swing.table.AbstractTableModel;
import javax.swing.table.DefaultTableModel;
import javax.swing.table.TableColumnModel;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Demo {
	private JFrame mainFrame;
	private JLabel headerLabel;
	private JLabel statusLabel;
	private JPanel headerPanel;
	private JPanel controlPanel;
	private static String sourceFileName;
	private JTable tblMap;
	private JTable tblCategory;
	private JPanel tablePanel;
	private JButton generateButton;
	private JPanel bottomPanel;	
	private HashMap<String, String> hashMapCategory;
	private static final int PRIMARY_CATEGORY_COL_INDEX=28;
	private static final int SECONDARY_CATEGORY_COL_INDEX=27;
	private static final int REASON_CODE_COL_INDEX=26;
	
	public Demo() throws Exception {
		prepareGUI();
	}

	public static void main(String[] args) throws Exception {
		Demo swingControlDemo = new Demo();
		swingControlDemo.showFileChooserDemo(sourceFileName);
	}

	private void prepareGUI() throws Exception {
		mainFrame = new JFrame("ASM Ticket Analysis");
		mainFrame.setSize(750, 600);
		hashMapCategory = new HashMap<String, String>();

		mainFrame.addWindowListener(new WindowAdapter() {
			public void windowClosing(WindowEvent windowEvent) {
				System.exit(0);
			}
		});

		headerPanel = new JPanel();
		headerLabel = new JLabel("", JLabel.CENTER);

		statusLabel = new JLabel("", JLabel.CENTER);
		statusLabel.setSize(50, 100);

		controlPanel = new JPanel();
		controlPanel.add(statusLabel);
		controlPanel.setLayout(new FlowLayout());

		headerPanel.add(headerLabel);
		headerPanel.add(controlPanel);
		
		DefaultTableModel dm = new DefaultTableModel(0, 0);
	    String headerColumns[] = new String[] { "Key Words", "Primary", "Secondary", "Reason" };
	    dm.setColumnIdentifiers(headerColumns);	    
	    tblCategory = new JTable();
	    tblCategory.setModel(dm);	    
	    loadCategoryData(dm);	
	    Dimension preferredSize = new Dimension(700,100);	    
	    JScrollPane jscrollCategory = new JScrollPane(tblCategory);
	    jscrollCategory.setPreferredSize(preferredSize);	   
	    headerPanel.add(jscrollCategory); 

		headerPanel.setLayout(new FlowLayout());
		headerPanel.setSize(100, 100);

		headerLabel.setText("Choose file:");
		final JFileChooser fileDialog = new JFileChooser();
		JButton showFileDialogButton = new JButton("Browse Ticket Dump File");
		
		headerPanel.setBorder(javax.swing.BorderFactory.createLineBorder(new java.awt.Color(0, 0, 0)));

		//mainFrame.add(headerPanel);
		mainFrame.add(headerPanel, BorderLayout.NORTH);

		showFileDialogButton.addActionListener(new ActionListener() {
			@Override
			public void actionPerformed(ActionEvent e) {
				int returnVal = fileDialog.showOpenDialog(mainFrame);

				if (returnVal == JFileChooser.APPROVE_OPTION) {
					java.io.File file = fileDialog.getSelectedFile();
					// statusLabel.setText("File Selected :" + file.getName());
					sourceFileName = file.getAbsolutePath();
					statusLabel.setText(sourceFileName);
					System.out.println(sourceFileName);

					// refreshScreen();
					// ////////////////////
					try {
						tblMap = new JTable(new ComboBoxTableModel(
								sourceFileName));
					} catch (IOException e1) {
						// TODO Auto-generated catch block
						e1.printStackTrace();
					}

					// Create the combo box editor
					@SuppressWarnings("unchecked")
					JComboBox comboBox = new JComboBox(ComboBoxTableModel
							.getValidStates());
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
					
					tablePanel = new JPanel();
					tablePanel.add(new JScrollPane(tblMap));
					//tablePanel.add();
					tablePanel.setSize(300, 400);
					
					tablePanel.setBorder(javax.swing.BorderFactory.createLineBorder(new java.awt.Color(0, 0, 0)));

					generateButton = new JButton();
					generateButton
							.setText("Migrate Data from Ticket Dump to ASM Template");
					generateButton.setSize(50, 50);

					generateButton.addActionListener(new ActionListener() {
						@Override
						public void actionPerformed(ActionEvent e) {
							// ////////////////////
							generateData();
							// ////////////////////
						}
					});

					bottomPanel = new JPanel();
					bottomPanel.add(generateButton);
					
					bottomPanel.setBorder(javax.swing.BorderFactory.createLineBorder(new java.awt.Color(0, 0, 0)));

					/*allPanel = new JPanel();
					allPanel.add(headerPanel);
					allPanel.add(tablePanel);
					allPanel.add(bottomPanel);*/

					//mainFrame.add(allPanel);
					//mainFrame.add(headerPanel, BorderLayout.NORTH);
					mainFrame.add(tablePanel, BorderLayout.CENTER);
					mainFrame.add(bottomPanel,BorderLayout.SOUTH);
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
	
	private void loadCategoryData(DefaultTableModel dm){
		Vector<Object> dataRow = new Vector<Object>();
        //Password
        dataRow.add("password/reset/login");        
        hashMapCategory.put("password", "Access;Login /Password Issues;Password reset");
        hashMapCategory.put("reset", "Access;Login /Password Issues;Password reset");
        hashMapCategory.put("login", "Access;Login /Password Issues;Password reset");
        dataRow.add("Access");
        dataRow.add("Login /Password Issues");
        dataRow.add("Password reset");
        dm.addRow(dataRow);
	    //Access
        dataRow = new Vector<Object>();				        
        dataRow.add("access/role/account/permission");
        hashMapCategory.put("access", "Access;Giving Permission to different role/application;Access Related Issues");
        hashMapCategory.put("role", "Access;Giving Permission to different role/application;Access Related Issues");
        hashMapCategory.put("account", "Access;Giving Permission to different role/application;Access Related Issues");
        hashMapCategory.put("permission", "Access;Giving Permission to different role/application;Access Related Issues");
        hashMapCategory.put("Authorization", "Access;Authorization error;Access Related Issues");
        hashMapCategory.put("User Creation", "Access;User Creation/Modification;Access Related Issues");
        dataRow.add("Access");
        dataRow.add("Giving Permission to different role/application");
        dataRow.add("Access Related Issues");        
        dm.addRow(dataRow);
        //Db
        dataRow = new Vector<Object>();				        
        dataRow.add("datafix/dbscript/script/sql/query/oracle");
        hashMapCategory.put("datafix", "Database;Execution of DML or SPs;Datafix");
        hashMapCategory.put("dbscript", "Database;Execution of DML or SPs;Datafix");
        hashMapCategory.put("script", "Database;Execution of DML or SPs;Datafix");
        hashMapCategory.put("sql", "Database;Execution of DML or SPs;Datafix");
        hashMapCategory.put("query", "Database;Execution of DML or SPs;Datafix");
        hashMapCategory.put("oracle", "Database;Execution of DML or SPs;Datafix");        
        dataRow.add("Database");
        dataRow.add("Static data changes");
        dataRow.add("Datafix");
        dm.addRow(dataRow);        
        //Restart
        dataRow = new Vector<Object>();				        
        dataRow.add("restart/reboot");
        hashMapCategory.put("restart", "Generic IT;Restart;Restart Server");
        hashMapCategory.put("Restart", "Generic IT;Restart;Restart Server");
        hashMapCategory.put("restarted", "Generic IT;Restart;Restart Server");
        hashMapCategory.put("reboot", "Generic IT;Restart;Restart Server");
        hashMapCategory.put("Reboot", "Generic IT;Restart;Restart Server");
        hashMapCategory.put("rebooted", "Generic IT;Restart;Restart Server");
        dataRow.add("Generic IT");
        dataRow.add("Restart");
        dataRow.add("Restart Server");
        dm.addRow(dataRow);
        //File Operations
        dataRow = new Vector<Object>();
        dataRow.add("File Operations");
        hashMapCategory.put("File Operations", "Directory / File Operations;File operations;File operations");
        hashMapCategory.put("File Operations", "Directory / File Operations;File operations;File operations");
        hashMapCategory.put("File Operations", "Directory / File Operations;File operations;File operations");        
        dataRow.add("Directory / File Operations");
        dataRow.add("File operations");
        dataRow.add("File operations");
        dm.addRow(dataRow);
        //Batch issue
        dataRow = new Vector<Object>();				        
        dataRow.add("Batch issue");
        hashMapCategory.put("Batch issue", "Batch issue;Batch failure (Job Status, Job Failure, Job not completed, Job Aborted);Batch issue");
        hashMapCategory.put("batch", "Batch issue;Batch failure (Job Status, Job Failure, Job not completed, Job Aborted);Batch issue");
        hashMapCategory.put("Batch Failure", "Batch issue;Batch failure (Job Status, Job Failure, Job not completed, Job Aborted);Batch issue");
        hashMapCategory.put("Job Failure", "Batch issue;Batch failure (Job Status, Job Failure, Job not completed, Job Aborted);Batch issue");
        hashMapCategory.put("Job not completed", "Batch issue;Batch failure (Job Status, Job Failure, Job not completed, Job Aborted);Batch issue");
        hashMapCategory.put("Job Aborted", "Batch issue;Batch failure (Job Status, Job Failure, Job not completed, Job Aborted);Batch issue");
        dataRow.add("Batch issue");
        dataRow.add("Batch failure (Job Status, Job Failure, Job not completed, Job Aborted)");
        dataRow.add("Batch issue");
        dm.addRow(dataRow);
        
	}

	private void generateData() {
		try {

			FileInputStream fis = new FileInputStream(sourceFileName);
			FileInputStream fisDestination = new FileInputStream("ASM.xlsx");
			XSSFWorkbook workbook = new XSSFWorkbook(fis);
			XSSFWorkbook workbookDestination = new XSSFWorkbook(fisDestination);
			XSSFSheet spreadsheet = workbook.getSheetAt(0);
			XSSFSheet spreadsheetDest = workbookDestination.getSheetAt(0);
			

			for (int rowIndex = 0; rowIndex < ComboBoxTableModel.colCount; rowIndex++) {
				String destinationColIndex = tblMap.getValueAt(rowIndex, 1).toString().trim();
				if (!destinationColIndex.isEmpty()) {
					String colIndex = tblMap.getValueAt(rowIndex, 0).toString()
							.trim();
					if (!colIndex.isEmpty()) {
						colIndex = colIndex.split(":")[0];

						int columnIndex = Integer.parseInt(colIndex);
						
						destinationColIndex = destinationColIndex.split(":")[0];
						int destinationColumnIndex = Integer.parseInt(destinationColIndex); 
						Iterator<Row> rowIterator = spreadsheet.iterator();
						XSSFRow row = (XSSFRow) rowIterator.next();
						while (rowIterator.hasNext()) {
							row = (XSSFRow) rowIterator.next();
							Iterator<Cell> cellIterator = row.cellIterator();
							
							while(cellIterator.hasNext()){
								Cell sourcecell = cellIterator.next();
								
								if (sourcecell.getColumnIndex() == columnIndex) {
																		
									setSourceCellDataToDestination(
											spreadsheetDest,
											destinationColumnIndex, sourcecell);																		
								}
								performSearchCategory(spreadsheetDest,destinationColumnIndex,sourcecell);
								
							}
						}

					}
				}
			}

			FileOutputStream out = new FileOutputStream("ASM.xlsx");
			workbookDestination.write(out);
			out.close();
			System.out.println("File Updated");
			fis.close();
			fisDestination.close();

		} catch (Exception ex) {
			System.out.println(ex);
		}
	}

	private void setSourceCellDataToDestination(XSSFSheet spreadsheetDest,
			int destinationColumnIndex, Cell sourceCell) {
		XSSFRow rowDestination = spreadsheetDest.getRow(sourceCell.getRowIndex());
		Cell columnDestination = null;
		if (rowDestination != null) {
			columnDestination = rowDestination.getCell(destinationColumnIndex);
			if (columnDestination == null) {
				columnDestination = rowDestination.createCell(destinationColumnIndex);
			}
		}else {
			rowDestination = spreadsheetDest.createRow(sourceCell.getRowIndex());
			columnDestination = rowDestination.createCell(destinationColumnIndex);
		}
		switch (sourceCell.getCellType())
		{
			case Cell.CELL_TYPE_STRING:
				
				columnDestination.setCellValue(sourceCell
						.getStringCellValue());											
			break;
			case Cell.CELL_TYPE_NUMERIC:											
				columnDestination.setCellValue(sourceCell
						.getNumericCellValue());		
			break;
		}
	}
	
	private void performSearchCategory(XSSFSheet spreadsheetDest,
			int destinationColumnIndex, Cell sourceCell){
		if (sourceCell.getCellType() == Cell.CELL_TYPE_STRING){
			for(Map.Entry<String, String> entry : hashMapCategory.entrySet()){
				if (sourceCell.getStringCellValue().contains(entry.getKey())){
					XSSFRow rowDestination = spreadsheetDest.getRow(sourceCell.getRowIndex());
					Cell columnDestination = null;
					if (rowDestination != null) {
						columnDestination = rowDestination.getCell(REASON_CODE_COL_INDEX);
						if (columnDestination == null) {
							columnDestination = rowDestination.createCell(destinationColumnIndex);
						}
					}else {
						rowDestination = spreadsheetDest.createRow(sourceCell.getRowIndex());
						columnDestination = rowDestination.createCell(REASON_CODE_COL_INDEX);
					}
					String[] categoryValues = entry.getValue().split(";");
					columnDestination.setCellValue(categoryValues[2]);	
					columnDestination = rowDestination.getCell(SECONDARY_CATEGORY_COL_INDEX);
					if (columnDestination == null){
						columnDestination = rowDestination.createCell(SECONDARY_CATEGORY_COL_INDEX);
					}
					columnDestination.setCellValue(categoryValues[1]);
					columnDestination = rowDestination.getCell(PRIMARY_CATEGORY_COL_INDEX);
					if (columnDestination == null){
						columnDestination = rowDestination.createCell(PRIMARY_CATEGORY_COL_INDEX);
					}
					columnDestination.setCellValue(categoryValues[0]);
				}
			}
		}
	}

}

@SuppressWarnings("serial")
class ComboBoxTableModel extends AbstractTableModel {

	protected static int colCount;
	protected Object[][] data;
	protected static final String[] validStates = {			
			" ", "0:Incident", "1:Type", "2:Priority", "3:Created",
			"4:Resolved", "5:Closed", "6:Status", "7:Assigned To",
			"8:Assignment Group", "9:Tower", "10:Severity",
			"11:Reassignment count", "12:Short Description", "13:Description",
			"14:Causing CI", "15:Effort (Hrs)", "16:KeDB referred",
			"17:Rd_Mon", "18:CR_Mon", "19:MON", "20:DAY", "21:TIME",
			"22:MTTR (Duration - Days)", "23:Rd_MTTR", "24:Product Type",
			"25:Technology", "26:Reason Code", "27:Secondary Category",
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
			data[rowIndex][0] = cell.getColumnIndex() + ":"
					+ cell.getStringCellValue();
			data[rowIndex][1] = validStates[0];
			rowIndex++;
		}
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

	@SuppressWarnings("unchecked")
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

	protected static final String[] columnNames = { "Source Dump",
			"Existing ASM Template" };

}
