import java.awt.BorderLayout;
import java.awt.Component;
import java.awt.event.WindowAdapter;
import java.awt.event.WindowEvent;
import java.io.File;
import java.util.Arrays;
import java.util.Enumeration;
import java.util.HashMap;
import java.util.Vector;
import java.awt.FileDialog;

import javax.swing.Box;
import javax.swing.BoxLayout;
import javax.swing.JButton;
import javax.swing.JFrame;
import javax.swing.JScrollPane;
import javax.swing.JTable;
import javax.swing.JViewport;
import javax.swing.table.DefaultTableModel;
import javax.swing.table.TableColumn;
import javax.swing.table.TableModel;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Workbook;

public class MainGUI_2 {
	static JFrame mainFrame;
	static JScrollPane scrollPane;
	static DBClass myDB;
	static FileDialog fdlg;
	static String curTableName = "";
	
	public static void main(String[] args) {
		mainFrame = new JFrame("SuperDB_Viewer");
		mainFrame.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
		scrollPane = new JScrollPane();
		String curPath = System.getProperty("user.dir");
		myDB = new DBClass(curPath);
		// добавляем массив названий для кнопок
		String buttonNames[] = { "Show students", "Show lecturers", "Show subjects", "Show universities",
				"Show exam_marks", "Show subj_lect" };
		// и массив названий таблиц в БД
		String tableNames[] = { "student", "lecturer", "subject", "university", "exam_marks", "subj_lect" };
		String tableNamesRus[] = { "Инфо о студентах", "Инфо о преподавателях", "Инфо о предметах", 
									"Инфо об университетах", "Инфо об экзаменах", "Инфо кто что ведет" };
		
		// создаем хеш-мап для сопоставления названия кнопки и таблицы
		HashMap<String, String> mapForTables = new HashMap<>();
		HashMap<String, String> mapForTablesRus = new HashMap<>();
		for (int i = 0; i < buttonNames.length; i++) { // в цикле сопоставляем кнопки и таблицы
			mapForTables.put(buttonNames[i], tableNames[i]);
			mapForTablesRus.put(buttonNames[i], tableNamesRus[i]);
		}
		mainFrame.add(setMenu(buttonNames, mapForTables, mapForTablesRus), BorderLayout.NORTH);
		mainFrame.add(setBottom(), BorderLayout.SOUTH);
		mainFrame.setSize(600, 400);
		mainFrame.setMinimumSize(mainFrame.getSize());
		mainFrame.setVisible(true);
		mainFrame.pack();
		mainFrame.addWindowListener(new WindowAdapter() { // слушатель закрытия окна, чтобы отключиться от БД
			@Override
			public void windowClosing(WindowEvent e) {
				System.out.println("Exit");
				myDB.closeConnection();
				e.getWindow().dispose();
			}
		});
		fdlg = new FileDialog(mainFrame, "");
		fdlg.setMode(FileDialog.SAVE); //делаем созданный диалог диалогом сохранения
	}

	private static Component setBottom() { //метод для создания нижних кнопок 
		Box bottom = new Box(BoxLayout.X_AXIS);
		JButton toWord = new JButton("toWord");
		toWord.addActionListener(e -> {
			String[] columnNames = getTableHeader();
			String[][] data = getTableData();
			File file = getFile("Save to Word file", "docx");
			if (!file.getName().contains("nullnull")) ToOffice.toWordDocx(columnNames, data, file, curTableName);
		});
		toWord.setEnabled(false); //делаем кнопку неактивной
		bottom.add(toWord);
		bottom.add(Box.createHorizontalGlue());
		JButton toExcel = new JButton("toExcel");
		toExcel.addActionListener(e->{
			String[] columnNames = getTableHeader();
			String[][] data = getTableData();
			File file = getFile("Save to Excel file", "xls");
			if (!file.getName().contains("nullnull")) ToOffice.toExcel(columnNames, data, file, curTableName);
		});
		toExcel.setEnabled(false); //делаем кнопку неактивной
		bottom.add(toExcel);
		return bottom;
	}
	
	private static File getFile(String caption, String ext) {//метод получения файла для сохранения
		System.out.println(caption);
		fdlg.setTitle(caption); //задаем ему заголовок
		fdlg.setFile("*."+ext);
		fdlg.setVisible(true); 
		String fileName = fdlg.getDirectory()+fdlg.getFile();
		if (!fileName.contains("."+ext)) fileName = fileName.concat("."+ext);
		File file = new File(fileName);
		System.out.println("file = "+file);
		return file;
	}
	
	private static String[] getTableHeader() { //для получения названия столбцов таблицы
		JViewport viewPort = (JViewport) scrollPane.getComponent(0);
		JTable tempTable = (JTable) viewPort.getComponent(0);
		TableModel tableModel = tempTable.getModel();
		int colCount = tableModel.getColumnCount();
		String[] columnNames = new String[colCount];
		for (int i = 0; i < colCount; i++) {
			columnNames[i] = tableModel.getColumnName(i);
		}
		System.out.println(" " + Arrays.asList(columnNames));
		return columnNames;
	}

	private static String[][] getTableData() { //для получения содержимого таблицы
		JViewport viewPort = (JViewport) scrollPane.getComponent(0);
		JTable tempTable = (JTable) viewPort.getComponent(0);
		TableModel tableModel = tempTable.getModel();
		int colCount = tableModel.getColumnCount();
		int rowCount = tableModel.getRowCount();
		String[][] data = new String[rowCount][colCount];
		for (int i = 0; i < rowCount; i++) {
			for (int j = 0; j < colCount; j++) {
				data[i][j] = (String) tableModel.getValueAt(i, j);
			}
		}
		for (int i = 0; i < data.length; i++) {
			System.out.println(Arrays.asList(data[i]));
		}
		return data;
	}

//метод setMenu принимает в качестве параметров массив названий кнопок и хеш-мап
	private static Component setMenu(String[] buttonNames, HashMap<String, String> mapForTables, HashMap<String, String> mapForTablesRus) {
		Box mainMenu = new Box(BoxLayout.X_AXIS);
		for (int i = 0; i < buttonNames.length; i++) {// цикл по всем названиям кнопок
			JButton tempButton = new JButton(buttonNames[i]);// создаем временную кнопку
			tempButton.addActionListener(e -> { // слушатель нажатия кнопки
				// создаем таблицу, получая ее из БД вызовом метода getTableWithJoin,
				// которому передаем имя таблицы, связанной с названием нажатой кнопки
				JTable tempTable = myDB.getTableWithJoin(mapForTables.get(e.getActionCommand()));
				curTableName = mapForTablesRus.get(e.getActionCommand());
				mainFrame.remove(scrollPane); // убираем предыдущую панель с таблицей
				// даже если не было еще панели с таблицей
				scrollPane = new JScrollPane(tempTable);// создаем новую панель с прокруткой
				// нужно для корректного отображения больших таблиц
				tempTable.setFillsViewportHeight(true);
				// вставляем панель с таблицей в центр окна
				mainFrame.add(scrollPane, BorderLayout.CENTER);
				mainFrame.pack(); // подстраиваем размеры окна под размеры таблицы
				//для активации нижних кнопок
				BorderLayout layout = (BorderLayout) mainFrame.getContentPane().getLayout();
				Box southBox = (Box) layout.getLayoutComponent(BorderLayout.SOUTH);
				for (int j = 0; j < southBox.getComponentCount(); j++) {
					southBox.getComponent(j).setEnabled(true);
				} 				
			});
			mainMenu.add(tempButton);
		}
		return mainMenu;
	}
}