import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.hssf.usermodel.HSSFCellStyle;
//import org.apache.poi.hssf.usermodel.HSSFErrorConstants;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hssf.util.HSSFColor;
//import org.apache.poi.hwpf.HWPFDocument;
import org.apache.poi.ss.usermodel.BorderStyle;
//import org.apache.poi.ss.examples.CellStyleDetails;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xwpf.usermodel.ParagraphAlignment;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTableCell;
import org.apache.poi.xwpf.usermodel.XWPFTableRow;

public class ToOffice {
	static Workbook book; // книга в Excel

//метод для записи данных в файл Excel в формате xls
//первый параметр - строковый массив с заголовками колонок
//второй параметр - список с данными	
	public static void toExcel(String[] tableTitles, String[][] data, File file, String curTableName) {
		System.out.println(data);
		book = new HSSFWorkbook(); // создаем книгу
		Sheet sheet = book.createSheet(curTableName); // создаем лист
		// создадим объединение из 4-х ячеек, чтоб сделать шапку таблицы
		sheet.addMergedRegion(new CellRangeAddress( // добавляем в книгу объединенный диапазон ячеек
				0, // начальный ряд диапазона
				0, // конечный ряд диапазона
				0, // начальная колонка диапазона
				tableTitles.length // конечная колонка диапазона
		));
		Row row = sheet.createRow(0); // создаем новый ряд
		Cell cell = row.createCell(0); // создаем ячейку
		cell.setCellValue(curTableName); // задаем текст ячейки
		HSSFCellStyle cellStyle = (HSSFCellStyle) book.createCellStyle(); // создаем стиль для ряда
		cellStyle.setAlignment(HorizontalAlignment.CENTER); // задаем выравнивание по центру
		Font font = book.createFont(); // создаем шрифт для объединенных ячеек
		font.setFontHeightInPoints((short) 24); // задаем размер шрифта
		font.setColor(HSSFColor.HSSFColorPredefined.RED.getIndex()); // задаем цвет шрифта
		cellStyle.setFont(font); // добавляем шрифт к стилю
		cell.setCellStyle(cellStyle); // устанавливаем стиль на ячейку

		headCreate(sheet, tableTitles); // вызываем метод для заголовка таблицы
		dataToSheet(sheet, data); // вызываем метод для заполнения данных таблицы

		for (int i = 0; i < tableTitles.length; i++) { // цикл для установки ширины ячеек по содержимому
			sheet.autoSizeColumn(i);
		}

		// Записываем всё в файл
		try {
			book.write(new FileOutputStream(file)); // пишем книгу в файл
			book.close(); // закрываем книгу
		} catch (IOException e) {
			e.printStackTrace();
		}
		System.out.println("xls file was written successfully");
	}

	private static int headCreate(Sheet sheet, String[] tableTitles) { // метод для создания заголовка таблицы
		// Нумерация рядов начинается с нуля
		Row row = sheet.createRow(1); // создаем новый ряд
		HSSFCellStyle cellStyle = (HSSFCellStyle) book.createCellStyle(); // создаем стиль для ряда
		cellStyle.setAlignment(HorizontalAlignment.CENTER); // задаем выравнивание по центру
		cellStyle.setBorderBottom(BorderStyle.THICK); // задаем нижнюю границу
		cellStyle.setBorderLeft(BorderStyle.THICK); // и все остальные
		cellStyle.setBorderRight(BorderStyle.THICK);
		cellStyle.setBorderTop(BorderStyle.THICK);
		Font font = book.createFont(); // создаем объект для параметров шрифта
		font.setBold(true); // делаем шрифт жирным
		cellStyle.setFont(font); // устанавливаем шрифт в стиль
		for (int i = 0; i < tableTitles.length; i++) { // цикл по заголовкам
			Cell temp = row.createCell(i); // создаем ячейку
			temp.setCellStyle(cellStyle); // задаем ей стиль
			temp.setCellValue(tableTitles[i]); // задаем ей значение
		}
		return 0;
	}

	private static void dataToSheet(Sheet sheet, String[][] data) { // метод для для заполнения данных таблицы
//		int i = 1; // счетчик кол-ва рядов
		HSSFCellStyle cellStyle = (HSSFCellStyle) book.createCellStyle(); // создаем стиль для ряда
		cellStyle.setBorderBottom(BorderStyle.THIN);
		cellStyle.setBorderLeft(BorderStyle.THIN);
		cellStyle.setBorderRight(BorderStyle.THIN);
		cellStyle.setBorderTop(BorderStyle.THIN);
		for (int i = 0; i < data.length; i++) {
			Row row = sheet.createRow(i+2); // создаем новый ряд
			for (int j = 0; j < data[i].length; j++) {
				Cell temp = row.createCell(j); // задаем первую ячейку
				temp.setCellValue(data[i][j]); // вставляем туда ФИО
				temp.setCellStyle(cellStyle);
			}
		}
	}

//метод для записи данных в docx файл, параметры - массив с заголовками для столбцов и список List с содержимым таблицы 
	public static void toWordDocx(String[] tableTitles, String[][] data, File file, String curTableName) {
		XWPFDocument document = new XWPFDocument(); // создаем документ Word
		XWPFParagraph paragraph = document.createParagraph(); // создаем абзац в документе
		XWPFRun run = paragraph.createRun(); // создаем объект для записи в полученный ранее абзац
		run.setText(curTableName); // текст перед таблицей
		XWPFTable table = document.createTable(); // создаем таблицу
		XWPFTableRow tableHead = table.getRow(0); // создаем первый ряд - заголовок таблицы
		for (int i = 0; i < tableTitles.length; i++) { // цикл по заголовкам
			XWPFParagraph paragraph1 = null; // объявляем объект для параграфа
			if (i == 0) { // если первая ячейка, то сразу берем у нее параграф
				paragraph1 = tableHead.getCell(0).getParagraphs().get(0);
			} else { // иначе создаем новую ячейку и берем у нее параграф
				XWPFTableCell cell = tableHead.addNewTableCell();
				paragraph1 = cell.getParagraphs().get(0);
			}
			paragraph1.setAlignment(ParagraphAlignment.CENTER); // устанавливаем выравнивание по центру
			XWPFRun tableHeadRun = paragraph1.createRun(); //// создаем объект для записи в абзац
			tableHeadRun.setBold(true); // устанавливаем жирный шрифт
			tableHeadRun.setText(tableTitles[i]); // задаем текст
		}
		for (int i = 0; i < data.length; i++) {
			XWPFTableRow tableRow = table.createRow(); // создаем ряд
			for (int j = 0; j < data[i].length; j++) {
				tableRow.getCell(j).setText(data[i][j]); // вставляем туда ФИО
			}
		}

		try {
			FileOutputStream out = new FileOutputStream(file); // создаем файловый поток вывода с новым файлом
			document.write(out); // пишем в файл из созданного объекта
			out.close(); // закрываем потоки вывода
			document.close();
		} catch (IOException e) {
			e.printStackTrace();
		}
		System.out.println("docx file was written successfully");
	}

//	public static void toWordDoc(String[] strings, String[][] data) throws Exception {
//		HWPFDocument doc = new HWPFDocument(new FileInputStream("document.doc"));
//		FileOutputStream out = new FileOutputStream(new File("document.doc"));
//
//	}
}
