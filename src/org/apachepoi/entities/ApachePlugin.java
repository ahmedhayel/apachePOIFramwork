package org.apachepoi.entities;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.lang.reflect.Field;
import java.util.ArrayList;
import java.util.List;

import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apachepoi.entities.type.ColorsChoice;

public class ApachePlugin {

	private String path;
	private String fileName;
	private String sheetName;
	private List<?> content;
	private ColorsChoice colorHeaders;
	private ColorsChoice colorContent;

	public ApachePlugin() {
		super();
		// TODO Auto-generated constructor stub
	}

	public ApachePlugin(String path, String fileName, String sheetName, List<?> content, ColorsChoice colorHeaders,
			ColorsChoice colorContent) throws FileNotFoundException {
		super();

		this.path = path;
		this.fileName = fileName;
		this.sheetName = sheetName;
		if (content == null) {
			throw new NullPointerException("content can not be null ");
		} else {
			this.content = content;
		}

		this.colorHeaders = colorHeaders;

		this.colorContent = colorContent;
	}

	public ApachePlugin(String path, String fileName, String sheetName) {
		super();
		this.path = path;
		this.fileName = fileName;
		this.sheetName = sheetName;
	}

	public String getPath() {
		return path;
	}

	public void setPath(String path) {
		this.path = path;
	}

	public String getFileName() {
		return fileName;
	}

	public void setFileName(String fileName) {
		this.fileName = fileName;
	}

	public String getSheetName() {
		return sheetName;
	}

	public void setSheetName(String sheetName) {
		this.sheetName = sheetName;
	}

	public List<?> getContent() {
		return content;
	}

	public void setContent(List<?> content) {
		this.content = content;
	}

	public ColorsChoice getColorHeaders() {
		return colorHeaders;
	}

	public void setColorHeaders(ColorsChoice colorHeaders) {
		this.colorHeaders = colorHeaders;
	}

	public ColorsChoice getColorContent() {
		return colorContent;
	}

	public void setColorContent(ColorsChoice colorContent) {
		this.colorContent = colorContent;
	}

	public void createNewExcelFile() {

		String FILE_NAME = path + fileName;

		System.out.println("CREATION OF " + fileName + " .....");

		try {

			XSSFWorkbook workbook = new XSSFWorkbook();

			XSSFSheet sheet = workbook.createSheet(sheetName);

			// header

			XSSFRow rowHeader = sheet.createRow(0);
			Field[] fieldsHeader = content.get(0).getClass().getDeclaredFields();

			for (int i = 0; i < fieldsHeader.length; i++) {

				XSSFCell cellheader = rowHeader.createCell(i);
				cellheader.setCellValue(fieldsHeader[i].getName());
				XSSFFont font = workbook.createFont();

				font.setFontHeightInPoints((short) 18);

				font.setFontName("Goudy Old Style");

				font.setBold(true);

				font.setColor(this.color(colorHeaders));

				XSSFCellStyle style = workbook.createCellStyle();

				style.setFont(font);

				cellheader.setCellStyle(style);

			}

			// write content

			for (int i = 0; i < content.size(); i++) {
				XSSFRow rowContent = sheet.createRow(i + 1);
				Field[] fieldsContent = content.get(i).getClass().getDeclaredFields();
				for (int j = 0; j < fieldsContent.length; j++) {

					fieldsContent[j].setAccessible(true);

					XSSFCell cellContent = rowContent.createCell(j);

					cellContent.setCellValue(fieldsContent[j].get(content.get(i)).toString());
					cellContent.setCellType(this.cellType(fieldsContent[j].getType().getSimpleName()));

					XSSFFont fontContent = workbook.createFont();

					fontContent.setFontHeightInPoints((short) 14);

					fontContent.setFontName("Calibri Light");

					fontContent.setColor(this.color(colorContent));

					XSSFCellStyle style = workbook.createCellStyle();

					style.setFont(fontContent);

					cellContent.setCellStyle(style);

				}
			}

			FileOutputStream outputStream = new FileOutputStream(new File(FILE_NAME));
			workbook.write(outputStream);
			outputStream.close();
			workbook.close();
			System.out.println("CREATION OF " + fileName + " .....");
		} catch (Exception e) {
			System.err.println(e);
		}
	}

	public void EditExistingExcelFile() {

		String FILE_NAME = path + fileName;
		System.out.println("EDITING OF " + fileName + " .....");
		try {

			FileInputStream excelFile = new FileInputStream(new File(FILE_NAME));
			XSSFWorkbook workbook = new XSSFWorkbook(excelFile);

			XSSFSheet sheet = workbook.getSheet(sheetName);

			// header

			XSSFRow rowHeader = sheet.createRow(0);
			Field[] fieldsHeader = content.get(0).getClass().getDeclaredFields();

			for (int i = 0; i < fieldsHeader.length; i++) {

				XSSFCell cellheader = rowHeader.createCell(i);
				cellheader.setCellValue(fieldsHeader[i].getName());
				XSSFFont font = workbook.createFont();

				font.setFontHeightInPoints((short) 18);

				font.setFontName("Goudy Old Style");

				font.setBold(true);

				font.setColor(this.color(colorHeaders));

				XSSFCellStyle style = workbook.createCellStyle();

				style.setFont(font);

				cellheader.setCellStyle(style);

			}

			// write content

			for (int i = 0; i < content.size(); i++) {
				XSSFRow rowContent = sheet.createRow(i + 1);
				Field[] fieldsContent = content.get(i).getClass().getDeclaredFields();
				for (int j = 0; j < fieldsContent.length; j++) {

					fieldsContent[j].setAccessible(true);

					XSSFCell cellContent = rowContent.createCell(j);

					cellContent.setCellValue(fieldsContent[j].get(content.get(i)).toString());
					cellContent.setCellType(this.cellType(fieldsContent[j].getType().getSimpleName()));

					XSSFFont fontContent = workbook.createFont();

					fontContent.setFontHeightInPoints((short) 12);

					fontContent.setFontName("Calibri Light");

					fontContent.setColor(this.color(colorContent));

					XSSFCellStyle style = workbook.createCellStyle();

					style.setFont(fontContent);

					cellContent.setCellStyle(style);

				}
			}

			FileOutputStream outputStream = new FileOutputStream(FILE_NAME);
			workbook.write(outputStream);
			outputStream.close();
			workbook.close();
			System.out.println("EDITING OF " + fileName + " COMPLET");
		} catch (Exception e) {
			System.err.println("you must close the file or sheet not found");
		}
	}

	public short color(ColorsChoice color) {
		switch (color) {
		case BLACK:
			return HSSFColor.BLACK.index;
		case BLUE:
			return HSSFColor.BLUE.index;
		case YELLOW:
			return HSSFColor.YELLOW.index;
		case DARK_GREEN:
			return HSSFColor.DARK_GREEN.index;
		case SKY_BLUE:
			return HSSFColor.SKY_BLUE.index;
		case WHITE:
			return HSSFColor.WHITE.index;
		default:
			return HSSFColor.WHITE.index;
		}

	}

	public int cellType(String fieldTyepName) {
		switch (fieldTyepName) {
		case "int":
		case "long":
		case "Long":
		case "short":
		case "float":
		case "double":
			return XSSFCell.CELL_TYPE_NUMERIC;
		case "String":
		case "Date":
			return XSSFCell.CELL_TYPE_STRING;
		default:
			return XSSFCell.CELL_TYPE_STRING;
		}

	}

	public List<Object[]> getDataFromExcelFile(int columnLenght) throws IOException {
		// object result

		List<Object[]> objects = new ArrayList<Object[]>();

		Object[] object;

		String FILE_NAME = path + fileName;

		// for formatting date

		DataFormatter formatter = new DataFormatter();

		// read file

		FileInputStream fis = new FileInputStream(new File(FILE_NAME));
		XSSFWorkbook wb = new XSSFWorkbook(fis);
		XSSFSheet sheet1 = wb.getSheet(sheetName);
		int i = 0;
		for (Row row : sheet1) {
			object = new Object[columnLenght];
			for (Cell cell : row) {

				Object o = null;
				switch (cell.getCellTypeEnum()) {
				case STRING:
					o = cell.getRichStringCellValue().getString();
					// System.out.print(cell.getRichStringCellValue().getString()
					// + " \t\t");
					break;
				case NUMERIC:
					if (DateUtil.isCellDateFormatted(cell)) {
						o = cell.getDateCellValue();

						System.out.print(cell.getDateCellValue() + " \t\t");
					} else {
						o = cell.getNumericCellValue();
						System.out.print(cell.getNumericCellValue() + " \t\t");
					}
					break;
				case BOOLEAN:
					o = cell.getBooleanCellValue();
					System.out.print(cell.getBooleanCellValue() + " \t\t");
					break;
				case FORMULA:
					o = cell.getCellFormula();
					System.out.print(cell.getCellFormula() + " \t\t");
					break;
				case BLANK:
					o = "";
					System.out.print("");
					break;
				}
				object[i] = o;
				i++;
			}
			System.out.println();
			objects.add(object);
			i = 0;
		}
		return objects;
	}
}
