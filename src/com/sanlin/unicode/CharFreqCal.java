package com.sanlin.unicode;
import java.io.BufferedReader;
import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStreamReader;
import java.io.File;
import java.util.Locale;

import jxl.CellView;
import jxl.Workbook;
import jxl.WorkbookSettings;
import jxl.format.UnderlineStyle;
import jxl.write.Label;
import jxl.write.Number;
import jxl.write.WritableCellFormat;
import jxl.write.WritableFont;
import jxl.write.WritableSheet;
import jxl.write.WritableWorkbook;
import jxl.write.WriteException;
import jxl.write.biff.RowsExceededException;

public class CharFreqCal {
	private WritableCellFormat timesBoldUnderline;
	private WritableCellFormat times;
	private String inputFile;
	private int[] freqList;
	private float totalLetters;

	public void setOutputFile(String inputFile) {
		this.inputFile = inputFile;
	}

	public void write() throws IOException, WriteException {
		File file = new File(inputFile);
		WorkbookSettings wbSettings = new WorkbookSettings();

		wbSettings.setLocale(new Locale("en", "EN"));

		WritableWorkbook workbook = Workbook.createWorkbook(file, wbSettings);
		workbook.createSheet("Report", 0);
		WritableSheet excelSheet = workbook.getSheet(0);
		createLabel(excelSheet);
		createContent(excelSheet);

		workbook.write();
		workbook.close();
	}

	private void createLabel(WritableSheet sheet) throws WriteException {
		// Lets create a times font
		WritableFont times10pt = new WritableFont(WritableFont.TIMES, 14);
		// Define the cell format
		times = new WritableCellFormat(times10pt);
		// Lets automatically wrap the cells
		times.setWrap(true);

		// create create a bold font with unterlines
		WritableFont times10ptBoldUnderline = new WritableFont(
				WritableFont.TIMES, 14, WritableFont.BOLD, false,
				UnderlineStyle.SINGLE);
		timesBoldUnderline = new WritableCellFormat(times10ptBoldUnderline);
		// Lets automatically wrap the cells
		timesBoldUnderline.setWrap(true);

		CellView cv = new CellView();
		cv.setFormat(times);
		cv.setFormat(timesBoldUnderline);
		cv.setAutosize(true);

		// Write a few headers
		addCaption(sheet, 0, 0, "HexaCode");
		addCaption(sheet, 1, 0, "Letter");
		addCaption(sheet, 2, 0, "Frequency");
		addCaption(sheet, 3, 0, "Percentage");

	}

	private void createContent(WritableSheet sheet) throws WriteException,
			RowsExceededException {
		// Write a few number
		int j = 1;
		for (int i = 1; i < freqList.length; i++) {
			// First column
			if (freqList[i] != 0) {
				addLabel(sheet, 0, j, "0x" + Integer.toHexString(i));
				addLabel(sheet, 1, j, (String.valueOf((char) i)));
				addNumber(sheet, 2, j, freqList[i]);

				addNumberFloat(sheet, 3, j,
						(freqList[i] / this.totalLetters) * 100);
				j++;
			}

		}
	}

	private void addCaption(WritableSheet sheet, int column, int row, String s)
			throws RowsExceededException, WriteException {
		Label label;
		label = new Label(column, row, s, timesBoldUnderline);
		sheet.addCell(label);
	}

	private void addNumber(WritableSheet sheet, int column, int row,
			Integer integer) throws WriteException, RowsExceededException {
		Number number;
		number = new Number(column, row, integer, times);
		sheet.addCell(number);
	}

	private void addNumberFloat(WritableSheet sheet, int column, int row,
			float floatNumber) throws WriteException, RowsExceededException {
		Number number;
		number = new Number(column, row, floatNumber, times);
		sheet.addCell(number);
	}

	private void addLabel(WritableSheet sheet, int column, int row, String s)
			throws WriteException, RowsExceededException {
		Label label;
		label = new Label(column, row, s, times);
		sheet.addCell(label);
	}

	public void countLetters(String filename) throws IOException {
		int[] freqs = new int[65535];
		int lCount = 0;

		try (BufferedReader in = new BufferedReader(new InputStreamReader(
				new FileInputStream(filename), "UTF-8"))) {
			String line;
			while ((line = in.readLine()) != null) {
				// System.out.print(line);
				for (int i = 0; i < line.length(); i++) {

					freqs[line.charAt(i) - '\u0000']++;
					lCount++;
				}
			}
		}
		this.freqList = freqs;
		this.totalLetters = lCount;
		System.out.println("Found \"" + lCount + "\" characters.");
	}

	public static void main(String[] args) throws IOException {
		if (args.length != 2) {
			System.out.println("Command syntx is -");
			System.out
					.println("java -jar CharFreqCal.jar input_file_name output_file_name\n" +
							"input_file_name\t - input text file name including '.txt' extension without 'space'.\n \t \t   eg. 'input.txt'\n" +
							"out_file_name\t - name of file output without 'space' and '.xls' extension.\n \t \t   eg. 'output_file'");
			return;
		}

		CharFreqCal freqCount = new CharFreqCal();
		freqCount.countLetters(args[0]);
		freqCount.setOutputFile(args[1] + ".xls");

		try {
			freqCount.write();
			System.out.println("File writed to " + freqCount.inputFile);
		} catch (WriteException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}

	}
}