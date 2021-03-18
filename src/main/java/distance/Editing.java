package distance;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Scanner;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Sheet;

public class Editing {

	// method to check if row exists when user inputs new data data for each iron
	// if row already exists, then method checks last row used by the iron
	// if row does not exist, create new row
	public static int lastFreeCell(HSSFSheet methSheet, int column) {
		int rowNum = methSheet.getLastRowNum();
		if (rowNum == 0)
			return 1;
		HSSFRow currentRow = methSheet.getRow(rowNum);
		while (currentRow.getCell(column) == null) {
			rowNum--;
			currentRow = methSheet.getRow(rowNum);
		}
		return rowNum + 1;
	}

	// method to calculate the averages of each iron
	public static double returnAverage(HSSFSheet avgSheet, int column) {
		int rowNum = avgSheet.getLastRowNum();
		double div = 0;
		double total = 0;
		double average = 0;
		if (rowNum == 0) {
			if (avgSheet.getRow(1) == null) {
				return 0;
			}
			total = Double.parseDouble(avgSheet.getRow(1).getCell(column).getStringCellValue());
			return total;
		}
		HSSFRow currentRow = avgSheet.getRow(rowNum);
		while (currentRow.getRowNum() > 0) {
			if (currentRow.getCell(column) != null) {
				div++;
				total += Double.parseDouble(avgSheet.getRow(currentRow.getRowNum()).getCell(column).getStringCellValue());
				average = total / div;
			} else {
				average = 0;
			}
			rowNum--;
			currentRow = avgSheet.getRow(rowNum);
		}
		return average;
	}

	// method to take and check that input string consists only digits (or empty for
	// no data entered)
	public static String[] input(Scanner input, String[] str) {
		String[] strArray = input.nextLine().split(" ");

		boolean checkNums = false;
		while (checkNums == false) {
			for (int i = 0; i < strArray.length; i++) {
				if (strArray[i].matches("[0-9]+") || strArray[i].isEmpty()) {
					checkNums = true;
				} else {
					System.out.println("Please use only digits");
					strArray = input.nextLine().split(" ");
					checkNums = false;
				}
			}
		}
		return strArray;
	}

	// method to delete a player's sheet
	public static void delete(File inputFileDelete) throws IOException {
		FileInputStream deleteFile = new FileInputStream(inputFileDelete);
		HSSFWorkbook workbook = new HSSFWorkbook(deleteFile);

		int count = 0;
		for (Sheet individualSheet : workbook) {
			System.out.println("Sheet number " + count + " name is :	" + individualSheet.getSheetName());
			count++;
		}
		System.out.println("Which player would you like to delete? ");
		Scanner input = new Scanner(System.in);
		String delete = input.nextLine();
		Sheet playerSheet = workbook.getSheet(delete);
		while (playerSheet == null) {
			System.out.println("player name is invalid, please input valid player name");
			delete = input.nextLine();
			playerSheet = workbook.getSheet(delete);
		}
		int deleteSheetAt = workbook.getSheetIndex(delete);
		workbook.removeSheetAt(deleteSheetAt);

		FileOutputStream output = new FileOutputStream(inputFileDelete);
		workbook.write(output);
		output.close();
		workbook.close();
		System.out.println(delete + " has been deleted");
	}

	// method to edit player's data
	public static void edit(File inputFile) throws IOException {
		FileInputStream getInfo = new FileInputStream(inputFile);
		HSSFWorkbook workbook = new HSSFWorkbook(getInfo);

		int count = 0;
		for (Sheet individualSheet : workbook) {
			System.out.println("Sheet number " + count + " name is :	" + individualSheet.getSheetName());
			count++;
		}

		System.out.println("Which player would you like to edit? ");
		Scanner input = new Scanner(System.in);
		String decision = input.nextLine();
		Sheet playerSheet = workbook.getSheet(decision);
		while (playerSheet == null) {
			System.out.println("player name is invalid, please input valid player name");
			decision = input.nextLine();
			playerSheet = workbook.getSheet(decision);
		}
		System.out.println("Valid Player Name \"" + decision + "\" found. Sheet is now open for editing \n");
		System.out.println("Please enter distance for the following clubs: ");
		System.out.println("Seperate distances by space for each iron ");
		System.out.println("If done entering distnces for the specified iron, or would like to skip, hit enter \n ");
		HSSFSheet currentSheet = workbook.getSheet(decision);

		// HSSFSheet currentSheet = playerList();
		/////////////////////////////////////////// print 4 iron
		System.out.println(currentSheet.getRow(0).getCell(0));

		String[] strArray = null;
		String[] Array = input(input, strArray);

		for (String element : Array) {
			int rowNumber = currentSheet.getLastRowNum();
			if (element.isEmpty())
				continue;
			else {
				if (currentSheet.getRow(rowNumber).getCell(0) != null) {
					rowNumber++;
					HSSFRow row = (HSSFRow) playerSheet.createRow((short) rowNumber);
					row.createCell(0).setCellValue(element);
				} else {
					rowNumber = lastFreeCell(currentSheet, 0);
					currentSheet.getRow(rowNumber).createCell(0).setCellValue(element);
				}
			}
		}
		double average = returnAverage(currentSheet, 0);
		if (average == 0) {
			System.out.println("No data yet");
		} else {
			System.out.println("The average is " + average);
		}

		// System.out.println("The average is " +returnAverage(currentSheet,0));
		System.out.println("********************************");

		/////////////////////////////////////////// print 5 iron
		System.out.println(currentSheet.getRow(0).getCell(1));

		Array = input(input, strArray);

		for (String element : Array) {
			int rowNumber = currentSheet.getLastRowNum();
			if (element.isEmpty())
				continue;
			else {
				if (currentSheet.getRow(rowNumber).getCell(1) != null) {
					rowNumber++;
					HSSFRow row = (HSSFRow) playerSheet.createRow((short) rowNumber);
					row.createCell(1).setCellValue(element);
				} else {
					rowNumber = lastFreeCell(currentSheet, 1);
					currentSheet.getRow(rowNumber).createCell(1).setCellValue(element);
				}
			}
		}

		average = returnAverage(currentSheet, 1);
		if (average == 0) {
			System.out.println("No data yet");
		} else {
			System.out.println("The average is " + average);
		}

		// System.out.println("The average is " +returnAverage(currentSheet,1));
		System.out.println("********************************");

		/////////////////////////////////////////// print 6 iron
		System.out.println(currentSheet.getRow(0).getCell(2));

		Array = input(input, strArray);

		for (String element : Array) {
			int rowNumber = currentSheet.getLastRowNum();
			if (element.isEmpty())
				continue;
			else {
				if (currentSheet.getRow(rowNumber).getCell(2) != null) {
					rowNumber++;
					HSSFRow row = (HSSFRow) playerSheet.createRow((short) rowNumber);
					row.createCell(2).setCellValue(element);
				} else {
					rowNumber = lastFreeCell(currentSheet, 2);
					currentSheet.getRow(rowNumber).createCell(2).setCellValue(element);
				}
			}
		}

		average = returnAverage(currentSheet, 2);
		if (average == 0) {
			System.out.println("No data yet");
		} else {
			System.out.println("The average is " + average);
		}

		// System.out.println("The average is " +returnAverage(currentSheet,2));
		System.out.println("********************************");

		/////////////////////////////////////////// print 7 iron
		System.out.println(currentSheet.getRow(0).getCell(3));

		Array = input(input, strArray);

		for (String element : Array) {
			int rowNumber = currentSheet.getLastRowNum();
			if (element.isEmpty())
				continue;
			else {
				if (currentSheet.getRow(rowNumber).getCell(3) != null) {
					rowNumber++;
					HSSFRow row = (HSSFRow) playerSheet.createRow((short) rowNumber);
					row.createCell(3).setCellValue(element);
				} else {
					rowNumber = lastFreeCell(currentSheet, 3);
					currentSheet.getRow(rowNumber).createCell(3).setCellValue(element);
				}
			}
		}

		average = returnAverage(currentSheet, 3);
		if (average == 0) {
			System.out.println("No data yet");
		} else {
			System.out.println("The average is " + average);
		}

		// System.out.println("The average is " +returnAverage(currentSheet,3));
		System.out.println("********************************");

		/////////////////////////////////////////// print 8 iron
		System.out.println(currentSheet.getRow(0).getCell(4));

		Array = input(input, strArray);

		for (String element : Array) {
			int rowNumber = currentSheet.getLastRowNum();
			if (element.isEmpty())
				continue;
			else {
				if (currentSheet.getRow(rowNumber).getCell(4) != null) {
					rowNumber++;
					HSSFRow row = (HSSFRow) playerSheet.createRow((short) rowNumber);
					row.createCell(4).setCellValue(element);
				} else {
					rowNumber = lastFreeCell(currentSheet, 4);
					currentSheet.getRow(rowNumber).createCell(4).setCellValue(element);
				}
			}
		}
		average = returnAverage(currentSheet, 4);
		if (average == 0) {
			System.out.println("No data yet");
		} else {
			System.out.println("The average is " + average);
		}

		// System.out.println("The average is " +returnAverage(currentSheet,4));
		System.out.println("********************************");

		/////////////////////////////////////////// print 9 iron
		System.out.println(currentSheet.getRow(0).getCell(5));

		Array = input(input, strArray);

		for (String element : Array) {
			int rowNumber = currentSheet.getLastRowNum();
			if (element.isEmpty())
				continue;
			else {
				if (currentSheet.getRow(rowNumber).getCell(5) != null) {
					rowNumber++;
					HSSFRow row = (HSSFRow) playerSheet.createRow((short) rowNumber);
					row.createCell(5).setCellValue(element);
				} else {
					rowNumber = lastFreeCell(currentSheet, 5);
					currentSheet.getRow(rowNumber).createCell(5).setCellValue(element);
				}
			}
		}

		average = returnAverage(currentSheet, 5);
		if (average == 0) {
			System.out.println("No data yet");
		} else {
			System.out.println("The average is " + average);
		}

		// System.out.println("The average is " +returnAverage(currentSheet,5));
		System.out.println("********************************");

		/////////////////////////////////////////// print p iron
		System.out.println(currentSheet.getRow(0).getCell(6));

		Array = input(input, strArray);

		for (String element : Array) {
			int rowNumber = currentSheet.getLastRowNum();
			if (element.isEmpty())
				continue;
			else {
				if (currentSheet.getRow(rowNumber).getCell(6) != null) {
					rowNumber++;
					HSSFRow row = (HSSFRow) playerSheet.createRow((short) rowNumber);
					row.createCell(6).setCellValue(element);
				} else {
					rowNumber = lastFreeCell(currentSheet, 6);
					currentSheet.getRow(rowNumber).createCell(6).setCellValue(element);
				}
			}
		}

		average = returnAverage(currentSheet, 6);
		if (average == 0) {
			System.out.println("No data yet");
		} else {
			System.out.println("The average is " + average);
		}

		// System.out.println("The average is " +returnAverage(currentSheet,6));
		System.out.println("********************************");

		FileOutputStream editWrite = new FileOutputStream(inputFile);
		workbook.write(editWrite);
		editWrite.close();
		workbook.close();

		// if you close, main class won't work as intended
		// input.close();

	}

}
