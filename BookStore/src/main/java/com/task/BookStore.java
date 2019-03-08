package com.task;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.PrintStream;
import java.io.UnsupportedEncodingException;
import java.util.Scanner;
import java.util.Set;
import java.util.TreeSet;

import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class BookStore {

	static Scanner scan = new Scanner(System.in);
	static int option = 0;
	static String input = "";
	static FileInputStream fis;
	static PrintStream out;
	static FileOutputStream outputStream;
	static Workbook wb;
	static DataFormatter formatter;
	static String FILE_NAME = "Book_Store.xlsx";
	static {
		try {
			fis = new FileInputStream(new File(FILE_NAME));
			wb = new XSSFWorkbook(fis);
			outputStream = new FileOutputStream(FILE_NAME);
			formatter = new DataFormatter();

		} catch (FileNotFoundException e) {
			System.out.println("The Excel file that loads all Book data not found!");
		} catch (UnsupportedEncodingException e) {
			System.out.println(e.getMessage());
		} catch (IOException e) {
			System.out.println(e.getMessage());
		}
	}

	public static void main(String[] str) throws IOException {

		BookStore BookStore = new BookStore();
		System.out.println("==== Book Manager ==== \n");
		// start process
		BookStore.showOptions();
	}

	/*
	 * This is just to show option for users and ask him to select specific
	 * option IF user enter wrong option, It should warn him to he should enter
	 * valid option.
	 */
	private void showOptions() {
		String error = "";
		System.out.println("\n 1) View all books \n" + " 2) Add a book \n" + " 3) Edit a book \n"
				+ " 4) Search for a book \n" + " 5) Save and exit");
		System.out.print("Option: ");

		input = scan.next();
		error = validateInput(input, 1, 5);

		if (error.isEmpty()) {
			option = Integer.parseInt(input);
			startProcess(option); // if option is valid start the whole process
		} else {
			System.out.println(error + "\n");
			showOptions(); // if option not valid, return user to select again
							// new option.
		}

	}

	/*
	 * This is to start the whole process.
	 * 
	 * @param int option : number that indicates user selection
	 */
	private void startProcess(int option) {

		switch (option) {
		case 1:
			viewAllBooks();
			break;
		case 2:
			addNewBook();
			break;
		case 3:
			editBook();
			break;
		case 4:
			searchForBook();
			break;
		case 5:
			saveAndExit();
			break;

		default:
			break;
		}
	}

	/*
	 * This is to search for specific keywords
	 */
	private void searchForBook() {

		System.out.println("\n=====Type in one or more keywords to search for=====");
		scan = new Scanner(System.in);
		System.out.println("Type your keywords sperated by spaces: ");

		String search = scan.nextLine();
		Sheet sheet = wb.getSheetAt(0);

		Set<Integer> matchedRows = new TreeSet<Integer>();
		String[] keywords = search.split("\\s+"); // split user input after each
													// space " "

		for (String string : keywords) {
			for (Row row : sheet) {
				String desc = row.getCell(3).toString();
				if (desc.toLowerCase().contains(string.toLowerCase())) {
					matchedRows.add(Integer.valueOf(formatter.formatCellValue(row.getCell(0))));
				}
			}
		}
		if (matchedRows.size() > 0) {
			System.out.println("\n=========Matched Records===========");

			for (Integer index : matchedRows) {
				showSpecificRecord(index);
				System.out.println();
			}
		} else {
			System.out.println("\n No matched record for your input!");
		}

		showOptions(); // Guide user to select again what he want.
	}

	/*
	 * This is to update already existed records.
	 */
	private void editBook() {

		int counter = 0;
		System.out.println("\n==== View Books ====\n");
		boolean firstRow = true;

		// print names of available records
		for (Sheet sheet : wb) {
			for (Row row : sheet) {
				if (firstRow) {
					firstRow = false;
					continue;
				}
				System.out.println("[" + ++counter + "] " + row.getCell(1));

			}
			while (true) {
				System.out.print("\n- To edit specific book, please enter the book ID, to return press <Enter>");
				editImpl(counter); // the update process implementation.
			}
		}
	}

	/*
	 * This is to add new book.
	 */
	private void addNewBook() {
		String title = "", author = "", description = "", error = "";
		scan = new Scanner(System.in);
		System.out.println("\n- Please enter the following information:");

		System.out.print("Title: ");
		title = scan.nextLine();
		error = validateInput(title, 2, 0);

		if (!error.isEmpty()) {
			System.out.println(error);
			addNewBook();
		}

		System.out.print("Author: ");
		author = scan.nextLine();
		error = validateInput(author, 2, 0);

		if (!error.isEmpty()) {
			System.out.println(error);
			addNewBook();
		}

		System.out.print("Description: ");
		description = scan.nextLine();
		error = validateInput(description, 2, 0);

		if (!error.isEmpty()) {
			System.out.println(error);
			addNewBook();
		}
		Sheet sheet = wb.getSheetAt(0);
		int newRow = sheet.getPhysicalNumberOfRows();
		sheet.createRow(newRow);
		Row row = sheet.getRow(newRow);
		row.createCell(0);
		row.getCell(0).setCellValue(newRow);
		row.createCell(1);
		row.getCell(1).setCellValue(title);
		row.createCell(2);
		row.getCell(2).setCellValue(author);
		row.createCell(3);
		row.getCell(3).setCellValue("this book: " + title + " for author: " + author + " -- " + description);

		System.out.println("\n Book [" + newRow + "] saved successfully!");
		showOptions();

	}

	/*
	 * This is to view all saved books
	 */
	private void viewAllBooks() {

		int counter = 0;
		System.out.println("\n==== View Books ====\n");
		boolean firstRow = true;
		for (Sheet sheet : wb) {
			for (Row row : sheet) {
				if (firstRow) {
					firstRow = false;
					continue;
				}
				System.out.println("[" + ++counter + "] " + row.getCell(1));

			}
			while (true) {
				System.out.print("\n- To view details enter the book ID, to return press <Enter>: ");
				showResultForViewBooks(counter);
			}

		}

	}

	private void showResultForViewBooks(int counter) {
		String err = "";
		scan = new Scanner(System.in);

		String input = scan.nextLine();
		if (input.length() > 0) {
			err = validateInput(input, 1, counter);
			if (err.isEmpty()) {
				option = Integer.valueOf(input);
				showSpecificRecord(option);
			} else {
				System.out.println("\n" + err);
				viewAllBooks();
			}

		} else {
			showOptions();
		}

	}

	private void showSpecificRecord(int option) {

		Row row = wb.getSheetAt(0).getRow(option);
		System.out.println("\n ID: " + formatter.formatCellValue(row.getCell(0)) + "\n Title: " + row.getCell(1)
				+ "\n Author: " + row.getCell(2) + "\n Description: " + row.getCell(3));

	}

	/*
	 * save and close the stream
	 */
	private void saveAndExit() {
		try {
			wb.write(outputStream);
			outputStream.close();
			wb.close();
			System.out.println("Library saved.");
			System.out.println("\n=========== THANK YOU ^_^ ===========");
			System.exit(1);
		} catch (IOException e) {
			System.out.println("Error while trying to save the file. -- " + e.getMessage());
		}

	}

	/*
	 * This is implementation of update records process
	 * 
	 * @param counter : number of all records in DB
	 */
	private void editImpl(int counter) {
		String title = "", author = "", description = "", err = "";
		boolean flag = false;// flag to indicates whether user need to change column value or no.

		scan = new Scanner(System.in);
		String input = scan.nextLine();
		if (input.length() > 0) {
			err = validateInput(input, 1, counter);

			if (!err.isEmpty()) {
				System.out.println(err + "\n");
				editBook();
			} else { // start updating
				option = Integer.valueOf(input);
				Row row = wb.getSheetAt(0).getRow(option);

				title = formatter.formatCellValue(row.getCell(1));
				author = formatter.formatCellValue(row.getCell(2));
				description = formatter.formatCellValue(row.getCell(3));

				System.out.println("\n ID: " + formatter.formatCellValue(row.getCell(0)) + "\n Title: " + row.getCell(1)
						+ "\n Author: " + row.getCell(2) + "\n Description: " + row.getCell(3));
				System.out.println("\n- Input the following information. To leave a field unchanged, hit <Enter>");
				System.out.print("\n Title [" + row.getCell(1) + "]: ");

				title = scan.nextLine();
				if (title.length() > 0) {
					row.getCell(1).setCellValue(title);
					flag = true;
				}

				System.out.print(" Author [" + row.getCell(2) + "]: ");
				author = scan.nextLine();
				if (author.length() > 0) {
					row.getCell(2).setCellValue(author);
					flag = true;
				}

				System.out.print(" Description [" + row.getCell(3) + "]: ");
				description = scan.nextLine();
				if (description.length() > 0) {
					row.getCell(3).setCellValue(
							"this book: " + row.getCell(1) + " for author: " + row.getCell(2) + " -- " + description);
					flag = true;
				}

				if (flag) {
					System.out.println("\n Book [" + option + "] updated successfully! \n");
					editBook();
				} else {
					System.out.println("\nYou didn't enter any value to be updated!!");
					editBook();
				}
			}
		} else {
			showOptions(); // Guide user to select again what he want.
		}

	}

	/*
	 * This is to validate on any input user will enter.
	 * 
	 * @param input: the value the we need to validate on.
	 * 
	 * @param type: the type of variable '1' when we want the method to deal
	 * with the var as an Integer and '2' when var as String.
	 * 
	 * @param range: This is to inforce user to enter number in specific range.
	 */
	@SuppressWarnings("unchecked")
	private <E> String validateInput(E input, int type, int range) {
		String err = "";
		try {

			if (type == 1) {
				input = (E) Integer.valueOf(input.toString());

				if (Integer.valueOf(input.toString()) > range || Integer.valueOf(input.toString()) <= 0) {
					err = "\nError, Please select number from 1 to " + range;
				}

			} else if (type == 2) {
				try {
					input = (E) Integer.valueOf(input.toString());
					err = "\nError, Please don't enter Numbers. Just Strings!";
				} catch (Exception e) {

					if (input.toString().trim().isEmpty()) {
						err = "\nError, Please don't enter Blank values. Just Strings!";
					}
				}
			}
			return err;
		} catch (Exception e) {
			if (type == 1) {
				err = "\nError, Please don't enter String or Blank values. Just Numbers!";
			} else if (type == 2) {
				err = "\nError, Please don't enter Numbers or Blank values. Just Strings!";
			}
			return err;
		}

	}
}
