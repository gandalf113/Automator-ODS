package main;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.StandardCopyOption;
import java.util.ArrayList;
import java.util.List;

import com.github.miachm.sods.Range;
import com.github.miachm.sods.Sheet;
import com.github.miachm.sods.SpreadSheet;

// Get corresponding contract file and sheet 

public class Backend {
	String[] FindFile(String path, String filename) {
		File dir = new File(path);
		MyFilenameFilter filter = new MyFilenameFilter(filename.toLowerCase());

		String[] flist = dir.list(filter);

		return flist;
	}

	void setCommiterValue(Sheet sheet, int row, double value) {
		Range cellToBeUpdated = sheet.getRange(row - 1, 1); // Cell in profile sheet
		cellToBeUpdated.setValue(value);
	}

	String GetClientData(Sheet sheet, int row) {
		Range cell = sheet.getRange(row - 1, 0);

		// Get client data as a string (something like Kabo5/6 20/3)
		String clientData = (String) cell.getValue();

		return clientData;
	}

	void UpdateProfileRow(Sheet sheet, int row) {
		String clientData = GetClientData(sheet, row);

		System.out.println(clientData);

		// Extract client data from cell
		String[] clientDataSplit = clientData.split("/");

		String contractFilename = clientDataSplit[0];
		String contractSheetname = clientDataSplit[1];
		int itemId = Integer.parseInt(clientDataSplit[2]);

		System.out.println("");

		// Find corresponding contract file (or files) by contractFilename
		String path = "res/Zbiór umów komisowych ods";
		String[] files = FindFile(path, contractFilename);

		// If any files are found
		if (files != null) {

			// Check if only a single file has been found
			if (files.length > 0) {
				File matchingContractFile = new File(path + "\\" + files[0]);

				System.out.println("Szukam w pliku: " + files[0]);
				System.out.println("");

				boolean success = true; // has saving the file succeeded?
				File newContractFile = new File(path + "\\temp\\" + files[0]);

				// Get commiter value from contract sheet and paste it into profile
				try {
					SpreadSheet contract = new SpreadSheet(matchingContractFile);

					Sheet contractSheet = contract.getSheet(contractSheetname);

					String itemName = (String) contractSheet.getRange(itemId + 13, 1).getValue();
					double commiterValue = (double) contractSheet.getRange(itemId + 13, 3).getValue();

					System.out.println("Nazwa przedmiotu: " + itemName);
					System.out.println("Kwota dla komitenta: " + commiterValue);

					setCommiterValue(sheet, row, commiterValue);

					try {
						updateContractSaleAmount(contractSheet, itemId);
						FixContractFormulas(contract);
//						contract.trimSheets();
						contract.save(newContractFile);
					} catch (OutOfMemoryError e) {
						success = false;
						System.out.println(files[0] + " jest zbyt du¿y.");
					}
				} catch (IOException e) {
					e.printStackTrace();
				}

				if (success) {

					try {
						Files.move(newContractFile.toPath(), matchingContractFile.toPath(),
								StandardCopyOption.REPLACE_EXISTING);
					} catch (IOException e) {
						e.printStackTrace();
					}

					System.out.println("Success!!");
				} else {
					System.out.println("Fail!!");
				}
			}
		} else {
			System.out.println("Znaleziono wiêcej ni¿ jeden plik!");

			System.out.println("Znalezione pliki: ");

			for (int i = 0; i < files.length; i++) {
				System.out.println(files[i]);
			}
		}
	}

	void updateContractSaleAmount(Sheet sheet, int itemId) {
		double soldAmount = (double) sheet.getRange(itemId + 13, 5).getValue();
		Range soldAmountCell = sheet.getRange(itemId + 13, 5);
		soldAmountCell.setValue(soldAmount + 1);
		System.out.println("Sprzedana iloœc teraz: " + soldAmountCell.getValue());
	}

	void updateRows(File profileFile, String sheetName, int firstRow, int lastRow) {
		try {
			SpreadSheet profile = new SpreadSheet(profileFile);
			Sheet profileSheet = profile.getSheet(sheetName);

			System.out.println(profileSheet.getName());

			for (int i = firstRow; i <= lastRow; i++) {
				UpdateProfileRow(profileSheet, i);
				System.out.println("---------------------------");
			}

			FixProfileFormulas(profile);
			profile.save(new File("res/Kwiecieñ 2020.ods"));
		} catch (IOException e) {
			if (e instanceof FileNotFoundException) {
				System.out.println("Zamknij plik zestawienia w OpenOfficie!");
			}
			e.printStackTrace();
		}
	}

	void FixProfileFormulas(SpreadSheet profile) {
		Sheet profileFirstSheet = profile.getSheet(0);
		Range range = profileFirstSheet.getRange(2, 0, 32, 4);

		String[][] formulas = range.getFormulas();

		// Fix formulas on the first sheet
		for (int i = 0; i < formulas.length; i++) {
			for (int j = 0; j < formulas[i].length; j++) {
				String formula = formulas[i][j];
				if (formula == null)
					continue;

				if (formula.contains("."))
					continue;
				String newFormula = addChar(formula, '.', formula.length() - 2);

				formulas[i][j] = newFormula;
				System.out.println(formulas[i][j]);
			}
		}

		// Fix formulas on all the other sheets
		for (int i = 1; i < profile.getNumSheets() - 1; i++) {
			Sheet sheet = profile.getSheet(i);

			// Upper cells (L6-M25)
			Range r1 = sheet.getRange(5, 11, 20, 1);
			String[][] f1s = r1.getFormulas();

			for (int j = 0; j < f1s.length; j++) {
				for (int j2 = 0; j2 < f1s[j].length; j2++) {
					String formula = f1s[j][j2];
					String newFormula = formula.replace(",", "") + "*100";
					f1s[j][j2] = newFormula;
				}
			}

			// Lower cells (L32-M51)
			Range r2 = sheet.getRange(31, 11, 20, 1);
			String[][] f2s = r2.getFormulas();

			for (int j = 0; j < f2s.length; j++) {
				for (int j2 = 0; j2 < f2s[j].length; j2++) {
					String formula = f2s[j][j2];
					String newFormula = formula.replace(",", "") + "*100";
					f2s[j][j2] = newFormula;
				}
			}
			r1.setFormulas(f1s);
			r2.setFormulas(f2s);
		}
		range.setFormulas(formulas);
	}

	void FixContractFormulas(SpreadSheet contract) {
		Sheet contractFirstSheet = contract.getSheet(0);

		Range cell = contractFirstSheet.getRange(3, 4);

		if (cell.getFormula() == null)
			return;

		if (cell.getFormula().contains("."))
			return;

		String[] splitFormula = cell.getFormula().split(":");

		splitFormula[0] = addChar(splitFormula[0], '.', splitFormula[0].length() - 3);
		splitFormula[1] = addChar(splitFormula[1], '.', splitFormula[1].length() - 4);

		String newFormula = splitFormula[0] + ":" + splitFormula[1];
		cell.setFormula(newFormula);
	}

	public String addChar(String str, char ch, int position) {
		int len = str.length();
		char[] updatedArr = new char[len + 1];
		str.getChars(0, position, updatedArr, 0);
		updatedArr[position] = ch;
		str.getChars(position, len, updatedArr, position + 1);
		return new String(updatedArr);
	}

	public List<String> GetSheets(String filepath) {
		List<String> sheets = new ArrayList<String>();
		
		try {
			SpreadSheet workbook = new SpreadSheet(new File(filepath));

			for (Sheet sheet : workbook.getSheets()) {
				sheets.add(sheet.getName());
			}
		} catch (IOException e) {
			e.printStackTrace();
		}
		
		return sheets;

	}

}
