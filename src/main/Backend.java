package main;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.nio.file.FileSystemException;
import java.nio.file.Files;
import java.nio.file.StandardCopyOption;
import java.util.ArrayList;
import java.util.List;

import com.github.miachm.sods.Range;
import com.github.miachm.sods.Sheet;
import com.github.miachm.sods.SpreadSheet;

// Get corresponding contract file and sheet 

public class Backend {
	private Main main;
	private List<String> dialogMessages = new ArrayList<String>();

	private List<SpreadSheet> contractsToBeSaved = new ArrayList<SpreadSheet>();
	private List<File> contractTempPathsToBeSaved = new ArrayList<File>();
	private List<File> contractFinalPathsToBeSaved = new ArrayList<File>();

	private boolean tempFolderNotFound;

	public Backend(Main main) {
		this.main = main;
	}

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

	void UpdateProfileRow(Sheet sheet, String contractsDirPath, int row) {
		String clientData = GetClientData(sheet, row);

		System.out.println(clientData);

		main.setMessage("Szukam dla " + clientData);

		// Extract client data from cell
		String[] clientDataSplit = clientData.split("/");

		String contractFilename = "", contractSheetname = "";
		int itemId = 0;

		try {
			contractFilename = clientDataSplit[0];
			contractSheetname = clientDataSplit[1];
			itemId = Integer.parseInt(clientDataSplit[2]);
		} catch (ArrayIndexOutOfBoundsException e) {
			System.out.println("Wyraz " + clientData + " nie jest prawid³owy. Rozdziel dane klienta znakiem /");
			dialogMessages.add("Wyraz " + clientData + " nie jest prawid³owy. Rozdziel dane klienta znakiem /");
		}

		System.out.println("");

		// Find corresponding contract file (or files) by contractFilename
		String path = contractsDirPath;
		String[] files = FindFile(path, contractFilename);

		// If any files are found
		if (files != null) {

			// Check if only a single file has been found
			if (files.length > 0 && files.length <= 1) {
				File matchingContractFile = new File(path + "\\" + files[0]);

				System.out.println("Szukam w pliku: " + files[0]);
				System.out.println("");

				File newContractFile = new File(path + "\\temp\\" + files[0]);
				tempFolderNotFound = false;

				// Get commiter value from contract sheet and paste it into profile
				try {
					SpreadSheet contract = new SpreadSheet(matchingContractFile);

					Sheet contractSheet = contract.getSheet(contractSheetname);

					String itemName = (String) contractSheet.getRange(itemId + 13, 1).getValue();
					double commiterValue = (double) contractSheet.getRange(itemId + 13, 3).getValue();

					System.out.println("Nazwa przedmiotu: " + itemName);
					System.out.println("Kwota dla komitenta: " + commiterValue);

					setCommiterValue(sheet, row, commiterValue);

					updateContractSaleAmount(contractSheet, itemId);
					FixContractFormulas(contract);

					// Add contract to a list so it can be saved at the end
					contractsToBeSaved.add(contract);
					contractTempPathsToBeSaved.add(newContractFile);
					contractFinalPathsToBeSaved.add(matchingContractFile);

				} catch (IOException e) {
					dialogMessages.add("Nie mogê odnalezæ pliku umowy dla " + contractFilename);
					e.printStackTrace();
				}

			} else if (files.length > 1) {
				dialogMessages.add("Znaleziono kilka plików pasuj¹cych do " + contractFilename + ". Nie wiem, który wybraæ!");
			}
		} else {
			dialogMessages.add("Nie znaleziono pliku pasuj¹cego do " + contractFilename);
		}
	}

	void updateContractSaleAmount(Sheet sheet, int itemId) {
		double soldAmount = (double) sheet.getRange(itemId + 13, 5).getValue();
		Range soldAmountCell = sheet.getRange(itemId + 13, 5);
		soldAmountCell.setValue(soldAmount + 1);
		System.out.println("Sprzedana iloœc teraz: " + soldAmountCell.getValue());
	}

	void updateRows(File profileFile, String contractsDirPath, String sheetName, int firstRow, int lastRow) {
		try {
			SpreadSheet profile = new SpreadSheet(profileFile);
			Sheet profileSheet = profile.getSheet(sheetName);

			int n = 0; // track progress
			for (int i = firstRow; i <= lastRow; i++) {
				int progress;
				if (lastRow - firstRow != 0) {
					progress = n * 100 / (lastRow - firstRow);
				} else {
					progress = 100;
				}
				System.out.println(progress);
				main.setProgress(progress);
				UpdateProfileRow(profileSheet, contractsDirPath, i);
				n++;
				System.out.println("---------------------------");

				if (tempFolderNotFound) {
					dialogMessages.add("Folder temp musi istnieæ w folderze z umowami!");
					main.setProgress(0);
					main.setMessage("Nie znaleziono folderu temp");
					main.showFinalErrors(dialogMessages);
					return;
				}
			}

			FixProfileFormulas(profile);
			profile.trimSheets();
			profile.save(new File("res/Kwiecieñ 2020.ods"));

			// Save all contract files
			for (int i = 0; i < contractsToBeSaved.size(); i++) {
				SpreadSheet contract = contractsToBeSaved.get(i);
				File outputFile = contractTempPathsToBeSaved.get(i);

				main.setMessage("Zapisywanie umowy " + outputFile.getName());

				boolean success = true; // has saving the file succeeded?

				try {
					contract.save(outputFile);
				} catch (FileNotFoundException e) {
					success = false;
					System.out.println("Folder temp musi istnieæ w folderze z umowami!");
					tempFolderNotFound = true;
				} catch (OutOfMemoryError e) {
					success = false;
					System.out.println("Plik " + outputFile.getName() + " jest zbyt du¿y.");
					dialogMessages.add("Plik " + outputFile.getName() + " jest zbyt du¿y.");
				}

				if (success) {
					try {
						Files.move(outputFile.toPath(), contractFinalPathsToBeSaved.get(i).toPath(),
								StandardCopyOption.REPLACE_EXISTING);
					} catch (java.nio.file.FileSystemException e) {
						dialogMessages.add("Nie mo¿na zapisaæ pliku " + outputFile.getName()
								+ " z folderu temp bo jest on otwarty w innym programie.");
					} catch (IOException e) {
						e.printStackTrace();
					}
					System.out.println("Successfuly saved " + contractFinalPathsToBeSaved.get(i).getName());
				} else {
					System.out.println("Failed saving " + contractFinalPathsToBeSaved.get(i).getName());
				}
			}

			main.setMessage("Gotowe!");
		} catch (IOException e) {
			if (e instanceof FileNotFoundException) {
				dialogMessages.add("Zamknij plik zestawienia w OpenOfficie!");
				main.setMessage("Zamknij wszystkie pliki w OpenOfficie i spróbuj ponownie.");
			}
		}
		contractsToBeSaved.clear();
		contractTempPathsToBeSaved.clear();
		contractFinalPathsToBeSaved.clear();
		main.showFinalErrors(dialogMessages);
	}

	void FixProfileFormulas(SpreadSheet profile) {
		Sheet profileFirstSheet = profile.getSheet(0);
		Range range01 = profileFirstSheet.getRange(2, 0, 32, 4);
		Range range02 = profileFirstSheet.getRange(9, 8, 1, 5);

		String[][] formulas01 = range01.getFormulas();
		String[][] formulas02 = range02.getFormulas();

		// Fix formulas on the first sheet
		for (int i = 0; i < formulas01.length; i++) {
			for (int j = 0; j < formulas01[i].length; j++) {
				String formula = formulas01[i][j];
				if (formula == null)
					continue;

				if (formula.contains("."))
					continue;
				String newFormula = addChar(formula, '.', formula.length() - 2);

				formulas01[i][j] = newFormula;
//				System.out.println(formulas01[i][j]);
			}
		}

		for (int i = 0; i < formulas02.length; i++) {
			for (int j = 0; j < formulas02[i].length; j++) {
				String formula = formulas02[i][j];
				if (formula == null)
					continue;

				if (formula.contains("."))
					continue;
				String newFormula = addChar(formula, '.', formula.length() - 3);
				newFormula = addChar(newFormula, '.', formula.length() - 14);

				formulas02[i][j] = newFormula;
				System.out.println(formulas02[i][j]);
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
					if (formula.contains("*100"))
						continue;
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
					if (formula.contains("*100"))
						continue;
					String newFormula = formula.replace(",", "") + "*100";
					f2s[j][j2] = newFormula;
				}
			}
			r1.setFormulas(f1s);
			r2.setFormulas(f2s);
		}
		range01.setFormulas(formulas01);
		range02.setFormulas(formulas02);
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
