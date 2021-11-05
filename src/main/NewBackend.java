package main;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.nio.file.FileAlreadyExistsException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.StandardCopyOption;
import java.util.ArrayList;
import java.util.List;

import com.github.miachm.sods.Range;
import com.github.miachm.sods.Sheet;
import com.github.miachm.sods.SpreadSheet;

public class NewBackend {

	// Copy a contract into backup folder
	void CreateContractBackup(File contractFile, File backupFolder) throws IOException {
		Path sourcePath = Path.of(contractFile.getAbsolutePath());
		Path targetPath = Path.of(backupFolder.getAbsolutePath());
		try {
			Files.copy(sourcePath, targetPath);
		} catch (FileAlreadyExistsException e) {
			return;
		}
	}

	// Move files from backup folder into the original folder
	void RestoreBackups(File mainFolder, File backupFolder) {
		// Populates the array with names of files and directories from backupFolder
		File[] backups = backupFolder.listFiles();

		// Iterate through all backup files and move them back into the main folder
		for (File backupFile : backups) {
			File newFile = new File(mainFolder.getAbsolutePath() + "/" + backupFile.getName());
			try {
				Files.move(backupFile.toPath(), newFile.toPath(), StandardCopyOption.REPLACE_EXISTING);
			} catch (IOException e) {
				e.printStackTrace();
			}
		}
	}

	// Get a list of all contract files
	String[] GetContractFilesList(Sheet profileSheet, String contractsFolderPath, int firstRow, int lastRow) {
		List<String> filePaths = new ArrayList();

		for (int i = firstRow; i <= lastRow; i++) {
			String clientName = GetClientData(profileSheet, i)[0];

			String[] files = FindFile(contractsFolderPath, clientName);
			filePaths.add(files[0]);
		}
		return filePaths.toArray(new String[0]);

	}

	// Get client info from specified row in profile sheet (eg. Kabo5/1942 12/7)
	String[] GetClientData(Sheet sheet, int row) {
		Range cell = sheet.getRange(row - 1, 0);

		// Get client data as a string (something like Kabo5/6 20/3)
		String clientData = (String) cell.getValue();
		String[] splitClientData = clientData.split("/");

		if (splitClientData.length != 3) {
			System.out.println(clientData + " musi byæ rozdzielone znakiem / na 3 czêœci");
			return null;
		} else
			return splitClientData;
	}

	// Search for a file
	String[] FindFile(String path, String filename) {
		File dir = new File(path);

		String fileToLookFor = filename.toLowerCase() + "-";
		MyFilenameFilter filter = new MyFilenameFilter(fileToLookFor);

		String[] flist = dir.list(filter);

		return flist;
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

	// Save file into temp folder, and move its original location if everything went
	// without errors
	void SaveSafely(SpreadSheet sheet, File outputFile, File originalFolder) {
		System.out.println("Zapisywanie arkusza " + outputFile.getName());

		boolean success = true; // Has saving the file succeeded?

		// Try saving the file, catch potential errors
		try {
			sheet.save(outputFile);
		} catch (FileNotFoundException e) {
			success = false;
			System.out.println("Folder temp musi istnieæ w folderze z umowami!");
			return;
		} catch (OutOfMemoryError e) {
			success = false;
			System.out.println("Plik " + outputFile.getName() + " jest zbyt du¿y.");
			return;
		} catch (IOException e) {
			success = false;
			e.printStackTrace();
			return;
		}

		// If everything went successfully, move the file back into original folder
		if (success) {
			try {
				File newPath = new File(originalFolder + "\\" + outputFile.getName());
				Files.move(outputFile.toPath(), newPath.toPath(), StandardCopyOption.REPLACE_EXISTING);
			} catch (java.nio.file.FileSystemException e) {
				e.printStackTrace();
				System.out.println("Nie mo¿na przenieœæ pliku " + outputFile.getName()
						+ " z folderu temp bo jest on otwarty w innym programie.");
				return;
			} catch (IOException e) {
				e.printStackTrace();
				return;
			}
			System.out.println("Successfuly saved " + outputFile.getName());
		} else {
			System.out.println("Failed saving " + outputFile.getName());
		}
	}

	// Increase soldAmount in specified contract sheet by one
	void updateContractSaleAmount(Sheet sheet, int itemId) {
		double soldAmount = (double) sheet.getRange(itemId + 13, 5).getValue();
		Range soldAmountCell = sheet.getRange(itemId + 13, 5);
		soldAmountCell.setValue(soldAmount + 1);
		System.out.println("Sprzedana iloœc teraz: " + soldAmountCell.getValue());
	}

	// Download commiterAmount for a specific row
	void UpdateProfileRow(Sheet sheet, String contractsFolderPath, int row) {
		// Get client data
		String[] clientData = GetClientData(sheet, row);

		// If the cell is empty, skip
		if (clientData == null) {
			return;
		}

		// Extract client data
		String clientName = clientData[0];
		String contractSheetname = clientData[1];
		int itemId = Integer.parseInt(clientData[2]);

		// Get a list of files that match clientName
		String[] matchingContracts = FindFile(contractsFolderPath, clientName);

		// If no matching files were found
		if (matchingContracts == null) {
			System.out.println("Nie znaleziono pliku pasuj¹cego do " + clientName);
			return;
		}

		// Only one file should be found
		if (matchingContracts.length == 1) {

			// Get the only contract file
			File matchingContract = new File(contractsFolderPath + "/" + matchingContracts[0]);

			System.out.println("(" + row + ") Good! I found " + matchingContract.getPath());

			try {
				// Get contract spreadsheet and sheet to work on
				SpreadSheet contract = new SpreadSheet(matchingContract);
				Sheet contractSheet = contract.getSheet(contractSheetname);

				// Get info about the item
				String itemName = (String) contractSheet.getRange(itemId + 13, 1).getValue();
				double commiterValue = (double) contractSheet.getRange(itemId + 13, 3).getValue();

				// Update commiterValue cell in profile sheet
				Range cellToBeUpdated = sheet.getRange(row - 1, 1);
				cellToBeUpdated.setValue(commiterValue);

				System.out.println("Item's name is " + itemName);

				// Increase soldAmount in matching contract sheet
				updateContractSaleAmount(contractSheet, itemId);

				// Cleanup, then safely save the contract
				FixContractFormulas(contract);
				SaveSafely(contract, new File(contractsFolderPath + "/temp/" + matchingContracts[0]),
						new File(contractsFolderPath));

			} catch (IOException e) {
				e.printStackTrace();
			} catch (NullPointerException e) {
				System.out.println("!!!!!!!!!! B³ad przy próbie odczytu z " + matchingContract.getName());
				return;
			}

		} else if (matchingContracts.length == 0) {
			System.out.println("Nie znaleziono pliku pasuj¹cego do " + clientName);
		}
	}

	public static void main(String[] args) {
		NewBackend main = new NewBackend();
		SpreadSheet profileSheet;
		String contractsFolderPath = "res/umowy";

		int firstRow = 7;
		int lastRow = 20;

		try {
			// Get desired spreadsheet, sheet and path to folder with contracts
			profileSheet = new SpreadSheet(new File("res/PaŸdziernik 2021.ods"));
			Sheet sheet = profileSheet.getSheet(1);

			// Get all contract files
			String[] contracts = main.GetContractFilesList(sheet, contractsFolderPath, 7, 20);

			// Create a backup
			for (String contract : contracts) {
				System.out.println(contractsFolderPath + "/" + contract);

				main.CreateContractBackup(new File(contractsFolderPath + "/" + contract),
						new File(contractsFolderPath + "/backup/" + contract));
			}

			// New line break
			System.out.println();

			// Update each row, one by one
			for (int row = firstRow; row <= lastRow; row++) {
				main.UpdateProfileRow(sheet, contractsFolderPath, row);
			}

			// Clenup and save
			main.FixProfileFormulas(profileSheet);
			profileSheet.trimSheets();
			String savePath = "res/Nowy Pazdziernik.ods";
			profileSheet.save(new File(savePath));

		}
		// If profile couldn't be saved, show error message and restore all contracts
		// from backup folder
		catch (IOException e) {
			System.out.println("Zamknij plik zestawienia w OpenOfficie!");
			main.RestoreBackups(new File(contractsFolderPath), new File(contractsFolderPath + "/backup/"));
			e.printStackTrace();
		}

	}
}
