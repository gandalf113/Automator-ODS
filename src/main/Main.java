package main;

import java.awt.EventQueue;
import java.awt.GridBagConstraints;
import java.awt.GridBagLayout;
import java.awt.Insets;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.io.File;
import java.util.Iterator;
import java.util.List;
import java.util.concurrent.Executors;

import javax.swing.DefaultComboBoxModel;
import javax.swing.JButton;
import javax.swing.JComboBox;
import javax.swing.JFileChooser;
import javax.swing.JFrame;
import javax.swing.JLabel;
import javax.swing.JOptionPane;
import javax.swing.JPanel;
import javax.swing.JProgressBar;
import javax.swing.JSpinner;
import javax.swing.JTextField;
import javax.swing.SpinnerModel;
import javax.swing.SpinnerNumberModel;

public class Main {

	private JFrame frmAutomatorOds;
	private Backend backend;
	private JTextField contractFolderTextField, profileFileTextField;
	private JProgressBar progressBar;
	private JLabel messageLabel;

	private String profileFilePath, contractFolderPath, sheetName;

	/**
	 * Launch the application.
	 */
	public static void main(String[] args) {
		EventQueue.invokeLater(new Runnable() {
			public void run() {
				try {
					Main window = new Main();
					window.frmAutomatorOds.setVisible(true);
				} catch (Exception e) {
					e.printStackTrace();
				}
			}
		});
	}

	public void setMessage(String text) {
		messageLabel.setText(text);
	}

	public void showSuccessDialog() {
		JOptionPane.showMessageDialog(frmAutomatorOds, "Gotowe. Proces przebieg³ pomyœlnie.", "Sukces",
				JOptionPane.INFORMATION_MESSAGE);
	}

	public void setProgress(int percentile) {
		progressBar.setValue(percentile);
	}

	public void showFinalErrors(List<String> messages) {
		if (messages.size() > 0) {
			String finalMessage = "";
			System.out.println("Errors!");
			for (String message : messages) {
				System.out.println(message);
				finalMessage += message + "\n";
			}
			showErrorDialog(finalMessage);
			messages.clear();
		} else {
			showSuccessDialog();
		}
	}

	public void showErrorDialog(String message) {
		JOptionPane.showMessageDialog(frmAutomatorOds, message, "Komunikat o b³êdach", JOptionPane.ERROR_MESSAGE);
	}

	/**
	 * Create the application.
	 */
	public Main() {
		initialize();
	}

	/**
	 * Initialize the contents of the frame.
	 */
	private void initialize() {
		backend = new Backend(this);

		frmAutomatorOds = new JFrame();
		frmAutomatorOds.setTitle("Automator ODS - Komis Olu\u015B");
		frmAutomatorOds.setResizable(false);
		frmAutomatorOds.setBounds(100, 100, 640, 480);
		frmAutomatorOds.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
		frmAutomatorOds.getContentPane().setLayout(null);

		JButton btnNewButton = new JButton("Aktualizuj arkusze");
		btnNewButton.setBounds(253, 390, 136, 32);
		frmAutomatorOds.getContentPane().add(btnNewButton);

		JPanel panel = new JPanel();
		panel.setBounds(93, 137, 429, 172);
		frmAutomatorOds.getContentPane().add(panel);
		GridBagLayout gbl_panel = new GridBagLayout();
		gbl_panel.columnWidths = new int[] { 162, 170, 70, 0 };
		gbl_panel.rowHeights = new int[] { 23, 20, 0, 0, 0, 0, 0 };
		gbl_panel.columnWeights = new double[] { 1.0, 1.0, 0.0, Double.MIN_VALUE };
		gbl_panel.rowWeights = new double[] { 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, Double.MIN_VALUE };
		panel.setLayout(gbl_panel);

		JLabel profileFileLabel = new JLabel("Plik zestawienia sprzeda\u017Cy: ");
		GridBagConstraints gbc_profileFileLabel = new GridBagConstraints();
		gbc_profileFileLabel.anchor = GridBagConstraints.WEST;
		gbc_profileFileLabel.insets = new Insets(0, 0, 5, 5);
		gbc_profileFileLabel.gridx = 0;
		gbc_profileFileLabel.gridy = 0;
		panel.add(profileFileLabel, gbc_profileFileLabel);

		profileFileTextField = new JTextField();
		profileFileTextField.setEditable(false);
		GridBagConstraints gbc_profileFileTextField = new GridBagConstraints();
		gbc_profileFileTextField.insets = new Insets(0, 0, 5, 5);
		gbc_profileFileTextField.fill = GridBagConstraints.BOTH;
		gbc_profileFileTextField.gridx = 1;
		gbc_profileFileTextField.gridy = 0;
		panel.add(profileFileTextField, gbc_profileFileTextField);
		profileFileTextField.setColumns(10);

		JButton openProfileFileButton = new JButton("Otw\u00F3rz");
		GridBagConstraints gbc_openProfileFileButton = new GridBagConstraints();
		gbc_openProfileFileButton.fill = GridBagConstraints.BOTH;
		gbc_openProfileFileButton.insets = new Insets(0, 0, 5, 0);
		gbc_openProfileFileButton.gridx = 2;
		gbc_openProfileFileButton.gridy = 0;
		panel.add(openProfileFileButton, gbc_openProfileFileButton);

		JLabel contractFolderLabel = new JLabel("Folder z umowami: ");
		GridBagConstraints gbc_contractFolderLabel = new GridBagConstraints();
		gbc_contractFolderLabel.insets = new Insets(0, 0, 5, 5);
		gbc_contractFolderLabel.anchor = GridBagConstraints.WEST;
		gbc_contractFolderLabel.gridx = 0;
		gbc_contractFolderLabel.gridy = 1;
		panel.add(contractFolderLabel, gbc_contractFolderLabel);

		contractFolderTextField = new JTextField();
		contractFolderTextField.setEditable(false);
		GridBagConstraints gbc_contractFolderTextField = new GridBagConstraints();
		gbc_contractFolderTextField.insets = new Insets(0, 0, 5, 5);
		gbc_contractFolderTextField.fill = GridBagConstraints.BOTH;
		gbc_contractFolderTextField.gridx = 1;
		gbc_contractFolderTextField.gridy = 1;
		panel.add(contractFolderTextField, gbc_contractFolderTextField);
		contractFolderTextField.setColumns(10);

		// Wybór folderu z umowami
		JButton openContractFolderButton = new JButton("Otw\u00F3rz");
		openContractFolderButton.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent e) {
				JFileChooser fileChooser = new JFileChooser();
				fileChooser.setFileSelectionMode(JFileChooser.DIRECTORIES_ONLY);
				int result = fileChooser.showOpenDialog(panel);
				if (result == JFileChooser.APPROVE_OPTION) {
					File selectedFile = fileChooser.getSelectedFile();
					contractFolderPath = selectedFile.getAbsolutePath();
					contractFolderTextField.setText(contractFolderPath);
				}
			}
		});
		GridBagConstraints gbc_openContractFolderButton = new GridBagConstraints();
		gbc_openContractFolderButton.fill = GridBagConstraints.BOTH;
		gbc_openContractFolderButton.insets = new Insets(0, 0, 5, 0);
		gbc_openContractFolderButton.gridx = 2;
		gbc_openContractFolderButton.gridy = 1;
		panel.add(openContractFolderButton, gbc_openContractFolderButton);

		JLabel profileSheetNameLabel = new JLabel("Nazwa arkusza: ");
		GridBagConstraints gbc_profileSheetNameLabel = new GridBagConstraints();
		gbc_profileSheetNameLabel.insets = new Insets(0, 0, 5, 5);
		gbc_profileSheetNameLabel.anchor = GridBagConstraints.WEST;
		gbc_profileSheetNameLabel.gridx = 0;
		gbc_profileSheetNameLabel.gridy = 3;
		panel.add(profileSheetNameLabel, gbc_profileSheetNameLabel);

		JComboBox sheetComboBox = new JComboBox();
		GridBagConstraints gbc_sheetComboBox = new GridBagConstraints();
		gbc_sheetComboBox.fill = GridBagConstraints.HORIZONTAL;
		gbc_sheetComboBox.insets = new Insets(0, 0, 5, 5);
		gbc_sheetComboBox.gridx = 1;
		gbc_sheetComboBox.gridy = 3;

		panel.add(sheetComboBox, gbc_sheetComboBox);

		JLabel firstRowLabel = new JLabel("Pierwszy wiersz: ");
		GridBagConstraints gbc_firstRowLabel = new GridBagConstraints();
		gbc_firstRowLabel.anchor = GridBagConstraints.WEST;
		gbc_firstRowLabel.insets = new Insets(0, 0, 5, 5);
		gbc_firstRowLabel.gridx = 0;
		gbc_firstRowLabel.gridy = 4;
		panel.add(firstRowLabel, gbc_firstRowLabel);

		SpinnerModel firstRowSpinnerModel = new SpinnerNumberModel(7, 0, 999, 1);

		JSpinner firstRowSpinner = new JSpinner(firstRowSpinnerModel);
		GridBagConstraints gbc_firstRowSpinner = new GridBagConstraints();
		gbc_firstRowSpinner.fill = GridBagConstraints.BOTH;
		gbc_firstRowSpinner.insets = new Insets(0, 0, 5, 5);
		gbc_firstRowSpinner.gridx = 1;
		gbc_firstRowSpinner.gridy = 4;
		panel.add(firstRowSpinner, gbc_firstRowSpinner);

		JLabel lastRowLabel = new JLabel("Ostatni wiersz:");
		GridBagConstraints gbc_lastRowLabel = new GridBagConstraints();
		gbc_lastRowLabel.anchor = GridBagConstraints.WEST;
		gbc_lastRowLabel.insets = new Insets(0, 0, 0, 5);
		gbc_lastRowLabel.gridx = 0;
		gbc_lastRowLabel.gridy = 5;
		panel.add(lastRowLabel, gbc_lastRowLabel);

		SpinnerModel lastRowSpinnerModel = new SpinnerNumberModel(8, 0, 999, 1);
		JSpinner lastRowSpinner = new JSpinner(lastRowSpinnerModel);
		GridBagConstraints gbc_lastRowSpinner = new GridBagConstraints();
		gbc_lastRowSpinner.insets = new Insets(0, 0, 0, 5);
		gbc_lastRowSpinner.fill = GridBagConstraints.BOTH;
		gbc_lastRowSpinner.gridx = 1;
		gbc_lastRowSpinner.gridy = 5;
		panel.add(lastRowSpinner, gbc_lastRowSpinner);

		progressBar = new JProgressBar();
		progressBar.setBounds(93, 351, 429, 20);
		frmAutomatorOds.getContentPane().add(progressBar);

		messageLabel = new JLabel("");
		messageLabel.setBounds(93, 328, 429, 13);
		frmAutomatorOds.getContentPane().add(messageLabel);

		btnNewButton.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent e) {
				int firstRow = (int) (firstRowSpinner).getValue();
				int lastRow = (int) (lastRowSpinner).getValue();
				String profileSheetName = String.valueOf(sheetComboBox.getSelectedItem());

				boolean canProceed = true;
				if (lastRow < firstRow) {
					canProceed = false;
					System.out.println("Ostatni wiersz musi byæ poni¿ej pierwszego");
					showErrorDialog("Ostatni wiersz musi byæ poni¿ej pierwszego");
				}
				if (profileFilePath == null || contractFolderPath == null || profileSheetName == null) {
					canProceed = false;
					System.out.println("Uzupe³nij wszystkie pola!");
					showErrorDialog("Uzupe³nij wszystkie pola!");
				}

				if (profileSheetName.equals("zestawienie sprzeda¿y")) {
					canProceed = false;
					System.out.println("Wybierz inny arkusz ni¿ zestawienie sprzeda¿y.");
					showErrorDialog("Wybierz inny arkusz ni¿ zestawienie sprzeda¿y.");
				}

				if (canProceed) {
					Executors.newSingleThreadExecutor().execute(new Runnable() {
						@Override
						public void run() {
							backend.updateRows(new File(profileFilePath), contractFolderPath, profileSheetName,
									firstRow, lastRow);
						}
					});
				}

			}
		});

		// Wybór pliku zestawienia sprzeda¿y .ods
		openProfileFileButton.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent e) {
				JFileChooser fileChooser = new JFileChooser();
				fileChooser.setFileSelectionMode(JFileChooser.FILES_ONLY);
				int result = fileChooser.showOpenDialog(panel);
				if (result == JFileChooser.APPROVE_OPTION) {
					File selectedFile = fileChooser.getSelectedFile();
					profileFilePath = selectedFile.getAbsolutePath();
					profileFileTextField.setText(profileFilePath);

					Object[] profileSheets = backend.GetSheets(selectedFile.getAbsolutePath()).toArray();
					sheetComboBox.setModel(new DefaultComboBoxModel(profileSheets));
				}

			}
		});
	}
}
