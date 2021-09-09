package main;

import java.awt.EventQueue;
import java.awt.GridBagConstraints;
import java.awt.GridBagLayout;
import java.awt.Insets;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.io.File;

import javax.swing.DefaultComboBoxModel;
import javax.swing.JButton;
import javax.swing.JComboBox;
import javax.swing.JFileChooser;
import javax.swing.JFrame;
import javax.swing.JLabel;
import javax.swing.JPanel;
import javax.swing.JSpinner;
import javax.swing.JTextField;

public class Main {

	private JFrame frmAutomatorOds;
	private Backend backend;
	private JTextField contractFolderTextField;

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
		backend = new Backend();
		
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
		panel.setBounds(93, 156, 429, 172);
		frmAutomatorOds.getContentPane().add(panel);
		GridBagLayout gbl_panel = new GridBagLayout();
		gbl_panel.columnWidths = new int[]{162, 170, 70, 0};
		gbl_panel.rowHeights = new int[]{23, 20, 0, 0, 0, 0, 0};
		gbl_panel.columnWeights = new double[]{1.0, 1.0, 0.0, Double.MIN_VALUE};
		gbl_panel.rowWeights = new double[]{0.0, 0.0, 0.0, 0.0, 0.0, 0.0, Double.MIN_VALUE};
		panel.setLayout(gbl_panel);
		Object[] profileSheets = backend.GetSheets("res/Kwiecieñ 2020.ods").toArray();
		
		JLabel contractFolderLabel = new JLabel("Folder z umowami: ");
		GridBagConstraints gbc_contractFolderLabel = new GridBagConstraints();
		gbc_contractFolderLabel.insets = new Insets(0, 0, 5, 5);
		gbc_contractFolderLabel.anchor = GridBagConstraints.WEST;
		gbc_contractFolderLabel.gridx = 0;
		gbc_contractFolderLabel.gridy = 1;
		panel.add(contractFolderLabel, gbc_contractFolderLabel);
		
		contractFolderTextField = new JTextField();
		GridBagConstraints gbc_contractFolderTextField = new GridBagConstraints();
		gbc_contractFolderTextField.insets = new Insets(0, 0, 5, 5);
		gbc_contractFolderTextField.fill = GridBagConstraints.BOTH;
		gbc_contractFolderTextField.gridx = 1;
		gbc_contractFolderTextField.gridy = 1;
		panel.add(contractFolderTextField, gbc_contractFolderTextField);
		contractFolderTextField.setColumns(10);
		
		JButton openContractFolderButton = new JButton("Otw\u00F3rz");
		openContractFolderButton.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent e) {
				JFileChooser fileChooser = new JFileChooser();
				fileChooser.setFileSelectionMode(JFileChooser.DIRECTORIES_ONLY);
				fileChooser.setCurrentDirectory(new File("res/"));
				int result = fileChooser.showOpenDialog(panel);
				if (result == JFileChooser.APPROVE_OPTION) {
				    File selectedFile = fileChooser.getSelectedFile();
				    contractFolderTextField.setText(selectedFile.getAbsolutePath()); 
				    System.out.println("Selected file: " + selectedFile.getAbsolutePath());
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
		sheetComboBox.setModel(new DefaultComboBoxModel(profileSheets));
		panel.add(sheetComboBox, gbc_sheetComboBox);
		

		
		JLabel firstRowLabel = new JLabel("Pierwszy wiersz: ");
		GridBagConstraints gbc_firstRowLabel = new GridBagConstraints();
		gbc_firstRowLabel.anchor = GridBagConstraints.WEST;
		gbc_firstRowLabel.insets = new Insets(0, 0, 5, 5);
		gbc_firstRowLabel.gridx = 0;
		gbc_firstRowLabel.gridy = 4;
		panel.add(firstRowLabel, gbc_firstRowLabel);
		
		JSpinner firstRowSpinner = new JSpinner();
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
		
		JSpinner lastRowSpinner = new JSpinner();
		GridBagConstraints gbc_lastRowSpinner = new GridBagConstraints();
		gbc_lastRowSpinner.insets = new Insets(0, 0, 0, 5);
		gbc_lastRowSpinner.fill = GridBagConstraints.BOTH;
		gbc_lastRowSpinner.gridx = 1;
		gbc_lastRowSpinner.gridy = 5;
		panel.add(lastRowSpinner, gbc_lastRowSpinner);
		
		btnNewButton.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent e) {
				int firstRow = (int) (firstRowSpinner).getValue();
				int lastRow = (int) (lastRowSpinner).getValue();
				String profileSheetName = String.valueOf(sheetComboBox.getSelectedItem());
				
				System.out.println(firstRow + ", " + lastRow);
				System.out.println(profileSheetName);
				


			}
		});
	}
}
