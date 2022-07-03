
/* 
This Application is built to help people generate millions of PDFs from a HTML template by populating data from 
XLS or XLSX file. This was preprared with the purpose of generating multiple copies of a letter/document that needs 
to be peronalised for the receiver with their name, contact details and details associated with them. A typical example 
would be generating letters with common intent for individual customers personalized for each of them such as information 
about a new product or a new service or intent to extend the service terms and period.

Copyright (C) 2022  Ghanithan Subramani
Git: https://github.com/ghanithan

This program is free software: you can redistribute it and/or modify
it under the terms of the GNU Affero General Public License as
published by the Free Software Foundation, either version 3 of the
License, or (at your option) any later version.

This program is distributed in the hope that it will be useful,
but WITHOUT ANY WARRANTY; without even the implied warranty of
MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
GNU Affero General Public License for more details.

You should have received a copy of the GNU Affero General Public License
along with this program.  If not, see https://www.gnu.org/licenses/agpl-3.0.en.html. 

This Application uses iText 7 core library along with pdfHTML plugin to convert HTML to PDF.
The iText 7 libraries are used under open source GNU AGPL license and I have made the source code of the application 
open source honouring the GNU AGPL license terms of the iText 7 library. The libraries are included in the build 
using the Maven repository through the Gradle build. You should have received a copy of the GNU Affero General Public License
along with this program.  If not, see https://itextpdf.com/en/how-buy/legal/agpl-gnu-affero-general-public-license.

Please visit the library's page https://itextpdf.com/en/products/itext-7/convert-html-css-to-pdf-pdfhtml for more details.

*/


package PDFReplicator;

import java.awt.BorderLayout;
import java.awt.EventQueue;
import java.awt.Toolkit;

import javax.swing.JFrame;
import javax.swing.JPanel;
import javax.swing.JProgressBar;
import javax.swing.border.EmptyBorder;
import javax.swing.event.ChangeEvent;
import javax.swing.event.ChangeListener;
import javax.swing.plaf.ProgressBarUI;

import org.apache.commons.io.FileUtils;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.format.CellDateFormatter;
import org.apache.poi.hssf.usermodel.HSSFFormulaEvaluator;
import org.apache.poi.xssf.usermodel.XSSFFormulaEvaluator;

import com.itextpdf.html2pdf.ConverterProperties;
import com.itextpdf.html2pdf.HtmlConverter;
import com.itextpdf.html2pdf.resolver.font.DefaultFontProvider;


import javax.swing.JTextField;
import javax.swing.SwingUtilities;
import javax.swing.Timer;
import javax.swing.JButton;
import javax.swing.JFileChooser;
import java.awt.event.ActionListener;
import java.io.ByteArrayInputStream;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.math.BigDecimal;
import java.nio.charset.Charset;
import java.nio.file.Paths;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;
import java.util.Date;
import java.util.Scanner;
import java.awt.event.ActionEvent;
import javax.swing.JLabel;
import javax.swing.JOptionPane;

public class App extends JFrame {

	private JPanel contentPane; // Panel to hold the UI elements in the window
	private JTextField textField; // The field to capture path of Excel file
	private JTextField textField_1; // THe field to capture path of HTML template
	private JProgressBar progress; // Progress bar to display the progress
	private int rows;
	private int columns;
	private int rowCount;
	private JButton convertpdf; // Button to initiate the process
	private Task task; // To initiate the PDF generation process in sepereate thread from the UI
	private Timer timer; // Timer to interrupt the pdf-conversion task and update the process bar
	
	List<String> headerArray = new ArrayList<String>(); // List to store the fields in header of the excel to iterate over
	/**
	 * Launch the application.
	 */
	public static void main(String[] args) {
		EventQueue.invokeLater(new Runnable() {
			public void run() {
				try {
					App frame = new App();
					
				
					frame.setVisible(true);
				} catch (Exception e) {
					e.printStackTrace();
				}
			}
		});
	}

	/**
	 * Create the frame.
	 */
	public App() {

        // Set the general properties of the window along with dimensions
		setTitle("PDFReplicator");
		setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
		setBounds(100, 100, 450, 300);
		contentPane = new JPanel();
		contentPane.setBorder(new EmptyBorder(5, 5, 5, 5));
		setContentPane(contentPane);
		contentPane.setLayout(null);
		
        // Create an input field to capture the path to excel file along with file locator
		textField = new JTextField();
		textField.setBounds(10, 39, 291, 20);
		contentPane.add(textField);
		textField.setColumns(10);
		JButton excelupload = new JButton("Open Excel");
		excelupload.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent e) {
				final JFileChooser chooser = new JFileChooser(""); // can set the current path to "./" to display the project folder by default
			
			int	response = chooser.showOpenDialog( App .this);
				if(response == JFileChooser.APPROVE_OPTION) {
					String textField1 = chooser.getSelectedFile().toString();
					textField.setText(textField1);
					
					System.out.println(textField1);
				}
				
			}
		});
		excelupload.setBounds(315, 38, 109, 23);
		contentPane.add(excelupload);
		JLabel lblNewLabel = new JLabel("SELECT THE EXCEL FILE");
		lblNewLabel.setBounds(20, 11, 199, 14);
		contentPane.add(lblNewLabel);
		
        // Create an input field to capture the path to HTML Template file along with file locator
		JLabel lblNewLabel_1 = new JLabel("SELECT THE HTML FILE");
		lblNewLabel_1.setBounds(20, 73, 166, 14);
		contentPane.add(lblNewLabel_1);
		textField_1 = new JTextField();
		textField_1.setBounds(10, 99, 291, 20);
		contentPane.add(textField_1);
		textField_1.setColumns(10);
		
		JButton htmlupload = new JButton("Open Html");
		htmlupload.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent e) {
				final JFileChooser chooser1 = new JFileChooser("");// can set the current path to "./" to display the project folder by default
				int	response1 = chooser1.showOpenDialog( App .this);
				if(response1 == JFileChooser.APPROVE_OPTION) {
					String textField2 = chooser1.getSelectedFile().toString();
					textField_1.setText(textField2);
					
					System.out.println(textField2);
				}
				
				
				
				
				
			}
		});
		htmlupload.setBounds(315, 98, 109, 23);
		contentPane.add(htmlupload);
		
    
        // Progress bar component
		progress = new JProgressBar(0, 100);
		progress.setBounds(10,220,415,30);         
		progress.setValue(0);    
		progress.setStringPainted(true);
		progress.setVisible(true);
		contentPane.add(progress);
		
		
	    // Button to initiate the task
		convertpdf = new JButton("Convert To PDF ");
		convertpdf.addActionListener(new ButtonListener());
		convertpdf.setBounds(118, 166, 154, 23);
		contentPane.add(convertpdf);
		
		timer = new Timer( 10 , new TimerListener() );
	}
	

    // Timer to check the progress of the task
class TimerListener implements ActionListener{
	public void actionPerformed(ActionEvent evt) {
	
		if(task.isAlive()) {
			progress.setValue(rowCount);
		}else {
			
			timer.stop();
			Toolkit.getDefaultToolkit().beep();
			JOptionPane.showMessageDialog( App .this, "PDFs are Ready in Output Folder.", "PDF Generated",1);
			progress.setValue(0);
			convertpdf.setEnabled(true);
		}
	}
}
	
// The Event listener binded wth the JButton convertpdf which will start a thread to generate PDFs
class ButtonListener implements ActionListener{
	public void actionPerformed(ActionEvent evt) {
			convertpdf.setEnabled(false);
			task = new Task();
			task.start(); 
			timer.start();
			
		}
			
	}

    // Primary task in which the PDF is generated using the HTML template and the data from excel file
	private class Task extends Thread{
		public Task() {
			
		}
		
		public void run() {
			
			rowCount = 0;
			System.out.println("Excel File Selected "+ textField.getText());
			System.out.println("HTML Template File Selected "+ textField_1.getText());
			
			
			String newHtmlFileextension = "";
			String newPdfFileextension = "";
			
			
			File file = new File(textField.getText());
			
			// Create an object of FileInputStream class to read excel file
			FileInputStream inputStream;
			try {
				inputStream = new FileInputStream(file);
                Double dnum; // To hold mobile number
                Date local_date; // To hold cell values in date format
                String dateFmt ; // To hold format of date
                String dateStrValue; // To hold string value of date
                Workbook AddCatalog = null;

                // Find the file extension by splitting file name in substring and getting only
                // extension name
                String fileExtensionName = textField.getText().substring(textField.getText().indexOf("."));

                // Check condition if the file is a .xls file or .xlsx file
                if (fileExtensionName.equals(".xls")) {
                    // If it is .xls file then create object of HSSFWorkbook class
                    AddCatalog = new HSSFWorkbook(inputStream);
                    HSSFFormulaEvaluator.evaluateAllFormulaCells(AddCatalog);

                    
                } else if (fileExtensionName.equals(".xlsx")) {
                    // If it is .xlsx file then create object of XSSFWorkbook class
                    AddCatalog = new XSSFWorkbook(inputStream);
                    XSSFFormulaEvaluator.evaluateAllFormulaCells(AddCatalog);

                }

                // Read sheet inside the workbook by its name
                Sheet AddCatalogSheet = AddCatalog.getSheetAt(0);
                
                int firstRowNum = AddCatalogSheet.getFirstRowNum();
                System.out.println("First Row is " + firstRowNum);
                
                int lastRowNum = AddCatalogSheet.getLastRowNum();
                System.out.println("Last Row is " + lastRowNum);
                //Read first row
                Row header = AddCatalogSheet.getRow(firstRowNum);
                
                int firstCellNum = header.getFirstCellNum();
                System.out.println("First Column is " + firstCellNum);
                int lastCellNum = header.getLastCellNum(); //returns last cell number PLUS ONE  
                System.out.println("Last Column is " + lastCellNum);
                
                rows = lastRowNum - firstRowNum + 1;
                
                columns = lastCellNum - firstCellNum;	
                
                System.out.println("number of rows is " + rows);
                
                progress.setMaximum(lastRowNum);
                            
                Iterator<Cell> headerIterator = header.cellIterator();
                while(headerIterator.hasNext()) {
                    Cell headerContent = headerIterator.next();
                    
                    headerArray.add("$"+headerContent.toString());
                }
                
                // Find number of rows in excel file
                //  int rowcount = AddCatalogSheet.getLastRowNum() - AddCatalogSheet.getFirstRowNum();
                //  System.out.println("Total row number: " + rowcount);
                
                //create directroies
                
                FileUtils.deleteQuietly(new File("output"));
                
                File htmldir = new File("output" +"/" +"HTML");
                htmldir.mkdirs();
                File pdfdir = new File("output" +"/" +"PDF");
                pdfdir.mkdirs();
                
                
                Iterator<Row> rowIterator = AddCatalogSheet.iterator();
                if(rowIterator.hasNext()) { // to skip header
                    rowIterator.next();
                }
                
                // Loop over the rows in the excel sheet 
                while(rowIterator.hasNext()) {
                    Row rowContent = rowIterator.next();
                    
                    List<String> rowArray = new ArrayList<String>();
                    File htmlTemplateFile = new File(textField_1.getText());
                    String htmlString = FileUtils.readFileToString(htmlTemplateFile,Charset.defaultCharset());
                    
                    
                    
                    Iterator<Cell> cellIterator = rowContent.cellIterator();
                    int count = 0;
                    while((cellIterator.hasNext()) && (count < columns)) {
                        Cell cellContent = cellIterator.next();
                        int i = cellContent.getColumnIndex();
                        //System.out.println(i);
                        if(cellContent.getCellType() == CellType.NUMERIC){
                            if (DateUtil.isCellDateFormatted(cellContent)) {
                                //   System.out.println("The cell contains a date value: " + cellContent.getDateCellValue());
                                   local_date = cellContent.getDateCellValue();
                                   dateFmt = cellContent.getCellStyle().getDataFormatString();
                                   dateStrValue = new CellDateFormatter(dateFmt).format(local_date); 
                                   cellContent.setBlank();
                                   cellContent.setCellValue(dateStrValue);
                                   
                                   
                               }else{
                                    dnum = cellContent.getNumericCellValue();
                                    BigDecimal bd = new BigDecimal(dnum);
                                    //System.out.println(bd.toPlainString());
                                    cellContent.setBlank();
                                    cellContent.setCellValue(bd.toPlainString());
                                    //System.out.println(cellContent.getStringCellValue());
                               }
                        }
                        rowArray.add(cellContent.toString());	
                        htmlString = htmlString.replace(headerArray.get(i), rowArray.get(i));
                        count = count +1;
                    }
                    String newFileextension = rowArray.get(1) + "/" + rowArray.get(2) + "/" + rowArray.get(2) + "_" + rowArray.get(0);
                    
                    
                    newHtmlFileextension = htmldir + "/" + newFileextension;
                    newPdfFileextension = pdfdir + "/" + newFileextension;
                    
                // commented the below lines related to saving html files before converting them to PDFs. It was not the best idea to store
                // the intermiate HTML file and consume more space. It was good for debugging though. So preserving the code.
                //	File createHtmldir = new File(htmldir + "/" + rowArray.get(1) + "/" + rowArray.get(2) );
                //	createHtmldir.mkdirs();
                    File createPdfdir = new File(pdfdir + "/" + rowArray.get(1) + "/" + rowArray.get(2) );
                    createPdfdir.mkdirs();
                    
                    
                //	File newHtmlFile = new File(newHtmlFileextension + ".html");
                    
                //	FileUtils.writeStringToFile(newHtmlFile, htmlString,Charset.forName("UTF-16"));
                    //System.out.println("newHtmlFile -->"+newHtmlFile);
                    
                    
                    //Code for converting html to pdf
                    
                    //System.out.println("textField_1.getText() --"+textField_1.getText());
                //	File htmlSource = new File(newHtmlFileextension +".html");
                    
                    //System.out.println("HTML BUILD PATH-->"+htmlSource);
                    
                    File pdfDest = new File(newPdfFileextension +".pdf");
                    
                    
                    //System.out.println("PDF BUILD PATH-->"+pdfDest);
                    // pdfHTML  code
                    
                    // Directly converting HTML string to PDF usign iTEXT 7 HTML to PDF library - pdfHTML plugin 
                    // On behalf of all the users of this opensource Application, I would like to appreciate iTEXT for sharing
                    // the library as open source library with AGPL license and to comply with their community requirement, 
                    // I am making the code as opensource under AGPL license

                    InputStream inputStream1 = new ByteArrayInputStream(htmlString.getBytes("UTF-16"));
                    ConverterProperties properties = new ConverterProperties();
                    properties.setFontProvider(new DefaultFontProvider(true, true, true));
                    HtmlConverter.convertToPdf( inputStream1, 
                    new FileOutputStream(pdfDest), properties);
                //  htmlSource.delete(); // was used to delete the PDF that was created as intermediate file
                
                    
                    rowCount = rowCount +1;
                    
                    
    //		        System.out.println(rowCount );
                    
                    try {
                        sleep(6); // putting the thread to sleep so that UI can be updated
                    }catch(InterruptedException e ) {
                        e.printStackTrace();
                    }
                    
                }
			} catch (IOException e1) {
				
				e1.printStackTrace();
			}
			String bulid_success = StringUtils.substringBeforeLast(newHtmlFileextension,"\\");
			System.out.println("PDF GENERATION SUCCESSFULL--"+ "PDFs Ready in Output Folder.");
		
			
			
			try {
                // I had written scripts in to create archive the pdf created in this process for easy sharing.
                // Would also recommend to use pdfCPU written in GO. Binaries for all target machines are avialable
                // https://github.com/pdfcpu/pdfcpu/releases
	            String[] command = {"cmd.exe", "/C", "Start",Paths.get("").toAbsolutePath().toString()+"\\" + "zip.bat"};
	            Process p =  Runtime.getRuntime().exec(command);           
	        } catch (IOException ex) {
	        	ex.printStackTrace();
	        }
			
			
		}
		
	}
}
