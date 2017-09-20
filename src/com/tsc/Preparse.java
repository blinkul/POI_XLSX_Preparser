package com.tsc;

import java.io.BufferedReader;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.FileReader;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.HashMap;
import java.util.Iterator;
import java.util.List;

import org.apache.commons.lang3.StringUtils;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.xmlbeans.SystemProperties;

public class Preparse {

	private String projectPath = SystemProperties.getProperty("user.dir");
	private String mainDirName = "GeneratorEXCEL";
	private String inDirName = "1.Procesare";
	private String outDirName = "2.Finalizat";
	private String arcDirName = "3.Arhiva Original";
	private String errDirName = "4.Erori";
	private String logDirName = "5.Log";
	private String usageName = "Usage_Info.txt";

	private File mainDir = new File(projectPath + File.separator + mainDirName);
	private File inDir = new File(projectPath + File.separator + mainDirName + File.separator + inDirName);
	private File outDir = new File(projectPath + File.separator + mainDirName + File.separator + outDirName);
	private File arcDir = new File(projectPath + File.separator + mainDirName + File.separator + arcDirName);
	private File errDir = new File(projectPath + File.separator + mainDirName + File.separator + errDirName);
	private File logDir = new File(projectPath + File.separator + mainDirName + File.separator + logDirName);
	private File usage = new File(projectPath + File.separator + mainDirName + File.separator + usageName);

	private HashMap<String, String> usages = new HashMap<>();
	
	List<String> colH = new ArrayList<>();  
	List<String> colAH = new ArrayList<>();  
	List<String> colAL = new ArrayList<>();
	
	private File[] loadedFiles;
		
	XSSFWorkbook workBook;
	XSSFRow row;	

	// Metoda invocata in main
	public void run() {

		// Verifica si creaza folderele si fisierul usage
		boolean showLog = true;
		checkFile(mainDir, true, showLog);
		checkFile(inDir, true, showLog);
		checkFile(outDir, true, showLog);
		checkFile(arcDir, true, showLog);
		checkFile(errDir, true, showLog);
		checkFile(logDir, true, showLog);
		checkFile(usage, false, showLog);

		// Citeste din usage file si adauga fiecare KEY = VALUE in <HashMap
		// usage>
		readFromUsageFile(usage, usages);

		// Arata in consola ce valori sunt in usage
		readMap(usages);

		loadedFiles = loadExcelFiles(inDir, true);

		processFiles();

	}

	private void processFiles() {

		boolean status;
		
		try {
			
			// Fa o copie fisierelor originale
			org.apache.commons.io.FileUtils.copyDirectory(inDir, arcDir);

			// Proceseaza fiecare fisier din inDir
			for (File file : loadedFiles) {

				status = processOneFile(file);				
				
				if(status){
					System.out.println("File complete successful: " + file.getName());
					move(file, outDir);
					file.deleteOnExit();
				} else {
					System.out.println("File failed: " + file.getName());
					move(file, errDir);
				}
				
				clearColumns();
			}
		} catch (IOException e) {
			e.printStackTrace();
		} 

	}
	
	private boolean move(File file, File destDir) {
		boolean status = false;

		try {	
						
			org.apache.commons.io.FileUtils.moveFileToDirectory(file, destDir, true);
			
			status = true;
		} catch (IOException e) {
			e.printStackTrace();			
		}
		
		return status;
	}

	//TODO processOneFile
	/*
	 * [X] File is processing
	 * [X] Load columns from worksheet into Lists
	 * 		1. Load column index 7 (H)
	 * 		2. Load column index 33 (AH)
	 * 		3. Load column index 37 (AL)
	 * [X] Process columns (Lists)
	 *		1. Replace non numeric values with EXTRACTED number from col index 33
	 *		2. Remove leading 0s from column index 7
	 *		3. Concatenate col index 7 + " " + value of key from column index 37
	 *				-if key is not available, return N/A
	 * [4] Write columns back to worksheet
	 */
	private boolean processOneFile(File file) {
		boolean status = false; 	
		
		try(FileInputStream fis = new FileInputStream(file);			
			XSSFWorkbook workBook = new XSSFWorkbook(fis)
			) {
			
			System.out.println("\nProcessing file: " + file.getName());
			
			CreationHelper createHelper = workBook.getCreationHelper();		
			XSSFSheet spreadSheet = workBook.getSheetAt(0);
			
			getInformationFromSpreadSheet(spreadSheet);
			
			editArrays(colH, colAH, colAL);
			
			showRecordsFromCol(colH, "COLUMN H");
//			showRecordsFromCol(colAL, "COLUMN AL");
			
			writeInformationToSpreadSheet(spreadSheet, colH);				

			FileOutputStream fos = new FileOutputStream(file);
			
			workBook.write(fos);
			
						
			status = true;
		
		} catch (IOException e) {			
			e.printStackTrace();
		}
		return status;
	}
	
	//1. Get the information from spread sheet
	private void getInformationFromSpreadSheet(XSSFSheet spreadSheet) {
		Iterator<Row> rowIterator = spreadSheet.iterator();
		
		String cellString;
		
		while(rowIterator.hasNext()) {
			row = (XSSFRow) rowIterator.next();
			Iterator<Cell> cellIterator = row.cellIterator();
			while(cellIterator.hasNext()) {
				Cell cell = cellIterator.next();
				
				switch(cell.getCellType()){
				case Cell.CELL_TYPE_NUMERIC:
					
					if(cell.getColumnIndex() == 7) {
						cellString = String.valueOf(cell.getNumericCellValue());
						colH.add(cellString);
					}
					
					if(cell.getColumnIndex() == 33) {
						cellString = String.valueOf(cell.getNumericCellValue());
						colAH.add(cellString);
					}
					
					if(cell.getColumnIndex() == 37) {
						cellString = String.valueOf(cell.getNumericCellValue());
						colAL.add(cellString);
					}	
					
					break;
				case Cell.CELL_TYPE_STRING:
					
					if(cell.getColumnIndex() == 7) {
						cellString = cell.getStringCellValue();
						colH.add(cellString);
					}
					
					if(cell.getColumnIndex() == 33) {
						cellString = cell.getStringCellValue();
						colAH.add(cellString);
					}
					
					if(cell.getColumnIndex() == 37) {
						cellString = cell.getStringCellValue();
						colAL.add(cellString);
					}
					
					break;
					
				}				
			}
		}		
	}
	
	//2. Edit the information from spread sheet
    private void editArrays(List<String> column1, List<String> column2, List<String> column3) {
		
		String newCellValue;	
		
		String value1;
		String value2;
				
		for(int i = 1; i < column1.size(); i++) {
				
			//CREATE VALUE 1
			//RPLACE NON NUMERIC FROM COL H AND REMOVE LEADING 0
			if(StringUtils.isNumeric(column1.get(i)) == false){
				value1 = removeLeadZeros(extractCostCenter(column2.get(i)));
			} else {
				value1 = removeLeadZeros(column1.get(i));
			}
			
			//CREATE VALUE 2
			if(usages.containsKey(column3.get(i))){
				value2 = usages.get(column3.get(i));
			} else {
				value2 = "N/A";
			}		
			
			//CREATE THE NEW VALUE
			newCellValue = value1 + " " +value2;			
			
			//SET THE NEW VALUE
			column1.set(i, newCellValue);
		}
	}
	
	private String removeLeadZeros(String string) {
		
		boolean isLeading = true;
		
		char[] charArr = string.toCharArray();
		StringBuilder sb = new StringBuilder();
		
		for(char c : charArr) {
			//00102
			if(isLeading) {
				if(c != '0') {
					isLeading = false;
					sb.append(c);
				}
			} else {
				sb.append(c);
			}			
		}
		
		return sb.toString();
		
	}
		
	private String extractCostCenter(String string) {
				
		char[] charArr = string.toCharArray();
		StringBuffer sb = new StringBuffer();
		
		for(char c : charArr) {
			if(c >= '0' && c <= '9') {
				sb.append(c);
			}
		}
		
		return sb.toString();	
	}
		
	private void writeInformationToSpreadSheet(XSSFSheet spreadSheet, List<String> list) {
		
		
		
		for(int i = 1; i < list.size() - 1; i++){
			Row row = spreadSheet.getRow(i);
//			Cell cell = row.getCell(7);
			Cell cell = row.createCell(7);
			
//			cell.setCellType(Cell.CELL_TYPE_STRING);
			cell.setCellValue(list.get(i));
			
		}
		
		
//		for(int i = 1; i <= totalRows; i++) {
//		    Row row = sheet.getRow(i);
//		    Cell cell = row.getCell(2);
//		    if (cell == null) {
//		        cell = row.createCell(2);
//		    }
//		    cell.setCellType(Cell.CELL_TYPE_STRING);
//		    cell.setCellValue("some value");
//		}
		
	}
	
	private void clearColumns(){
		colH.clear();
		colAH.clear();
		colAL.clear();
	}

	// Incarca fisierele excel din filderul inDir
	private File[] loadExcelFiles(File inDir, boolean showLog) {

		File[] files = inDir.listFiles();

		if (showLog) {
			System.out.println("\nFisiere disponibile in directorul \"" + inDirName + "\":");
			for (File file : files) {
				System.out.println(file.getName());
			}
		}

		return files;
	}

	// Citeste din usage file si adauga fiecare KEY = VALUE in <HashMap usage>
	public boolean readFromUsageFile(File usage, HashMap<String, String> usages) {
		boolean status = false;
		String line;

		String[] lineArr;

		try (BufferedReader br = new BufferedReader(new FileReader(usage))) {

			while ((line = br.readLine()) != null) {

				if (line.contains("=")) {

					lineArr = line.split("=");

					usages.put(lineArr[0].trim(), lineArr[1].trim());
				} else {
					System.out.println("\nIncorrect key one line:" + line + "\nMissing \"=\" sign");
				}
			}

		} catch (IOException ex) {
			ex.getMessage();
			ex.printStackTrace();
		}

		status = true;
		return status;
	}

	// Arata in consola ce valori sunt in usage
	public void readMap(HashMap<String, String> usages) {

		System.out.println();
		System.out.println(Arrays.asList(usages));

	}

	private void checkFile(File file, boolean isDir, boolean showLog) {
		try {

			if (isDir) {
				if ((file.exists() && file.isDirectory()) == false) {
					file.mkdir();
					if (showLog) {
						System.out.println("INFO: " + file.getName() + " was created: " + file.getAbsolutePath());
					}
				} else {
					if (showLog) {
						System.out.println("INFO: " + file.getName() + " exists: " + file.getAbsolutePath());
					}
				}

			} else {
				if ((file.exists() && file.isFile()) == false) {
					file.createNewFile();
					if (showLog) {
						System.out.println("INFO: " + file.getName() + " was created: " + file.getAbsolutePath());
					}
				} else {
					if (showLog) {
						System.out.println("INFO: " + file.getName() + " exists: " + file.getAbsolutePath());
					}
				}
			}
		} catch (IOException e) {
			e.printStackTrace();
		}
	}
	
	private void showRecordsFromCol(List<String> column, String columnName) {
		System.out.println("RECORDS FROM COL " + columnName + "***************");
		for(String s : column) {
			System.out.println(s);
		}
		System.out.println("*************************");
	}
}
