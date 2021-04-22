package com.example.demo;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.text.MessageFormat;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.HashMap;
import java.util.HashSet;
import java.util.List;
import java.util.Map;
import java.util.Properties;
import java.util.stream.Collectors;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

import com.monitorjbl.xlsx.StreamingReader;

public class ExcelToWord {

//	String excelDir;// excel檔案路徑
	List<String> queryColArray;// 要抓取的欄位
//	File excelFolder; // excel資料夾
	String JCLNameLast;// 存放JCLName
//	String systemName;
	Map<Integer, String> KeyNameIndex = new HashMap<Integer, String>();

	ExcelToWord() {
		Properties pro = new Properties();
		// 設定檔位置
		String config = "C:\\Users\\si1153\\Documents\\workspace-spring-tool-suite-4-4.8.1.RELEASE\\ExcelToWord\\src\\main\\resources\\config.properties";
		try {
			// 讀取設定檔
			pro.load(new FileInputStream(config));
			// 讀取excel資料夾位置
			String excelDir = pro.getProperty("excelDir");
			// 讀取需要抓取的欄位名稱
			queryColArray = Arrays.asList(pro.getProperty("queryColArray").split(","));
			// 取資料夾
			File excelFolder = new File(excelDir);
			System.out.print(MessageFormat.format("excelDir:{0} 有 {1} 個Excel檔案", excelDir, excelFolder.list().length));
			//excel檔名
			String systemName = null;
			
			for (File file : excelFolder.listFiles()) {
				// 讀取excel檔案
				Workbook wb = getExcelFile(file.getPath());
				// 解析Excel to List
				List<Map<String, String>> excelInfoList = parseExcel(wb);
				// EX: excel檔名 帳務作業流程清單(BANK)_1090430 取 帳務作業流程清單(BANK)
				systemName = file.getName().split("_")[0];
				outPutToWork(excelInfoList);
			}
		} catch (FileNotFoundException e) {
			e.printStackTrace();
		} catch (IOException e) {
			e.printStackTrace();
		}
	}

	/**
	 * 讀取excel檔案
	 * 
	 * @param path excel檔案路徑
	 * @return excel內容
	 */
	public Workbook getExcelFile(String path) {
		Workbook wb = null;
		try {
			if (path == null) {
				return null;
			}
			String extString = path.substring(path.lastIndexOf(".")).toLowerCase();
			FileInputStream in = new FileInputStream(path);
			wb = StreamingReader.builder().rowCacheSize(100)// 存到記憶體行數，預設10行。
					.bufferSize(4096)// 讀取到記憶體的上限，預設1024
					.open(in);
		} catch (FileNotFoundException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}

//		old
//		File file = new File(path);
//		FileInputStream is;
//		try {
//			is = new FileInputStream(path);
//			if (".xls".equals(extString)) {
//				wb = new HSSFWorkbook(is);
//			} else if (".xlsx".equals(extString)) {
//				wb = new XSSFWorkbook(FileUtils.openInputStream(file));
//			}
//		} catch (FileNotFoundException e) {
//			e.printStackTrace();
//		} catch (IOException e) {
//			e.printStackTrace();
//		}
		return wb;
	}

	/**
	 * 解析Sheet
	 * 
	 * @param workbook Excel檔案
	 * @return 整個Sheet的資料
	 */
	public List<Map<String, String>> parseExcel(Workbook workbook) {
		// Sheet的資料
		List<Map<String, String>> excelDataList = new ArrayList<>();
		Sheet sheet;
		// 存放DNS欄位的欄位號
		int dnsIndex = 0;
		// 遍歷每一個sheet
		for (int sheetNum = 0; sheetNum < workbook.getNumberOfSheets(); sheetNum++) {
			sheet = workbook.getSheetAt(sheetNum);
			int rowNum = 1;
			if (sheet == null) {
				continue;
			}

			// 先取header
			for (Cell cell : sheet.getRow(rowNum)) {
				if (queryColArray.contains(cell.getStringCellValue())) {
					if (cell.getStringCellValue() == "DNS") {
						dnsIndex = cell.getColumnIndex();
					} else {
						KeyNameIndex.put(cell.getColumnIndex(), cell.getStringCellValue());
					}
				}
			}

			// 開始讀取sheet
			for (Row row : sheet) {
				if (rowNum == 1) {
					rowNum++;
					continue;
				}
				/*
				 * OLD code
				 * Row firstRow = sheet.getRow(firstRowNum); if (null == firstRow) {
				 * System.out.println("解析Excel失敗"); } int rowStart = sheetNum;// 起始去掉首欄 int
				 * rowEnd = sheet.getPhysicalNumberOfRows();OLD int dnsIndex = 0; old for (Cell
				 * cell : firstRow) { if (cell.getStringCellValue().equals("DSN")) { dnsIndex =
				 * cell.getColumnIndex(); } } for (int rowNum = rowStart; rowNum < rowEnd;
				 * rowNum++) { Row row = sheet.getRow(rowNum); if (null == row) { continue; } //
				 * 解析Row的資料 excelDataList.add(convertRowToData(row, firstRow, dnsIndex)); }-
				 */
				// 解析Row的資料
				excelDataList.add(convertRowToData(row,dnsIndex));
				rowNum++;
			}
		}
		return excelDataList;
	}

	/**
	 * 將資料重組並輸出Word
	 * 
	 * @param excelDataList 整理過的Excel檔案
	 */
	public void outPutToWork(List<Map<String, String>> excelDataList) {
		// 抓出不重複的JCL
		HashSet<String> jclKeys = new HashSet<>();
		excelDataList.forEach(cn -> {
			jclKeys.add(cn.get("JCL Name"));
		});
		// 將不重複的相同JCL_NAME的資料Group to List並輸出word
		jclKeys.forEach(classKey -> {
			List<Map<String, String>> toWordList = new ArrayList<>();
			toWordList = excelDataList.stream().filter(student -> student.get("Class") == classKey)
					.collect(Collectors.toList());
			System.out.println(excelDataList.stream().filter(student -> student.get("Class") == classKey)
					.collect(Collectors.toList()));
			// 輸出Word
			toWordList.clear();
		});

	}

	/**
	 * 解析ROW
	 * 
	 * @param row      資料行
	 * @param firstRow 標頭
	 * @param dnsIndex Dns的列數
	 * @return 整row的欄位
	 */
	public Map<String, String> convertRowToData(Row row,int dnsIndex) {
		Map<String, String> excelDateMap = new HashMap<String, String>();
		String firstRowName = null;
		for (Cell cell : row) {
			// 1.先抓現在第幾個Column
			int cellNum = cell.getColumnIndex();
			// 2.再去抓Header的欄位名稱
			firstRowName = KeyNameIndex.get(cell.getColumnIndex());

//			String firstRowName = firstRow.getCell(cell.getColumnIndex()).getStringCellValue();
			// 3.判斷是否為需要抓的欄位
			if (!queryColArray.contains(firstRowName)) {
				continue;
			}

			// "TWS AD Name,JCL Name,STEP Name,PROGRAM Name,DISP Status"
			// 抓到的欄位如果是JCL Name 會需要做空值塞值
			if (firstRowName.equals("JCL Name")) {
				if (cell.getStringCellValue().isEmpty() || cell.getStringCellValue() == null) {
					cell.setCellValue(JCLNameLast);
				} else {
					JCLNameLast = cell.getStringCellValue();
				}
			}
			// 如果是DISP Status，要抓DSN的值帶過來
			if (firstRowName.equals("DISP Status")) {
				switch (firstRowName) {
				case "MOD":
					firstRowName = "OUTPUT FILE";
					cell.setCellValue(row.getCell(dnsIndex).getStringCellValue());
					break;
				case "OLD":
					firstRowName = "INPUT FILE";
					cell.setCellValue(row.getCell(dnsIndex).getStringCellValue());
					break;
				case "SHR":
					firstRowName = "INPUT FILE";
					cell.setCellValue(row.getCell(dnsIndex).getStringCellValue());
					break;
				case "TLB645":
					break;
				default:
					break;
				}
			}

			excelDateMap.put(firstRowName, cell.getStringCellValue());
		}
		return excelDateMap;
	}

}
