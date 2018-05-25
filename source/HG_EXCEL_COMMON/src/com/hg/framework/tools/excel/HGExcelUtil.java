/**
 * Copyright (c) 2014-2017 HG R&D Team.  All Rights Reserved.
 * This software is published under the HG Solution Team.
 * License version 2.0, a copy of which has been included with this
 * distribution in the LICENSE.txt file.
 *
 * @File name:  HGExcelUtil.java * @Create on:  2018年5月8日
 * @Author   :  
 *
 * @ChangeList
 * ---------------------------------------------------
 * NO      Date               Editor             ChangeReasons
 * 1       2018年5月8日            kuangyj            Create 
 *
 */
package com.hg.framework.tools.excel;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;

import org.apache.log4j.Logger;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.hg.framework.tools.excel.vo.HGExcelCellVO;
import com.hg.framework.tools.excel.vo.HGExcelRowVO;
import com.hg.framework.tools.excel.vo.HGExcelSheetVO;

/**
 * 处理EXCEL的基础类
 * 
 * @author kuangyj
 *
 */
public class HGExcelUtil {
	private static Logger logger = Logger.getLogger(HGExcelUtil.class);
	private static HGExcelUtil excelUtil = new HGExcelUtil();
	private HGExcelUtil() {		
	}
	
	/**
	 * DESC: get the instance of HVExcelUtil  
	 * 
	 * @return
	 */
	public static HGExcelUtil getInstance() {
		if (excelUtil == null) {
			excelUtil = new HGExcelUtil();
		}
		return excelUtil;	
	}
	
	/**
	 * DESC: 解析EXCEL内容到对象中  <br>
	 *       通过EXCEL文件后缀名自动识别OFFICE 2003 或者OFFICE 2005+版本 <br>
	 *       优先对excludeSheet的内容生效【exclude内容不为空时则include内容不做处理】 <br>
	 *       excludeSheet --- 排除的sheet, 在该内容中的Sheet不做解析 <br> 
	 *       includeSheet --- 包含的sheet, 仅解析该内容中的Sheet内容  <br>
	 * 
	 * @param fileName        待解析的EXCEL文件
	 * @param includeSheet    要处理的SHEET
	 * @param excludeSheet    不做处理的SHEET
	 * @Param matchType    -- 匹配模式<br>
	 * 						1 ： 等于<br>
	 * 						2： 包含<br>
	 * 						3：Start With<br>
	 * 						4 ： End With<br>
	 * @return
	 */
	public HashMap<String, HGExcelSheetVO> parseExcelData(String fileName, List<String> includeSheet, List<String> excludeSheet, int matchType) {
		HashMap<String, HGExcelSheetVO> sheetMap = new HashMap<String, HGExcelSheetVO>();
		if (fileName != null && fileName.endsWith(".xls")) {
			//office 2003 
			sheetMap = this.parseExcelData4XLS(fileName, includeSheet, excludeSheet, matchType);
		} else if (fileName != null && fileName.endsWith(".xlsx")){
			//office 2005+
			sheetMap = this.parseExcelData4XLSX(fileName, includeSheet, excludeSheet, matchType);
		}
 		return sheetMap;
	}
	
	/**
	 * DESC: 修改指定的EXCEL单元格内容   <br>
	 *       通过文件后缀名自动识别OFFICE 2003或者OFFICE 2005+版本 <br>
	 * 	     data -- key ： SHEET内容   <br>
	 *       VALUE： 待填写的CELL数据内容,HVExcelCellVO对象  <br> 
	 * 
	 * @param fileName
	 * @param dataMap
	 * @return
	 */
	public boolean editExcelData(String fileName, HashMap<String, List<HGExcelCellVO>> dataMap) {
		boolean result = false;
		if (fileName != null && fileName.endsWith(".xls")) {
			//office 2003 
			result = this.editExcelData4XLS(fileName, dataMap);
		} else if (fileName != null && fileName.endsWith(".xlsx")){
			//office 2005+
			result = this.editExcelData4XLSX(fileName, dataMap);
		}
		return result;
	}
	
	/**
	 * DESC: 复制文件 
	 * 
	 * @param sourceFileName  -- 源文件
	 * @param targetFileName  -- 目标文件
	 * @param overFlag  -- 是否覆盖
	 */
	public boolean copyFile(String sourceFileName, String targetFileName, boolean overFlag) {
		boolean result = false;
		File sourceFile = new File(sourceFileName);
		if (!sourceFile.exists()) {
			logger.error("sourceFileName [" + sourceFileName + "] Not Found!");
			return false;
		} else if (!sourceFile.isFile()){
			logger.error("sourceFileName [" + sourceFileName + "] isn't file!");
			return false;
		} 
		File targetFile = new File(targetFileName);
		if (targetFile.exists()) {
			if (overFlag) {
				new File(targetFileName).delete();
			}
		} else {
			if (!targetFile.getParentFile().exists()) {
				if (!targetFile.getParentFile().mkdirs()) {
					logger.error("targetFileName [" + targetFileName + "] can;t make dir!");
					return false;
				}
			}
		}
		//copy the file
		int bytes = 0;
		InputStream in = null;
		OutputStream out = null;
		try {
			in = new FileInputStream(sourceFile);
			out = new FileOutputStream(targetFile);
			byte[] buff = new byte[1024];
			while ((bytes = in.read(buff)) != -1) {
				out.write(buff, 0, bytes);
			}
			result = true;
		} catch (FileNotFoundException e) {
			logger.error(e);
			e.printStackTrace();
		} catch (IOException e) {
			logger.error(e);
			e.printStackTrace();
		} finally {
			if (out != null) {
				try {
					out.close();
				} catch (IOException e) {}
			} 
			if (in != null) {
				try {
					in.close();
				} catch (IOException e) {}
			}
		}
		return result;
	}
	
	//===================== private method ========================
	/**
	 * DESC: 解析OFFICE 2003版本的EXCEL内容到对象中  
	 *       优先对excludeSheet的内容生效【exclude内容不为空时则include内容不做处理】
	 *       excludeSheet --- 排除的sheet, 在该内容中的Sheet不做解析
	 *       includeSheet --- 包含的sheet, 仅解析该内容中的Sheet内容
	 * 
	 * @param fileName
	 * @param includeSheet
	 * @param excludeSheet
	 * @param matchType  -- 匹配模式<br>
	 * 						1 ： 等于<br>
	 * 						2： 包含<br>
	 * 						3：Start With<br>
	 * 						4 ： End With<br>
	 * @return
	 */
	private HashMap<String, HGExcelSheetVO> parseExcelData4XLS(String fileName, List<String> includeSheet, List<String> excludeSheet, int matchType) {
		HashMap<String, HGExcelSheetVO> sheetMap = new HashMap<String, HGExcelSheetVO>();
		//read xls
		HSSFWorkbook workbook = null;
		File excelFile = new File(fileName);
		InputStream is = null;
		try {
			is = new FileInputStream(excelFile);
			workbook = new HSSFWorkbook(is);
			for (int sheetIndexOf = 0; sheetIndexOf < workbook.getNumberOfSheets(); sheetIndexOf++) {
				HSSFSheet sheet = workbook.getSheetAt(sheetIndexOf);
				if (sheet != null) {
					String sheetName = sheet.getSheetName();
					HGExcelSheetVO sheetVO = new HGExcelSheetVO();
					sheetVO.setSheetName(sheetName);
					boolean flag = false;
					if (excludeSheet != null && excludeSheet.size() > 0) {
						if (!this.checkStrExist(sheetName, excludeSheet, matchType)) {
							flag = true;
						}
					} else if (includeSheet != null && includeSheet.size() > 0) {
						if (this.checkStrExist(sheetName, includeSheet, matchType)) {
							flag = true;
						}
					} else {
						flag = true;
					}
					//parse the row 
					if (flag) {
						int rowIndexOf = sheet.getFirstRowNum();
						int rowCounts = sheet.getLastRowNum();
						List<HGExcelRowVO> rowList = new ArrayList<HGExcelRowVO>();
						for (; rowIndexOf < rowCounts; rowIndexOf++) {
							HSSFRow row = sheet.getRow(rowIndexOf);
							if (row != null) {
								HGExcelRowVO rowVO = new HGExcelRowVO();
								rowVO.setRowIndex(rowIndexOf);
								List<HGExcelCellVO> cellList = new ArrayList<HGExcelCellVO>();
								int cellIndexOf = row.getFirstCellNum();
								int cellCounts = row.getLastCellNum();
								for (; cellIndexOf < cellCounts; cellIndexOf++) {
									HSSFCell cell = row.getCell(cellIndexOf);
									if (cell != null) {
										HGExcelCellVO cellVO = new HGExcelCellVO();
										cellVO.setCellIndex(cellIndexOf);
										cellVO.setRowIndex(rowIndexOf);
										cellVO.setCellType(cell.getCellType());
										switch (cell.getCellType()) {
											case Cell.CELL_TYPE_BLANK : cellVO.setCellValue("");break;
											case Cell.CELL_TYPE_BOOLEAN : cellVO.setCellValue(String.valueOf(cell.getBooleanCellValue()));break;
											case Cell.CELL_TYPE_NUMERIC : cellVO.setCellValue(String.valueOf(cell.getNumericCellValue()).trim());break;
											case Cell.CELL_TYPE_STRING : cellVO.setCellValue(cell.getStringCellValue().trim());break;
										}
										cellList.add(cellVO);										
									}
								}
								rowVO.setCellList(cellList);
								rowList.add(rowVO);
							}
							
						}
						sheetVO.setRowList(rowList);
						sheetMap.put(sheetName, sheetVO);
					}
				}
			}
		} catch (FileNotFoundException e) {
			logger.error(e);
			e.printStackTrace();
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		} finally {
			if (is != null) {
				try {
					is.close();
				} catch (IOException e) {}
			}
		}
		return sheetMap;
	}

	/**
	 * DESC: 解析OFFICE 2005及以上版本的EXCEL内容到对象中  
	 *       优先对excludeSheet的内容生效【exclude内容不为空时则include内容不做处理】
	 *       excludeSheet --- 排除的sheet, 在该内容中的Sheet不做解析
	 *       includeSheet --- 包含的sheet, 仅解析该内容中的Sheet内容  
	 * 
	 * @param fileName
	 * @param includeSheet
	 * @param excludeSheet
	 * @param matchType  -- 匹配模式<br>
	 * 						1 ： 等于<br>
	 * 						2： 包含<br>
	 * 						3：Start With<br>
	 * 						4 ： End With<br>
	 * @return
	 */
	private HashMap<String, HGExcelSheetVO> parseExcelData4XLSX(String fileName, List<String> includeSheet, List<String> excludeSheet, int matchType) {
		HashMap<String, HGExcelSheetVO> sheetMap = new HashMap<String, HGExcelSheetVO>();
		//read xls
		XSSFWorkbook workbook = null;
		File excelFile = new File(fileName);
		InputStream is = null;
		try {
			is = new FileInputStream(excelFile);
			workbook = new XSSFWorkbook(is);			
			for (int sheetIndexOf = 0; sheetIndexOf < workbook.getNumberOfSheets(); sheetIndexOf++) {
				XSSFSheet sheet = workbook.getSheetAt(sheetIndexOf);
				if (sheet != null) {
					String sheetName = sheet.getSheetName();
					HGExcelSheetVO sheetVO = new HGExcelSheetVO();
					sheetVO.setSheetName(sheetName);
					boolean flag = false;
					if (excludeSheet != null && excludeSheet.size() > 0) {
						if (!this.checkStrExist(sheetName, excludeSheet, matchType)) {
							flag = true;
						}
					} else if (includeSheet != null && includeSheet.size() > 0) {
						if (this.checkStrExist(sheetName, includeSheet, matchType)) {
							flag = true;
						}
					} else {
						flag = true;
					}
					//parse the row 
					if (flag) {
						int rowIndexOf = sheet.getFirstRowNum();
						int rowCounts = sheet.getLastRowNum();
						List<HGExcelRowVO> rowList = new ArrayList<HGExcelRowVO>();
						for (; rowIndexOf < rowCounts; rowIndexOf++) {
							XSSFRow row = sheet.getRow(rowIndexOf);
							if (row != null) {
								HGExcelRowVO rowVO = new HGExcelRowVO();
								rowVO.setRowIndex(rowIndexOf);
								List<HGExcelCellVO> cellList = new ArrayList<HGExcelCellVO>();
								int cellIndexOf = row.getFirstCellNum();
								int cellCounts = row.getLastCellNum();
								for (; cellIndexOf < cellCounts; cellIndexOf++) {
									XSSFCell cell = row.getCell(cellIndexOf);
									if (cell != null) {
										HGExcelCellVO cellVO = new HGExcelCellVO();
										cellVO.setCellIndex(cellIndexOf);
										cellVO.setRowIndex(rowIndexOf);
										cellVO.setCellType(cell.getCellType());
										switch (cell.getCellType()) {
											case Cell.CELL_TYPE_BLANK : cellVO.setCellValue("");break;
											case Cell.CELL_TYPE_BOOLEAN : cellVO.setCellValue(String.valueOf(cell.getBooleanCellValue()));break;
											case Cell.CELL_TYPE_NUMERIC : cellVO.setCellValue(String.valueOf(cell.getNumericCellValue()));break;
											case Cell.CELL_TYPE_STRING : cellVO.setCellValue(cell.getStringCellValue());break;
										}
										cellList.add(cellVO);										
									}
								}
								rowVO.setCellList(cellList);
								rowList.add(rowVO);
							}
							
						}
						sheetVO.setRowList(rowList);
						sheetMap.put(sheetName, sheetVO);
					}
				}
			}
		} catch (FileNotFoundException e) {
			logger.error(e);
			e.printStackTrace();
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		} finally {
			if (is != null) {
				try {
					is.close();
				} catch (IOException e) {}
			}
		}
		return sheetMap;
	}

	/**
	 * DESC: 修改OFFICE 2003版本的EXCEL单元格内容，将data中的数据内容写入到EXCEL对应的单元格中  <br>》
	 *       data -- key ： SHEET内容   <br>
	 *       VALUE： 待填写的CELL数据内容,HVExcelCellVO对象  <br> 
	 * 
	 * @param fileName
	 * @param data
	 * @return
	 */
	private boolean editExcelData4XLS(String fileName, HashMap<String, List<HGExcelCellVO>> data) {
		boolean result = false;
		//read xls 
		HSSFWorkbook workbook = null;
		File excelFile = new File(fileName);
		InputStream is = null;
		FileOutputStream fileOut = null;
		try {
			is = new FileInputStream(excelFile);
			workbook = new HSSFWorkbook(is);
			if (workbook != null && data != null && data.size() > 0) {
				for (String sheetName : data.keySet()) {
					//get the sheet 
					if (workbook.getSheet(sheetName) != null) {
						HSSFSheet sheet = workbook.getSheet(sheetName);
						List<HGExcelCellVO> cellList = data.get(sheetName);
						if (cellList != null && cellList.size() > 0) {
							for (HGExcelCellVO cellVO : cellList) {
								//get the row 
								if (sheet.getFirstRowNum() <= cellVO.getRowIndex() && cellVO.getRowIndex() <= sheet.getLastRowNum()) {
									HSSFRow row = sheet.getRow(cellVO.getRowIndex());
									//get the cell 
									if (row != null && row.getFirstCellNum() <= cellVO.getCellIndex() && cellVO.getCellIndex() <= row.getLastCellNum()) {
										HSSFCell cell = row.getCell(cellVO.getCellIndex());
										switch(cell.getCellType()) {
											case Cell.CELL_TYPE_BLANK : {
												cell.setCellType(Cell.CELL_TYPE_STRING);
												cell.setCellValue(cellVO.getCellValue()); break;
											}
											case Cell.CELL_TYPE_NUMERIC : {
												Double d = Double.valueOf(cellVO.getCellValue());
												cell.setCellValue(d);
												break;
											}
											case Cell.CELL_TYPE_STRING : {
												cell.setCellValue(cellVO.getCellValue());
												break;
											}
											default : {
												cell.setCellType(Cell.CELL_TYPE_STRING);
												cell.setCellValue(cellVO.getCellValue());
												break;
											}
										}
									}
								}
							}
						}
					}
				}
			}					
		} catch (FileNotFoundException e) {
			logger.error("FileNotFound Error : " + e);
			e.printStackTrace();
		} catch (IOException e) {
			logger.error("IOException Error : " + e);
			e.printStackTrace();
		} finally {			
			if (is != null) {
				try {
					is.close();
				} catch (IOException e) {}
			}
		}
		//write the file
		try {			
			if (workbook != null) {
				fileOut = new FileOutputStream(fileName);
				workbook.write(fileOut);	
				result = true;
			}	
		} catch (IOException e) {
			logger.error("IOException Error : " + e);
			e.printStackTrace();
		} finally {
			if (fileOut != null) {
				try {
					fileOut.close();
				} catch (IOException e) {}
			}
		}
		return result;
	}
	
	/**
	 * DESC: 修改OFFICE 2005+版本 EXCEL单元格内容，将data中的数据内容写入到EXECL 对应的单元个格  <br> 
	 *       data -- key ： SHEET内容   <br>
	 *       VALUE： 待填写的CELL数据内容,HVExcelCellVO对象  <br> 
	 * 
	 * @param fileName
	 * @param data
	 * @return
	 */
	private boolean editExcelData4XLSX(String fileName, HashMap<String, List<HGExcelCellVO>> data) {
		boolean result = false;
		//read xls 
		XSSFWorkbook workbook = null;
		File excelFile = new File(fileName);
		InputStream is = null;
		FileOutputStream fileOut = null;
		try {
			is = new FileInputStream(excelFile);
			workbook = new XSSFWorkbook(is);
			if (workbook != null && data != null && data.size() > 0) {
				for (String sheetName : data.keySet()) {
					//get the sheet 
					if (workbook.getSheet(sheetName) != null) {
						XSSFSheet sheet = workbook.getSheet(sheetName);
						List<HGExcelCellVO> cellList = data.get(sheetName);
						if (cellList != null && cellList.size() > 0) {
							for (HGExcelCellVO cellVO : cellList) {
								//get the row 
								if (sheet.getFirstRowNum() <= cellVO.getRowIndex() && cellVO.getRowIndex() <= sheet.getLastRowNum()) {
									XSSFRow row = sheet.getRow(cellVO.getRowIndex());
									//get the cell 
									if (row != null && row.getFirstCellNum() <= cellVO.getCellIndex() && cellVO.getCellIndex() <= row.getLastCellNum()) {
										XSSFCell cell = row.getCell(cellVO.getCellIndex());
										switch(cell.getCellType()) {
											case Cell.CELL_TYPE_BLANK : {
												cell.setCellValue(cellVO.getCellValue()); break;
											}
											case Cell.CELL_TYPE_NUMERIC : {
												Double d = Double.valueOf(cellVO.getCellValue());
												cell.setCellValue(d);
												break;
											}
											case Cell.CELL_TYPE_STRING : {
												cell.setCellValue(cellVO.getCellValue());
												break;
											}
										}
										if (cell.getCellType() == Cell.CELL_TYPE_STRING) {
											cell.setCellValue(cellVO.getCellValue());
										}
									}
								}
							}
						}
					}
				}
			}
			//write the file
			if (workbook != null) {
				fileOut = new FileOutputStream(fileName);
				workbook.write(fileOut);	
				result = true;
			}			
		} catch (FileNotFoundException e) {
			logger.error("FileNotFound Error : " + e);
			e.printStackTrace();
		} catch (IOException e) {
			logger.error("IOException Error : " + e);
			e.printStackTrace();
		} finally {
			if (fileOut != null) {
				try {
					fileOut.close();
				} catch (IOException e) {}
			}
			if (is != null) {
				try {
					is.close();
				} catch (IOException e) {}
			}
		}
		return result;
	}
	
	/**
	 * DESC:  检测匹配的字符串是否在列表中体现 <br>
	 * 
	 * @param str
	 * @param strList
	 * @param matchType  -- 匹配模式<br>
	 * 						1 ： 等于<br>
	 * 						2： 包含<br>
	 * 						3：Start With<br>
	 * 						4 ： End With<br>
	 * @return	
	 */
	private boolean checkStrExist(String str, List<String> strList, int matchType) {
		boolean result = false;
		if (str != null && strList != null && strList.size() > 0) {
			for (String s : strList) {
				switch (matchType) {
					case 1 : {
						result = s.equals(str) ? true : false;
						break; //等于
					}
					case 2 : {
						result = s.indexOf(str) > -1 ? true : false;
						break; //包含
					} 
					case 3 : {
						result = s.startsWith(str) ? true : false;
						break; //start with
					}
					case 4 : {
						result = s.endsWith(str) ? true : false;
						break; //end with
					}
				}
				if (result) {
					break;
				}
			}
		}
		return result;
	}	
}
